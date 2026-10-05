"""Reference calculation of the development model in plain Python.

This is the specification. The workbook formulas written by the builder must
reproduce these numbers; tests compare the two after LibreOffice recalculates
the workbook. Keys match the builder's row keys ("S01.cost.total" and so on).

Conventions
- Periods t = 1..N (index t-1 in arrays). L is the last actual month: months
  1..L are actuals (zero in this version until the Home Hub feed exists) and
  opening balances apply at the end of month L.
- Amounts exclude GST unless the key says otherwise.
- Equity is the plug: positive = contribution, negative = distribution. Cash
  therefore stays at nil in every month.
"""

from __future__ import annotations

import numpy as np
from scipy.stats import beta as beta_dist

from .recipe import Recipe, Site


def irr(flows: np.ndarray, guess: float = 0.01) -> float:
    """Monthly IRR: the NPV root nearest to `guess` (Excel's IRR starts its
    Newton search at the guess; the model passes 0.01 a month). Flows after a sale can change sign twice (GST paid
    the month after settlement), so a single bracket is not enough."""
    f = np.asarray(flows, dtype=float)
    if not (np.any(f < -1e-9) and np.any(f > 1e-9)):
        return float("nan")
    t = np.arange(len(f))

    def npv(r):
        return float(np.sum(f / (1 + r) ** t))

    grid = np.linspace(-0.9, 1.0, 1901)
    with np.errstate(all="ignore"):
        vals = np.array([npv(x) for x in grid])
    roots = []
    for i in range(len(grid) - 1):
        a, b = vals[i], vals[i + 1]
        if np.isfinite(a) and np.isfinite(b) and a * b <= 0:
            lo, hi = grid[i], grid[i + 1]
            for _ in range(100):
                mid = (lo + hi) / 2
                if npv(lo) * npv(mid) <= 0:
                    hi = mid
                else:
                    lo = mid
            roots.append((lo + hi) / 2)
    if not roots:
        return float("nan")
    return min(roots, key=lambda x: abs(x - guess))


def shares(profile: str, start: int, months: int, alpha: float, beta: float, n: int) -> np.ndarray:
    out = np.zeros(n)
    for t in range(1, n + 1):
        k = t - start + 1
        if profile == "Lump sum":
            out[t - 1] = 1.0 if k == 1 else 0.0
        elif 1 <= k <= months:
            if profile == "Flat":
                out[t - 1] = 1.0 / months
            else:
                out[t - 1] = beta_dist.cdf(k / months, alpha, beta) - beta_dist.cdf((k - 1) / months, alpha, beta)
    return out


def site_calc(r: Recipe, s: Site) -> dict[str, object]:
    n, L, g = r.periods, r.last_actual, r.gst_rate
    t = np.arange(1, n + 1)
    fc = (t > L).astype(float)
    reg = 1.0 if s.gst_registered == "Yes" else 0.0
    out: dict[str, object] = {}
    k = s.code

    # ---- derived scalars
    total_budget = sum(x.budget for x in s.lines)
    sale_idx = r.month_index(s.sale_date)
    debt_idx = r.month_index(s.debt_start)
    standard = s.sale_gst == "Standard-rated"
    revenue_amt = s.sale_price / (1 + g) if (reg and standard) else s.sale_price
    out_gst_amt = s.sale_price - revenue_amt
    if s.limit_basis == "Override":
        limit = s.limit_override
    elif s.limit_basis == "LTC on budget":
        limit = s.max_ltc * total_budget
    elif s.limit_basis == "LVR on As Complete":
        limit = s.max_lvr * s.sale_price
    else:
        limit = s.max_lvr * s.btr_value
    opening_equity = s.opening_wip - s.opening_debt
    out.update({f"{k}.d.total_budget": total_budget, f"{k}.d.sale_idx": sale_idx, f"{k}.d.debt_idx": debt_idx,
                f"{k}.d.revenue": revenue_amt, f"{k}.d.output_gst": out_gst_amt, f"{k}.d.limit": limit,
                f"{k}.d.opening_equity": opening_equity})

    # ---- costs
    phased_fc = np.zeros(n)
    line_phased = {}
    for x in s.lines:
        sh = shares(x.profile, r.month_index(x.start), x.months, s.alpha, s.beta, n)
        out[f"{k}.share.{x.key}"] = sh
        line_phased[x.key] = x.budget * sh
        phased_fc += x.budget * sh * fc
    denom = phased_fc.sum()
    scale = (total_budget - s.opening_wip) / denom if abs(denom) > 1e-9 else 0.0
    out[f"{k}.cost.scale"] = scale
    cost = np.zeros(n)
    in_gst = np.zeros(n)
    for x in s.lines:
        c = line_phased[x.key] * fc * scale
        out[f"{k}.cost.{x.key}"] = c
        cost += c
        if x.gst == "Yes":
            in_gst += c * g
    unrec = in_gst * (1 - reg)
    cost_paid = cost + in_gst
    cum_cost = np.where(fc > 0, s.opening_wip + np.cumsum((cost + unrec) * fc), 0.0)
    out.update({f"{k}.cost.total": cost, f"{k}.cost.input_gst": in_gst, f"{k}.cost.unrec_gst": unrec,
                f"{k}.cost.paid": cost_paid, f"{k}.cost.cum": cum_cost})

    # ---- sales
    sale = ((t == sale_idx) & (t > L)).astype(float)
    sold = (t >= sale_idx).astype(float)
    revenue = revenue_amt * sale
    out_gst = out_gst_amt * sale
    proceeds = s.sale_price * sale
    via = 1.0 if s.via_spv == "Yes" else 0.0
    out.update({f"{k}.sale.flag": sale, f"{k}.sale.sold": sold, f"{k}.sale.revenue": revenue,
                f"{k}.sale.output_gst": out_gst, f"{k}.sale.proceeds": proceeds,
                f"{k}.exit.spv_price": s.sale_price * via * sale,
                f"{k}.exit.fund_price": s.btr_value * via * sale,
                f"{k}.exit.hhlp_margin": (s.btr_value - s.sale_price) * via * sale,
                f"{k}.exit.gst_cost": out_gst_amt * via * sale,
                f"{k}.exit.fund_debt": r.fund_lvr * s.btr_value * via * sale,
                f"{k}.exit.fund_equity": (1 - r.fund_lvr) * s.btr_value * via * sale})

    # ---- GST
    net_gst = (out_gst - in_gst) * reg
    gst_cash = np.zeros(n)
    for i in range(n):
        src = (i + 1) - r.gst_lag
        if src > L:
            gst_cash[i] = net_gst[src - 1]
    gst_bal = np.zeros(n)
    prev = 0.0
    for i in range(n):
        if i + 1 <= L:
            gst_bal[i] = 0.0
            continue
        prev = prev + net_gst[i] - gst_cash[i]
        gst_bal[i] = prev
    out.update({f"{k}.gst.net": net_gst, f"{k}.gst.cash": gst_cash, f"{k}.gst.balance": gst_bal})

    # ---- funding
    cap = s.capitalise == "Yes"
    release = s.release_equity == "Yes"
    keys = ["ceiling", "open", "interest", "line_fee", "estab_fee", "fin_cost", "headroom", "fin_cap", "fin_cash",
            "need", "gap", "draw", "close", "cash_open", "cash", "equity", "contrib", "distrib", "eq_bal", "ltc", "lvr",
            "irr_eq", "irr_proj"]
    F = {x: np.zeros(n) for x in keys}
    d_prev = 0.0
    eq_bal = 0.0
    for i in range(n):
        ti = i + 1
        if ti <= L:
            continue
        cc = cum_cost[i]
        if sold[i]:
            ceil = 0.0
        elif ti >= debt_idx:
            ceil = min(limit, max(0.0, cc - s.equity_commitment) if s.method == "Equity first" else s.max_ltc * cc)
        else:
            ceil = s.land_loan
        d0 = s.opening_debt if ti == L + 1 else d_prev
        interest = d0 * s.rate / 12
        lf = limit * s.line_fee / 12 if (ti >= debt_idx and not sold[i]) else 0.0
        ef = limit * s.estab_fee if ti == debt_idx else 0.0
        fin = interest + lf + ef
        head = max(0.0, ceil - d0)
        fcap = min(fin, head) if cap else 0.0
        fcash = fin - fcap
        need = cost_paid[i] + gst_cash[i] - proceeds[i]
        gap = ceil - d0 - fcap
        if gap < 0 or release:
            draw = gap
        else:
            draw = min(gap, max(0.0, need + fcash))
        d = d0 + fcap + draw
        # cash is held back to meet GST payable, so output GST is not
        # distributed and then called back the next month
        cash_open = 0.0 if ti == L + 1 else F["cash"][i - 1]
        cash = max(0.0, gst_bal[i])
        eq = need + fcash - draw + (cash - cash_open)
        eq_bal = (opening_equity if ti == L + 1 else eq_bal) + eq
        vals = dict(ceiling=ceil, open=d0, interest=interest, line_fee=lf, estab_fee=ef, fin_cost=fin, headroom=head,
                    fin_cap=fcap, fin_cash=fcash, need=need, gap=gap, draw=draw, close=d, cash_open=cash_open,
                    cash=cash, equity=eq,
                    contrib=max(0.0, eq), distrib=max(0.0, -eq), eq_bal=eq_bal,
                    ltc=d / cc if cc > 0 else 0.0, lvr=d / s.sale_price if s.sale_price > 0 else 0.0,
                    irr_eq=-eq - (opening_equity if ti == L + 1 else 0.0),
                    irr_proj=-need - (cash - cash_open) - (s.opening_wip if ti == L + 1 else 0.0))
        for key, v in vals.items():
            F[key][i] = v
        d_prev = d
    for key, v in F.items():
        out[f"{k}.fund.{key}"] = v

    # ---- statements
    cap_acc = s.cap_borrowing == "Yes"
    fin_wip = F["fin_cost"] if cap_acc else np.zeros(n)
    fin_exp = F["fin_cost"] - fin_wip
    wip_open = np.zeros(n)
    adds = (cost + unrec) * fc + fin_wip
    cos = np.zeros(n)
    wip = np.zeros(n)
    prev = 0.0
    for i in range(n):
        ti = i + 1
        if ti <= L:
            continue
        wo = s.opening_wip if ti == L + 1 else prev
        wip_open[i] = wo
        cos[i] = (wo + adds[i]) if sold[i] else 0.0
        wip[i] = wo + adds[i] - cos[i]
        prev = wip[i]
    profit = revenue - cos - fin_exp
    re = np.where(fc > 0, np.cumsum(profit * fc), 0.0)
    eq_net = np.where(fc > 0, F["eq_bal"], 0.0)
    bs_check = wip + F["cash"] - (F["close"] + gst_bal + eq_net + re)
    out.update({f"{k}.fs.wip_open": wip_open, f"{k}.fs.wip_adds": adds, f"{k}.fs.cos": cos, f"{k}.fs.wip": wip,
                f"{k}.fs.fin_wip": fin_wip, f"{k}.fs.fin_exp": fin_exp, f"{k}.fs.revenue": revenue,
                f"{k}.fs.profit": profit, f"{k}.fs.re": re, f"{k}.fs.bs_check": bs_check})

    # ---- returns
    eq_irr_m = irr(F["irr_eq"])
    proj_irr_m = irr(F["irr_proj"])
    contrib_total = F["contrib"].sum() + max(0.0, opening_equity)
    out.update({
        f"{k}.ret.revenue": revenue.sum(),
        f"{k}.ret.dev_cost": s.opening_wip + ((cost + unrec) * fc).sum(),
        f"{k}.ret.finance": F["fin_cost"].sum(),
        f"{k}.ret.profit": profit.sum(),
        f"{k}.ret.margin": profit.sum() / revenue.sum() if revenue.sum() else 0.0,
        f"{k}.ret.poc": profit.sum() / (cos.sum() + fin_exp.sum()) if (cos.sum() + fin_exp.sum()) else 0.0,
        f"{k}.ret.peak_debt": F["close"].max(),
        f"{k}.ret.peak_equity": F["eq_bal"].max(),
        f"{k}.ret.eq_irr": (1 + eq_irr_m) ** 12 - 1 if eq_irr_m == eq_irr_m else 0.0,
        f"{k}.ret.proj_irr": (1 + proj_irr_m) ** 12 - 1 if proj_irr_m == proj_irr_m else 0.0,
        f"{k}.ret.multiple": F["distrib"].sum() / contrib_total if contrib_total else 0.0,
        f"{k}.ret.hhlp_margin": (s.btr_value - s.sale_price) * via if sale.sum() else 0.0,
        f"{k}.ret.group_margin": profit.sum() + ((s.btr_value - s.sale_price) * via if sale.sum() else 0.0),
    })
    return out


def calculate(r: Recipe) -> dict[str, object]:
    res: dict[str, object] = {}
    for s in r.sites:
        res.update(site_calc(r, s))
    return res
