"""Reference calculation of the full financial model (library/hfg), in plain Python.

This is the specification the workbook must reproduce. It runs every module of a recipe month by
month exactly as the library's formulas do, in the engine's block order, so a recipe built by the
engine and recalculated by LibreOffice must give the same numbers.

The same function generates the fictional history: run with no actual months, the true historical
drivers and the opening balances, it gives the income statement and balance sheet the recipe then
types into the historical statements. Run again with the actual months set, the model reads that
history in actual months and forecasts the rest.

Conventions
- t = 0..T-1 is period t+1. prev(x, t) is x[t-1], or the opening balance in the first month for a
  balance with a historical line (0 otherwise), as the engine renders [x@prev].
- Every amount is in dollars, excluding GST unless the row says otherwise.
- `init` supplies states a model cannot know (dues and accumulators from before the first month);
  only the history generator passes it. The model always starts them at 0.
"""

from __future__ import annotations

from dataclasses import dataclass, field

AREAS = ["Income summary", "Balance summary", "Cash summary", "Budget summary", "Scenario summary", "Scenarios",
         "Historical IS", "Historical BS", "Revenue and expenses", "Working capital", "Assets", "Capital", "Tax",
         "Other items", "Financials", "Budget", "Checks"]

MODULE_AREA = {
    "fm.revenue": "Revenue and expenses", "fm.cogs": "Revenue and expenses", "fm.staff": "Revenue and expenses",
    "fm.opex": "Revenue and expenses", "fm.other_income": "Revenue and expenses", "fm.other_expense": "Revenue and expenses",
    "fm.collection": "Working capital", "fm.payment": "Working capital", "fm.inventory": "Working capital",
    "fm.payroll": "Working capital", "fm.fixed_asset": "Assets", "fm.intangible": "Assets", "fm.capex": "Assets",
    "fm.debt": "Capital", "fm.equity": "Capital", "fm.gst": "Tax", "fm.income_tax": "Tax",
    "fm.interest_cash": "Other items", "fm.other_ca": "Other items", "fm.other_nca": "Other items",
    "fm.other_cl": "Other items", "fm.other_ncl": "Other items", "fm.other_equity": "Other items",
    "fm.statements": "Financials", "fm.checks": "Checks",
}

GST_RATE_DEFAULT = 0.15


@dataclass
class Inst:
    module: str
    number: int
    settings: dict
    name: str | None = None
    data: dict = field(default_factory=dict)

    @property
    def uid(self) -> str:
        return f"{self.module}#{self.number}"


def block_order(instances: list[Inst]) -> list[Inst]:
    """The engine's block order: by area, then the order instances were inserted."""
    order = {i.uid: k for k, i in enumerate(instances)}
    return sorted(instances, key=lambda i: (AREAS.index(MODULE_AREA[i.module]), order[i.uid]))


def series(inst: Inst, key: str, T: int, default: float | None = None) -> list[float]:
    v = inst.data.get("series", {}).get(key, default)
    if isinstance(v, list):
        return [float(x) if x is not None else 0.0 for x in (v + [None] * T)[:T]]
    return [float(v) if v is not None else 0.0] * T


def hist(inst: Inst, key: str, T: int) -> list[float]:
    v = inst.data.get("history", {}).get(key, [])
    return [float(x) if x is not None else 0.0 for x in (list(v) + [None] * T)[:T]]


def opening(inst: Inst, key: str) -> float:
    return float(inst.data.get("opening", {}).get(key, 0.0))


def scen(inst: Inst, key: str, active: int) -> float:
    v = inst.data.get("scenarios", {}).get(key, [])
    return float(v[active - 1]) if len(v) >= active else 0.0


def run(instances: list[Inst], T: int, last_actual: int, fy_end: int, start_month: int, start_year: int,
        scenario: int = 1, init: dict | None = None) -> dict[str, list[float]]:
    """Every row of every block, by "uid/key" (and the statements' totals by "fs/key")."""
    init = init or {}
    blocks = block_order(instances)
    out: dict[str, list[float]] = {}
    actual = [t + 1 <= last_actual for t in range(T)]
    # timeline: calendar month and year of each period's end, month of the financial year
    cal_month = [((start_month - 1 + t) % 12) + 1 for t in range(T)]
    fy_month = [((cm - fy_end - 1) % 12) + 1 for cm in cal_month]

    def setting(i: Inst, key: str, default):
        return i.settings.get(key, default)

    gst_rate = next((setting(i, "rate", GST_RATE_DEFAULT) for i in instances if i.module == "fm.gst"), GST_RATE_DEFAULT)

    def gst_of(i: Inst) -> float:
        return gst_rate if setting(i, "gst", 1) == 1 else 0.0

    def prev(x: list[float], t: int, op: float = 0.0, n: int = 1) -> float:
        return x[t - n] if t - n >= 0 else (op if n == 1 else 0.0)

    def lag(x: list[float], t: int, n: int) -> float:
        return x[t - n] if t - n >= 0 else 0.0

    # list positions
    def positions(mods: tuple[str, ...], first: int) -> dict[str, int]:
        return {i.uid: first + k + 1 for k, i in enumerate(b for b in blocks if b.module in mods)}

    pos_rev = positions(("fm.revenue",), 0)
    pos_coll = positions(("fm.collection",), 1)
    pos_pay = positions(("fm.payment",), 1)
    pos_stock = positions(("fm.inventory",), 1)
    pos_class = positions(("fm.fixed_asset", "fm.intangible"), 0)

    links: dict[str, list[tuple[Inst, list[float]]]] = {}

    def send(link: str, i: Inst, arr: list[float]) -> None:
        links.setdefault(link, []).append((i, arr))

    def total(link: str) -> list[float]:
        return [sum(a[t] for _, a in links.get(link, [])) for t in range(T)]

    # The statements' links (fs.cash, fs.npbt, eq.npat) feed modules earlier in block order, which read
    # them one month back (cash) or in the same month (profit). So the whole model is worked out one
    # month at a time, every block in order, as Excel's dependency order does.
    rows: dict[str, list[float]] = {}

    def R(uid: str, key: str) -> list[float]:
        k = f"{uid}/{key}"
        if k not in rows:
            rows[k] = [0.0] * T
        return rows[k]

    def P(uid: str, key: str, t: int, op_key: str | None = None, n: int = 1, inst: Inst | None = None) -> float:
        x = R(uid, key)
        if t - n >= 0:
            return x[t - n]
        if n == 1 and op_key is not None and inst is not None:
            return opening(inst, op_key)
        return 0.0

    by_mod = lambda m: [b for b in blocks if b.module == m]
    revs, cogs, staff, opex = by_mod("fm.revenue"), by_mod("fm.cogs"), by_mod("fm.staff"), by_mod("fm.opex")
    oinc, oexp = by_mod("fm.other_income"), by_mod("fm.other_expense")
    colls, pays, stocks = by_mod("fm.collection"), by_mod("fm.payment"), by_mod("fm.inventory")
    fas, ints, capexes = by_mod("fm.fixed_asset"), by_mod("fm.intangible"), by_mod("fm.capex")
    debts = by_mod("fm.debt")
    others = {m: by_mod(m) for m in ("fm.other_ca", "fm.other_nca", "fm.other_cl", "fm.other_ncl", "fm.other_equity")}
    single = lambda m: next(iter(by_mod(m)), None)
    payroll, equity, gst, itax, icash, fs = (single("fm.payroll"), single("fm.equity"), single("fm.gst"),
                                             single("fm.income_tax"), single("fm.interest_cash"), single("fm.statements"))

    for t in range(T):
        a = actual[t]
        p = t + 1
        # ---- revenue lines
        for i in revs:
            u = i.uid
            m = setting(i, "method", 1)
            amount, price, volume, growth = (series(i, k, T)[t] for k in ("amount", "price", "volume", "growth"))
            pre = [amount, price * volume, lag(R(u, "revenue"), t, 12) * (1 + growth)][m - 1]
            R(u, "pre")[t] = pre
            R(u, "revenue")[t] = hist(i, "revenue", T)[t] if a else pre * (1 + scen(i, "revenue", scenario))
            R(u, "gst_charged")[t] = R(u, "revenue")[t] * gst_of(i)
            R(u, "invoiced")[t] = R(u, "revenue")[t] + R(u, "gst_charged")[t]
            R(u, "cash_sales")[t] = R(u, "invoiced")[t] if setting(i, "collect", 1) == 1 else 0.0
        # ---- cost of sales lines
        rev_lines = [R(i.uid, "revenue") for i in revs]
        total_rev = sum(x[t] for x in rev_lines)
        for i in cogs:
            u = i.uid
            m = setting(i, "method", 1)
            share, amount, growth = (series(i, k, T)[t] for k in ("share", "amount", "growth"))
            base = rev_lines[setting(i, "source", 1) - 1][t] if rev_lines else 0.0
            pre = [base * share, amount, lag(R(u, "cost"), t, 12) * (1 + growth)][m - 1]
            R(u, "pre")[t] = pre
            R(u, "cost")[t] = hist(i, "cost", T)[t] if a else pre * (1 + scen(i, "cost", scenario))
            stocked = setting(i, "stock", 1) > 1
            R(u, "stocked")[t] = R(u, "cost")[t] if stocked else 0.0
            R(u, "bought")[t] = 0.0 if stocked else R(u, "cost")[t]
            R(u, "gst_paid")[t] = R(u, "bought")[t] * gst_of(i)
            R(u, "billed")[t] = R(u, "bought")[t] + R(u, "gst_paid")[t]
            R(u, "cash_paid")[t] = R(u, "billed")[t] if setting(i, "pay", 1) == 1 else 0.0
        # ---- staff
        for i in staff:
            u = i.uid
            fte, salary = series(i, "fte", T)[t], series(i, "salary", T)[t]
            R(u, "salaries")[t] = hist(i, "salaries", T)[t] if a else fte * salary / 12 * (1 + scen(i, "salaries", scenario))
            R(u, "ks")[t] = hist(i, "ks", T)[t] if a else R(u, "salaries")[t] * setting(i, "kiwisaver", 0.03)
        # ---- operating expenses
        for i in opex:
            u = i.uid
            m = setting(i, "method", 1)
            amount, growth, share = (series(i, k, T)[t] for k in ("amount", "growth", "share"))
            pre = [amount, lag(R(u, "cost"), t, 12) * (1 + growth), total_rev * share][m - 1]
            R(u, "pre")[t] = pre
            R(u, "cost")[t] = hist(i, "cost", T)[t] if a else pre * (1 + scen(i, "cost", scenario))
            R(u, "gst_paid")[t] = R(u, "cost")[t] * gst_of(i)
            R(u, "billed")[t] = R(u, "cost")[t] + R(u, "gst_paid")[t]
            R(u, "cash_paid")[t] = R(u, "billed")[t] if setting(i, "pay", 1) == 1 else 0.0
        for i in oinc:
            R(i.uid, "income")[t] = hist(i, "income", T)[t] if a else series(i, "amount", T)[t]
        for i in oexp:
            R(i.uid, "expense")[t] = hist(i, "expense", T)[t] if a else series(i, "amount", T)[t]
            R(i.uid, "cash")[t] = -R(i.uid, "expense")[t]

        # ---- collection profiles
        for i in colls:
            u = i.uid
            days = setting(i, "days", 30)
            sales = sum(R(r.uid, "invoiced")[t] for r in revs if setting(r, "collect", 1) == pos_coll[u])
            R(u, "sales")[t] = sales
            op = P(u, "closing", t, "closing", inst=i)
            R(u, "opening")[t] = op
            target = (sales * min(days, 30) + P(u, "sales", t) * min(max(days - 30, 0), 30)
                      + P(u, "sales", t, n=2) * min(max(days - 60, 0), 30)) / 30
            R(u, "receipts")[t] = op + sales - (hist(i, "closing", T)[t] if a else target)
            R(u, "closing")[t] = hist(i, "closing", T)[t] if a else op + sales - R(u, "receipts")[t]
        # ---- stock profiles
        for i in stocks:
            u = i.uid
            used = sum(R(c.uid, "stocked")[t] for c in cogs if setting(c, "stock", 1) == pos_stock[u])
            R(u, "used")[t] = used
            op = P(u, "closing", t, "closing", inst=i)
            R(u, "opening")[t] = op
            R(u, "purchases")[t] = (hist(i, "closing", T)[t] - op + used) if a else used * setting(i, "days", 30) / 30 - op + used
            R(u, "closing")[t] = hist(i, "closing", T)[t] if a else op + R(u, "purchases")[t] - used
            R(u, "gst_paid")[t] = R(u, "purchases")[t] * gst_of(i)
            R(u, "billed")[t] = R(u, "purchases")[t] + R(u, "gst_paid")[t]
            R(u, "cash_paid")[t] = R(u, "billed")[t] if setting(i, "pay", 1) == 1 else 0.0
        # ---- payment profiles (bills from cost of sales, operating expenses and stock)
        for i in pays:
            u = i.uid
            days = setting(i, "days", 30)
            billed = sum(R(x.uid, "billed")[t] for x in cogs + opex + stocks if setting(x, "pay", 1) == pos_pay[u])
            R(u, "billed")[t] = billed
            op = P(u, "closing", t, "closing", inst=i)
            R(u, "opening")[t] = op
            target = (billed * min(days, 30) + P(u, "billed", t) * min(max(days - 30, 0), 30)
                      + P(u, "billed", t, n=2) * min(max(days - 60, 0), 30)) / 30
            R(u, "payments")[t] = op + billed - (hist(i, "closing", T)[t] if a else target)
            R(u, "closing")[t] = hist(i, "closing", T)[t] if a else op + billed - R(u, "payments")[t]
        # ---- payroll liabilities
        sal = sum(R(i.uid, "salaries")[t] for i in staff)
        ks = sum(R(i.uid, "ks")[t] for i in staff)
        if payroll:
            u, i = payroll.uid, payroll
            R(u, "owed")[t] = sal * setting(i, "withheld", 0.22) + ks
            R(u, "payable")[t] = hist(i, "payable", T)[t] if a else R(u, "owed")[t]
            R(u, "leave")[t] = hist(i, "leave", T)[t] if a else sal * setting(i, "leave", 1.2)
            R(u, "movement")[t] = hist(i, "movement", T)[t] if a else R(u, "leave")[t] - P(u, "leave", t, "leave", inst=i)
            R(u, "paid")[t] = sal + ks + P(u, "payable", t, "payable", inst=i) - R(u, "payable")[t]

        # ---- assets: capital expenditure lines, then each class
        for i in capexes:
            u = i.uid
            R(u, "amount")[t] = series(i, "amount", T)[t]
            R(u, "gst_paid")[t] = R(u, "amount")[t] * gst_of(i)
            R(u, "paid")[t] = R(u, "amount")[t] + R(u, "gst_paid")[t]
        for i, flow in [(x, "dep") for x in fas] + [(x, "amort") for x in ints]:
            u = i.uid
            add = sum(R(c.uid, "amount")[t] for c in capexes if setting(c, "class", 1) == pos_class[u])
            R(u, "additions")[t] = add
            op = P(u, "nbv", t, "nbv", inst=i)
            R(u, "opening")[t] = op
            R(u, flow)[t] = hist(i, flow, T)[t] if a else (op + add) * setting(i, "rate", 0.2) / 12
            R(u, "nbv")[t] = hist(i, "nbv", T)[t] if a else op + add - R(u, flow)[t]

        # ---- debt
        for i in debts:
            u = i.uid
            op = P(u, "balance", t, "balance", inst=i)
            R(u, "opening")[t] = op
            hb = hist(i, "balance", T)[t]
            R(u, "drawn")[t] = max(0.0, hb - op) if a else series(i, "draw", T)[t]
            R(u, "repaid")[t] = max(0.0, op - hb) if a else min(series(i, "repay", T)[t], op + R(u, "drawn")[t])
            R(u, "balance")[t] = hb if a else op + R(u, "drawn")[t] - R(u, "repaid")[t]
            R(u, "rate_now")[t] = setting(i, "rate", 0.07) * (1 + scen(i, "rate_now", scenario))
            R(u, "interest")[t] = hist(i, "interest", T)[t] if a else op * R(u, "rate_now")[t] / 12

        # ---- other balance sheet items
        signs = {"fm.other_ca": -1, "fm.other_nca": -1, "fm.other_cl": 1, "fm.other_ncl": 1, "fm.other_equity": 1}
        for mod, items in others.items():
            for i in items:
                u = i.uid
                pb = P(u, "balance", t, "balance", inst=i)
                R(u, "balance")[t] = hist(i, "balance", T)[t] if a else pb + series(i, "change", T)[t]
                R(u, "cash")[t] = signs[mod] * (R(u, "balance")[t] - pb)

        # ---- interest on cash (reads last month's closing cash from the statements)
        if icash:
            u, i = icash.uid, icash
            R(u, "opening")[t] = R("fs", "cash")[t - 1] if t > 0 else 0.0   # [cash@prev] reads the collected row: 0 before the first month
            o = R(u, "opening")[t]
            R(u, "interest")[t] = hist(i, "interest", T)[t] if a else (o * setting(i, "rate", 0.025) if o >= 0 else o * setting(i, "od_rate", 0.095)) / 12

        # ---- GST
        out_gst = sum(R(i.uid, "gst_charged")[t] for i in revs)
        in_gst = sum(R(i.uid, "gst_paid")[t] for i in cogs + opex + stocks + capexes)
        if gst:
            u, i = gst.uid, gst
            R(u, "net")[t] = out_gst - in_gst
            n = [1, 2, 6][setting(i, "cycle", 2) - 1]
            R(u, "period_end")[t] = 1.0 if cal_month[t] % n == 0 else 0.0
            pe_prev = P(u, "period_end", t) if t > 0 else init.get("gst_period_end_prev", 0.0)
            acc_prev = P(u, "accrued", t) if t > 0 else init.get("gst_accrued_prev", 0.0)
            R(u, "accrued")[t] = (0.0 if pe_prev == 1 else acc_prev) + R(u, "net")[t]
            R(u, "due")[t] = R(u, "accrued")[t] if R(u, "period_end")[t] == 1 else 0.0
            due_prev = P(u, "due", t) if t > 0 else init.get("gst_due_prev", 0.0)
            pay_prev = P(u, "payable", t, "payable", inst=i)
            R(u, "paid")[t] = (pay_prev + R(u, "net")[t] - hist(i, "payable", T)[t]) if a else due_prev
            R(u, "payable")[t] = hist(i, "payable", T)[t] if a else pay_prev + R(u, "net")[t] - R(u, "paid")[t]

        # ---- statements: income statement
        def ssum(mods, key):
            return sum(R(x.uid, key)[t] for x in mods)
        S = lambda k: R("fs", k)
        S("revenue")[t] = total_rev
        S("cogs")[t] = ssum(cogs, "cost")
        S("gross")[t] = S("revenue")[t] - S("cogs")[t]
        S("other_income")[t] = ssum(oinc, "income")
        S("staff")[t] = sal + ks + (R(payroll.uid, "movement")[t] if payroll else 0.0)
        S("opex")[t] = ssum(opex, "cost")
        S("other_expense")[t] = ssum(oexp, "expense")
        S("ebitda")[t] = S("gross")[t] + S("other_income")[t] - S("staff")[t] - S("opex")[t] - S("other_expense")[t]
        S("da")[t] = ssum(fas, "dep") + ssum(ints, "amort")
        S("ebit")[t] = S("ebitda")[t] - S("da")[t]
        S("interest_income")[t] = R(icash.uid, "interest")[t] if icash else 0.0
        S("interest_expense")[t] = ssum(debts, "interest")
        S("interest")[t] = S("interest_expense")[t] - S("interest_income")[t]
        S("npbt")[t] = S("ebit")[t] - S("interest")[t]

        # ---- income tax (reads this month's profit before tax)
        if itax:
            u, i = itax.uid, itax
            R(u, "expense")[t] = hist(i, "expense", T)[t] if a else S("npbt")[t] * setting(i, "rate", 0.28)
            fm1 = fy_month[t] == 1
            ytd_prev = P(u, "ytd", t) if t > 0 else init.get("tax_ytd_prev", 0.0)
            ly_prev = P(u, "last_year", t) if t > 0 else init.get("tax_last_year_prev", 0.0)
            ty_prev = P(u, "two_years", t) if t > 0 else init.get("tax_two_years_prev", 0.0)
            pfy_prev = P(u, "paid_fy", t) if t > 0 else init.get("tax_paid_fy_prev", 0.0)
            R(u, "ytd")[t] = (0.0 if fm1 else ytd_prev) + R(u, "expense")[t]
            R(u, "last_year")[t] = ytd_prev if fm1 else ly_prev
            R(u, "two_years")[t] = ly_prev if fm1 else ty_prev
            up = setting(i, "uplift", 0.05)
            pay_prev = P(u, "payable", t, "payable", inst=i)
            fm = fy_month[t]
            if a:
                paid = pay_prev + R(u, "expense")[t] - hist(i, "payable", T)[t]
            elif fm in (5, 10):
                paid = R(u, "last_year")[t] * (1 + up) / 3
            elif fm == 2:
                paid = R(u, "two_years")[t] * (1 + up) / 3
            elif fm == 11:
                paid = pay_prev - (ytd_prev - pfy_prev)
            else:
                paid = 0.0
            R(u, "paid")[t] = paid
            R(u, "paid_fy")[t] = (0.0 if fm1 else pfy_prev) + (paid if fm in (5, 10) else 0.0)
            R(u, "payable")[t] = hist(i, "payable", T)[t] if a else pay_prev + R(u, "expense")[t] - paid
        S("tax")[t] = R(itax.uid, "expense")[t] if itax else 0.0
        S("npat")[t] = S("npbt")[t] - S("tax")[t]

        # ---- equity and dividends (profit this month, cash last month)
        if equity:
            u, i = equity.uid, equity
            cap_prev = P(u, "capital", t, "capital", inst=i)
            R(u, "capital")[t] = hist(i, "capital", T)[t] if a else cap_prev + series(i, "issue", T)[t]
            R(u, "issued")[t] = R(u, "capital")[t] - cap_prev
            ytd_prev = P(u, "ytd", t) if t > 0 else init.get("npat_ytd_prev", 0.0)
            R(u, "ytd")[t] = (0.0 if fy_month[t] == 1 else ytd_prev) + S("npat")[t]
            R(u, "proposed")[t] = max(0.0, R(u, "ytd")[t] * setting(i, "payout", 0.5)) if fy_month[t] == 12 else 0.0
            re_prev = P(u, "retained", t, "retained", inst=i)
            cash_prev = R("fs", "cash")[t - 1] if t > 0 else 0.0
            if a:
                declared = re_prev + S("npat")[t] - hist(i, "retained", T)[t]
            elif setting(i, "limit", True):
                declared = min(R(u, "proposed")[t], max(0.0, cash_prev))
            else:
                declared = R(u, "proposed")[t]
            R(u, "declared")[t] = declared
            R(u, "retained")[t] = hist(i, "retained", T)[t] if a else re_prev + S("npat")[t] - declared
            payable_prev = P(u, "payable", t, "payable", inst=i)
            if a:
                paid = payable_prev + declared - hist(i, "payable", T)[t]
            else:
                d = min(max(setting(i, "delay", 3), 0), 6)
                hist_decl = init.get("declared_before", [0.0] * 6)   # declared 1..6 months before the first month
                paid = R(u, "declared")[t - d] if t - d >= 0 else (hist_decl[d - t - 1] if d - t - 1 < len(hist_decl) else 0.0)
            R(u, "paid")[t] = paid
            R(u, "payable")[t] = hist(i, "payable", T)[t] if a else payable_prev + declared - paid

        # ---- statements: balance sheet and cash flow
        cf = {
            "receipts": ssum(revs, "cash_sales") + ssum(colls, "receipts"),
            "other_operating": ssum(oinc, "income") + ssum(oexp, "cash") + ssum(others["fm.other_ca"], "cash") + ssum(others["fm.other_cl"], "cash"),
            "interest_received": R(icash.uid, "interest")[t] if icash else 0.0,
            "payments": ssum(cogs, "cash_paid") + ssum(opex, "cash_paid") + ssum(stocks, "cash_paid") + ssum(pays, "payments"),
            "staff": R(payroll.uid, "paid")[t] if payroll else 0.0,
            "gst": R(gst.uid, "paid")[t] if gst else 0.0,
            "interest_paid": ssum(debts, "interest"),
            "tax_paid": R(itax.uid, "paid")[t] if itax else 0.0,
            "capex": ssum(capexes, "paid"),
            "other_investing": ssum(others["fm.other_nca"], "cash"),
            "debt_drawn": ssum(debts, "drawn"),
            "equity_issued": R(equity.uid, "issued")[t] if equity else 0.0,
            "debt_repaid": ssum(debts, "repaid"),
            "dividends_paid": R(equity.uid, "paid")[t] if equity else 0.0,
            "other_financing": ssum(others["fm.other_ncl"], "cash") + ssum(others["fm.other_equity"], "cash"),
        }
        for k, v in cf.items():
            S("cf." + k)[t] = v
        S("cf_operating")[t] = (cf["receipts"] + cf["other_operating"] + cf["interest_received"] - cf["payments"] - cf["staff"]
                                - cf["gst"] - cf["interest_paid"] - cf["tax_paid"])
        S("cf_investing")[t] = cf["other_investing"] - cf["capex"]
        S("cf_financing")[t] = cf["debt_drawn"] + cf["equity_issued"] - cf["debt_repaid"] - cf["dividends_paid"] + cf["other_financing"]
        S("net_cash")[t] = S("cf_operating")[t] + S("cf_investing")[t] + S("cf_financing")[t]
        cash_prev = R("fs", "cash")[t - 1] if t > 0 else (opening(fs, "cash") if fs else 0.0)
        S("cash")[t] = hist(fs, "cash", T)[t] if (a and fs) else cash_prev + S("net_cash")[t]
        S("current_assets")[t] = S("cash")[t] + ssum(colls, "closing") + ssum(stocks, "closing") + ssum(others["fm.other_ca"], "balance")
        S("noncurrent_assets")[t] = ssum(fas, "nbv") + ssum(ints, "nbv") + ssum(others["fm.other_nca"], "balance")
        S("assets")[t] = S("current_assets")[t] + S("noncurrent_assets")[t]
        S("current_liabilities")[t] = (ssum(pays, "closing") + (R(payroll.uid, "payable")[t] if payroll else 0.0)
                                       + (R(gst.uid, "payable")[t] if gst else 0.0) + (R(itax.uid, "payable")[t] if itax else 0.0)
                                       + (R(equity.uid, "payable")[t] if equity else 0.0) + ssum(others["fm.other_cl"], "balance"))
        S("noncurrent_liabilities")[t] = (R(payroll.uid, "leave")[t] if payroll else 0.0) + ssum(debts, "balance") + ssum(others["fm.other_ncl"], "balance")
        S("liabilities")[t] = S("current_liabilities")[t] + S("noncurrent_liabilities")[t]
        S("net_assets")[t] = S("assets")[t] - S("liabilities")[t]
        S("equity")[t] = ((R(equity.uid, "capital")[t] + R(equity.uid, "retained")[t]) if equity else 0.0) + ssum(others["fm.other_equity"], "balance")
        S("debt")[t] = ssum(debts, "balance")

    return rows
