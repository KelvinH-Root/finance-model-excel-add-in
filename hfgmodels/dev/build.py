"""Development model: sheet definitions.

Sheets (inserted into the template's Financial Model section):
  Dev_Sites    inputs: model-wide, then one block per site
  Dev_Costs    cost phasing by line, GST on costs, cumulative cost
  Dev_Sales    settlement, revenue, output GST, two-step exit (SPV, Fund)
  Dev_GST      GST position, settlement with IRD, balance
  Dev_Funding  facility ceiling, finance costs, draws, equity plug, ratios
  Dev_FS       WIP, income statement, balance sheet, cash flow, checks
  Dev_Returns  site and portfolio returns
  Dev_Checks   every error and alert check in plain English (Appendices)

Each time series sheet starts with a portfolio block (sites flagged for
inclusion), followed by one block per site in the same order as Dev_Sites.
"""

from __future__ import annotations

from ..layout import Model
from .recipe import FUNDING_METHODS, GST_TREATMENTS, LIMIT_BASES, PROFILES, Recipe

YN = '"Yes,No"'
L = "DD_Ts_Last_Hist_Mth"  # last actual month (timeline index)


def _list(values) -> str:
    return '"' + ",".join(values) + '"'


def month_idx(d: str) -> str:
    return f"(YEAR({d})-YEAR(Ts_Start_Date))*12+MONTH({d})-MONTH(Ts_Start_Date)+1"


def build_model(r: Recipe) -> Model:
    m = Model(r.periods)
    sites = r.sites
    items = r.line_items
    fc = "dev.flag.fc"

    # ============================================================ Dev_Sites
    S = m.sheet("Dev_Sites", "Development Sites", timeline=False, after=r.insert_after, last_col="O")
    S.extra_cols = [(8, 8, 16.0), (9, 9, 12.0), (10, 15, 11.0)]
    S.section("Model-wide assumptions")
    S.scalar("g.gst_rate", "GST rate", r.gst_rate, style="in_pct", name="Dev_GST_Rate", input=True)
    S.scalar("g.gst_lag", "GST settled with IRD after (months)", r.gst_lag, style="in_num", name="Dev_GST_Lag", input=True)
    S.scalar("g.fund_lvr", "Te Piringa (NZHF) purchase debt on BTR price", r.fund_lvr, style="in_pct",
             name="Dev_Fund_LVR", input=True)
    S.scalar("g.last_actual", "Last actual month (set on the Time sheet)", fn=lambda x: f"={L}", style="num")
    S.scalar("g.last_actual_date", "Last actual month end", fn=lambda x: "=Ts_Last_Hist_Mth_End_Date", style="date_d")

    for s in sites:
        k = s.code
        S.section(("f", f'"{k}: "&{k}_Name'))
        S.sub("General")
        S.scalar(f"{k}.in.name", "Site name", s.name, style="in_text", name=f"{k}_Name", input=True)
        S.scalar(f"{k}.in.entity", "Entity (development LP)", s.entity, style="in_text", name=f"{k}_Entity", input=True)
        S.scalar(f"{k}.in.include", "Include in portfolio", s.include, style="in_text", name=f"{k}_Include", input=True,
                 validation=YN)
        S.scalar(f"{k}.in.gst_reg", "Entity is GST registered", s.gst_registered, style="in_text", name=f"{k}_GST_Reg",
                 input=True, validation=YN)
        S.scalar(f"{k}.in.units", "Units (homes)", s.units, style="in_num", name=f"{k}_Units", input=True)
        S.blank()
        S.sub("Opening position at the last actual month")
        S.scalar(f"{k}.in.open_wip", "Work in progress (excluding GST)", s.opening_wip, style="in_num",
                 name=f"{k}_Open_WIP", input=True, unit="$")
        S.scalar(f"{k}.in.open_debt", "Facility balance", s.opening_debt, style="in_num", name=f"{k}_Open_Debt",
                 input=True, unit="$")
        S.blank()
        S.sub("Cost budget (excluding GST)")
        S.head("Line", cells={"H": ("Budget $", "head_r"), "I": ("Profile", "head_r"), "J": ("Start", "head_r"),
                              "K": ("Months", "head_r"), "L": ("GST on cost", "head_r"),
                              "M": ("Intergroup", "head_r"), "N": ("Start #", "head_r"), "O": ("End #", "head_r")})
        for ln in s.lines:
            S.table(f"{k}.line.{ln.key}", ln.label, cells={
                "H": (ln.budget, "in_num"),
                "I": (ln.profile, "in_text", _list(PROFILES)),
                "J": (ln.start, "in_date"),
                "K": (ln.months, "in_num"),
                "L": (ln.gst, "in_text", YN),
                "M": (ln.counterparty, "in_text"),
                "N": (lambda x, kk=f"{k}.line.{ln.key}": "=" + month_idx(x.cell(kk, "J")), "num"),
                "O": (lambda x, kk=f"{k}.line.{ln.key}": f"={x.cell(kk, 'N')}+IF({x.cell(kk, 'I')}=\"Lump sum\",1,{x.cell(kk, 'K')})-1", "num"),
            })
        first, lastk = f"{k}.line.{s.lines[0].key}", f"{k}.line.{s.lines[-1].key}"
        S.scalar(f"{k}.d.total_budget", "Total budget",
                 fn=lambda x, a=first, b=lastk: f"=SUM({x.cell(a, 'H')}:{x.cell(b, 'H')})", style="total",
                 name=f"{k}_Budget", unit="$")
        S.scalar(f"{k}.in.alpha", "S-curve shape: alpha", s.alpha, style="in_dec", name=f"{k}_Alpha", input=True)
        S.scalar(f"{k}.in.beta", "S-curve shape: beta (alpha = beta gives a symmetric curve)", s.beta, style="in_dec",
                 name=f"{k}_Beta", input=True)
        S.blank()
        S.sub("Sale and exit")
        S.scalar(f"{k}.in.sale_date", "Settlement date", s.sale_date, style="in_date", name=f"{k}_Sale_Date", input=True)
        S.scalar(f"{k}.in.sale_price", "Sale price: As Complete (BTS) value, including any GST", s.sale_price,
                 style="in_num", name=f"{k}_Sale_Price", input=True, unit="$")
        S.scalar(f"{k}.in.sale_gst", "GST on the sale", s.sale_gst, style="in_text", name=f"{k}_Sale_GST", input=True,
                 validation=_list(GST_TREATMENTS))
        S.scalar(f"{k}.in.via_spv", "Sold to an SPV, SPV sold on to the Fund", s.via_spv, style="in_text",
                 name=f"{k}_Via_SPV", input=True, validation=YN)
        S.scalar(f"{k}.in.btr", "Build to Rent (BTR) value paid by the Fund", s.btr_value, style="in_num",
                 name=f"{k}_BTR", input=True, unit="$")
        S.blank()
        S.sub("Funding")
        S.scalar(f"{k}.in.method", "Funding method", s.method, style="in_text", name=f"{k}_Fund_Method", input=True,
                 validation=_list(FUNDING_METHODS))
        S.scalar(f"{k}.in.eq_commit", "Equity committed before debt (equity first method)", s.equity_commitment,
                 style="in_num", name=f"{k}_Equity_Commit", input=True, unit="$")
        S.scalar(f"{k}.in.limit_basis", "Facility limit basis", s.limit_basis, style="in_text", name=f"{k}_Limit_Basis",
                 input=True, validation=_list(LIMIT_BASES))
        S.scalar(f"{k}.in.limit_override", "Facility limit (when basis is Override)", s.limit_override, style="in_num",
                 name=f"{k}_Limit_Override", input=True, unit="$")
        S.scalar(f"{k}.in.max_ltc", "Maximum loan to cost (LTC)", s.max_ltc, style="in_pct", name=f"{k}_Max_LTC", input=True)
        S.scalar(f"{k}.in.max_lvr", "Maximum loan to value (LVR) on As Complete value", s.max_lvr, style="in_pct",
                 name=f"{k}_Max_LVR", input=True)
        S.scalar(f"{k}.in.debt_start", "Facility available from", s.debt_start, style="in_date", name=f"{k}_Debt_Start",
                 input=True)
        S.scalar(f"{k}.in.land_loan", "Pre-development (land) loan held until then", s.land_loan, style="in_num",
                 name=f"{k}_Land_Loan", input=True, unit="$")
        S.scalar(f"{k}.in.rate", "Interest rate (all-in, per year)", s.rate, style="in_pct2", name=f"{k}_Rate", input=True)
        S.scalar(f"{k}.in.line_fee", "Line fee on the limit (per year)", s.line_fee, style="in_pct2", name=f"{k}_Line_Fee",
                 input=True)
        S.scalar(f"{k}.in.estab_fee", "Establishment fee on the limit", s.estab_fee, style="in_pct2",
                 name=f"{k}_Estab_Fee", input=True)
        S.scalar(f"{k}.in.cap_fin", "Capitalise finance costs into the facility", s.capitalise, style="in_text",
                 name=f"{k}_Cap_Fin", input=True, validation=YN)
        S.scalar(f"{k}.in.release", "Draw to the ceiling and release equity", s.release_equity, style="in_text",
                 name=f"{k}_Release", input=True, validation=YN)
        S.blank()
        S.sub("Accounting")
        S.scalar(f"{k}.in.cap_borrow", "Capitalise borrowing costs to WIP (PBE IPSAS 5 policy)", s.cap_borrowing,
                 style="in_text", name=f"{k}_Cap_Borrow", input=True, validation=YN)
        S.blank()
        S.sub("Derived values")
        S.scalar(f"{k}.d.sale_idx", "Settlement month number", fn=lambda x, k=k: "=" + month_idx(f"{k}_Sale_Date"),
                 name=f"{k}_Sale_Idx")
        S.scalar(f"{k}.d.debt_idx", "Facility start month number", fn=lambda x, k=k: "=" + month_idx(f"{k}_Debt_Start"),
                 name=f"{k}_Debt_Idx")
        S.scalar(f"{k}.d.limit", "Facility limit", name=f"{k}_Limit", unit="$", fn=lambda x, k=k: (
            f'=IF({k}_Limit_Basis="Override",{k}_Limit_Override,IF({k}_Limit_Basis="LTC on budget",{k}_Max_LTC*{k}_Budget,'
            f'IF({k}_Limit_Basis="LVR on As Complete",{k}_Max_LVR*{k}_Sale_Price,{k}_Max_LVR*{k}_BTR)))'))
        S.scalar(f"{k}.d.revenue", "Sale revenue (excluding GST)", name=f"{k}_Revenue", unit="$", fn=lambda x, k=k: (
            f'=IF(AND({k}_GST_Reg="Yes",{k}_Sale_GST="Standard-rated"),{k}_Sale_Price/(1+Dev_GST_Rate),{k}_Sale_Price)'))
        S.scalar(f"{k}.d.output_gst", "Output GST on the sale", name=f"{k}_Output_GST", unit="$",
                 fn=lambda x, k=k: f"={k}_Sale_Price-{k}_Revenue")
        S.scalar(f"{k}.d.open_equity", "Opening equity (WIP less facility)", name=f"{k}_Open_Equity", unit="$",
                 fn=lambda x, k=k: f"={k}_Open_WIP-{k}_Open_Debt")

    def inc(k):
        return f'IF({k}_Include="Yes",1,0)'

    def port(sheet, key, label, item, style="num", unit="$", total="sum", indent=1):
        sheet.ts(key, label, lambda x, item=item: "=" + "+".join(f"{x.v(f'{s.code}.{item}')}*{inc(s.code)}" for s in sites),
                 style=style, unit=unit, total=total, indent=indent)

    # ============================================================ Dev_Costs
    C = m.sheet("Dev_Costs", "Development Costs", timeline=True, after="Dev_Sites")
    C.section("Timing")
    C.ts(fc, "Forecast month (after the last actual month)", lambda x: f"=IF({x.cnt}>{L},1,0)", unit="flag",
         total="sum")
    C.section("Portfolio")
    port(C, "port.cost.total", "Development cost (excluding GST)", "cost.total")
    port(C, "port.cost.input_gst", "GST on costs", "cost.input_gst")
    port(C, "port.cost.paid", "Costs paid (including GST)", "cost.paid", )
    for s in sites:
        k = s.code
        C.section(("f", f'"{k}: "&{k}_Name'))
        C.sub("Phasing share by line")
        for ln in s.lines:
            lk = f"{k}.line.{ln.key}"

            def share(x, lk=lk, k=k):
                prof, n, st = x.cell(lk, "I"), x.cell(lk, "K"), x.cell(lk, "N")
                kk = f"({x.cnt}-{st}+1)"
                return (f'=IF({prof}="Lump sum",IF({kk}=1,1,0),IF(AND({kk}>=1,{kk}<={n}),IF({prof}="Flat",1/{n},'
                        f"_xlfn.BETA.DIST({kk}/{n},{k}_Alpha,{k}_Beta,TRUE)-_xlfn.BETA.DIST(({kk}-1)/{n},{k}_Alpha,{k}_Beta,TRUE)),0))")
            C.ts(f"{k}.share.{ln.key}", ln.label, share, style="pct", unit="%", outline=1)
        C.sub("Forecast cost (excluding GST)")
        C.ts(f"{k}.cost.phased_fc", "Budget phased into forecast months", lambda x, k=k, s=s: "=(" + "+".join(
            f"{x.cell(f'{k}.line.{ln.key}', 'H')}*{x.v(f'{k}.share.{ln.key}')}" for ln in s.lines) + f")*{x.v(fc)}")
        C.scalar(f"{k}.cost.scale", "Cost to complete scaling factor",
                 name=f"{k}_Cost_Scale", style="dec",
                 fn=lambda x, k=k: f"=IF(ABS({x.tot(f'{k}.cost.phased_fc')})<0.005,0,({k}_Budget-{k}_Open_WIP)/{x.tot(f'{k}.cost.phased_fc')})")
        for ln in s.lines:
            C.ts(f"{k}.cost.{ln.key}", ln.label, lambda x, k=k, ln=ln: (
                f"={x.cell(f'{k}.line.{ln.key}', 'H')}*{x.v(f'{k}.share.{ln.key}')}*{x.v(fc)}*{k}_Cost_Scale"), indent=2)
        C.ts(f"{k}.cost.total", "Total development cost", lambda x, k=k, s=s: (
            f"=SUM({x.v(f'{k}.cost.{s.lines[0].key}')}:{x.v(f'{k}.cost.{s.lines[-1].key}')})"), style="total")
        C.ts(f"{k}.cost.input_gst", "GST on costs", lambda x, k=k, s=s: "=Dev_GST_Rate*(" + "+".join(
            f"{x.v(f'{k}.cost.{ln.key}')}*IF({x.cell(f'{k}.line.{ln.key}', 'L')}=\"Yes\",1,0)" for ln in s.lines) + ")")
        C.ts(f"{k}.cost.unrec_gst", "GST not recoverable (entity not registered)", lambda x, k=k: (
            f'={x.v(f"{k}.cost.input_gst")}*IF({k}_GST_Reg="Yes",0,1)'))
        C.ts(f"{k}.cost.paid", "Costs paid (including GST)", lambda x, k=k: (
            f"={x.v(f'{k}.cost.total')}+{x.v(f'{k}.cost.input_gst')}"))
        C.ts(f"{k}.cost.cum", "Cumulative cost for loan to cost (includes opening WIP)", total="last", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{k}_Open_WIP+{x.v(f'{k}.cost.total')}+{x.v(f'{k}.cost.unrec_gst')},0)" if x.first else
            f"=IF({x.v(fc)}=1,IF({x.cnt}={L}+1,{k}_Open_WIP,{x.prev(f'{k}.cost.cum')})+{x.v(f'{k}.cost.total')}+{x.v(f'{k}.cost.unrec_gst')},0)"))

    # ============================================================ Dev_Sales
    A = m.sheet("Dev_Sales", "Development Sales and Exit", timeline=True, after="Dev_Costs")
    A.section("Portfolio")
    port(A, "port.sale.revenue", "Sale revenue (excluding GST)", "sale.revenue")
    port(A, "port.sale.output_gst", "Output GST on sales", "sale.output_gst")
    port(A, "port.sale.proceeds", "Sale proceeds (including GST)", "sale.proceeds")
    port(A, "port.exit.hhlp_margin", "Margin to HHLP on the sale to the Fund", "exit.hhlp_margin")
    port(A, "port.exit.gst_cost", "GST on the internal sale (group cost)", "exit.gst_cost")
    port(A, "port.exit.fund_price", "Te Piringa (NZHF) purchases at BTR", "exit.fund_price")
    port(A, "port.exit.fund_debt", "Te Piringa debt at the fund LVR", "exit.fund_debt")
    port(A, "port.exit.fund_equity", "Te Piringa equity required", "exit.fund_equity")
    for s in sites:
        k = s.code
        A.section(("f", f'"{k}: "&{k}_Name'))
        A.sub("Settlement")
        A.ts(f"{k}.sale.flag", "Settlement month", lambda x, k=k: f"=IF(AND({x.cnt}={k}_Sale_Idx,{x.v(fc)}=1),1,0)",
             unit="flag")
        A.ts(f"{k}.sale.sold", "Sold (settlement month and after)", lambda x, k=k: f"=IF({x.cnt}>={k}_Sale_Idx,1,0)",
             unit="flag", total="none")
        A.ts(f"{k}.sale.revenue", "Sale revenue (excluding GST)", lambda x, k=k: f"={k}_Revenue*{x.v(f'{k}.sale.flag')}")
        A.ts(f"{k}.sale.output_gst", "Output GST", lambda x, k=k: f"={k}_Output_GST*{x.v(f'{k}.sale.flag')}")
        A.ts(f"{k}.sale.proceeds", "Sale proceeds (including GST)", lambda x, k=k: f"={k}_Sale_Price*{x.v(f'{k}.sale.flag')}",
             style="total")
        A.sub("Exit through the SPV to the Fund")
        via = f'IF({k}_Via_SPV="Yes",1,0)'
        A.ts(f"{k}.exit.spv_price", "SPV purchase price (As Complete, including GST)",
             lambda x, k=k, via=via: f"={k}_Sale_Price*{via}*{x.v(f'{k}.sale.flag')}")
        A.ts(f"{k}.exit.fund_price", "Fund purchase price (BTR)",
             lambda x, k=k, via=via: f"={k}_BTR*{via}*{x.v(f'{k}.sale.flag')}")
        A.ts(f"{k}.exit.hhlp_margin", "Margin to HHLP", lambda x, k=k: (
            f"={x.v(f'{k}.exit.fund_price')}-{x.v(f'{k}.exit.spv_price')}"), style="total")
        A.ts(f"{k}.exit.gst_cost", "GST on the internal sale (group cost)",
             lambda x, k=k, via=via: f"={k}_Output_GST*{via}*{x.v(f'{k}.sale.flag')}")
        A.ts(f"{k}.exit.fund_debt", "Fund debt at the fund LVR",
             lambda x, k=k: f"={x.v(f'{k}.exit.fund_price')}*Dev_Fund_LVR")
        A.ts(f"{k}.exit.fund_equity", "Fund equity required",
             lambda x, k=k: f"={x.v(f'{k}.exit.fund_price')}-{x.v(f'{k}.exit.fund_debt')}")

    # ============================================================ Dev_GST
    G = m.sheet("Dev_GST", "Development GST", timeline=True, after="Dev_Sales")
    G.section("Portfolio")
    port(G, "port.gst.net", "Net GST for the month: payable / (refundable)", "gst.net")
    port(G, "port.gst.cash", "GST paid to / (refunded by) IRD", "gst.cash")
    port(G, "port.gst.balance", "GST payable / (receivable)", "gst.balance", total="last")
    for s in sites:
        k = s.code
        G.section(("f", f'"{k}: "&{k}_Name'))
        G.ts(f"{k}.gst.net", "Net GST for the month: payable / (refundable)", lambda x, k=k: (
            f'=({x.v(f"{k}.sale.output_gst")}-{x.v(f"{k}.cost.input_gst")})*IF({k}_GST_Reg="Yes",1,0)'))
        G.ts(f"{k}.gst.cash", "GST paid to / (refunded by) IRD", lambda x, k=k: (
            f"=IF({x.cnt}-Dev_GST_Lag>{L},INDEX({x.rng(f'{k}.gst.net')},{x.cnt}-Dev_GST_Lag),0)"))
        G.ts(f"{k}.gst.balance", "GST payable / (receivable)", total="last", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{x.v(f'{k}.gst.net')}-{x.v(f'{k}.gst.cash')},0)" if x.first else
            f"=IF({x.v(fc)}=1,IF({x.cnt}={L}+1,0,{x.prev(f'{k}.gst.balance')})+{x.v(f'{k}.gst.net')}-{x.v(f'{k}.gst.cash')},0)"))

    # ============================================================ Dev_Funding
    F = m.sheet("Dev_Funding", "Development Funding", timeline=True, after="Dev_GST")
    F.section("Portfolio")
    port(F, "port.fund.close", "Facility balance", "fund.close", total="max")
    port(F, "port.fund.fin_cost", "Finance costs", "fund.fin_cost")
    port(F, "port.fund.draw", "Facility drawn / (repaid)", "fund.draw")
    port(F, "port.fund.equity", "Equity contributed / (distributed)", "fund.equity")
    port(F, "port.fund.eq_bal", "Net equity invested", "fund.eq_bal", total="max")
    port(F, "port.fund.cash", "Cash held for GST payable", "fund.cash", total="last")
    port(F, "port.fund.irr_eq", "Equity cash flows for IRR", "fund.irr_eq")
    port(F, "port.fund.irr_proj", "Project cash flows for IRR (before finance)", "fund.irr_proj")
    for s in sites:
        k = s.code
        v = lambda key, k=k: f"{k}.{key}"  # noqa: E731
        F.section(("f", f'"{k}: "&{k}_Name'))
        F.sub("Facility ceiling")
        F.ts(v("fund.ceiling"), "Debt ceiling", total="max", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=0,0,IF({x.v(f'{k}.sale.sold')}=1,0,IF({x.cnt}>={k}_Debt_Idx,MIN({k}_Limit,"
            f'IF({k}_Fund_Method="Equity first",MAX(0,{x.v(f"{k}.cost.cum")}-{k}_Equity_Commit),{k}_Max_LTC*{x.v(f"{k}.cost.cum")})),'
            f"{k}_Land_Loan)))"))
        F.sub("Finance costs")
        F.ts(v("fund.open"), "Opening facility balance", total="none", fn=lambda x, k=k: (
            f"=IF({x.cnt}={L}+1,{k}_Open_Debt,0)" if x.first else
            f"=IF({x.cnt}={L}+1,{k}_Open_Debt,IF({x.cnt}>{L},{x.prev(f'{k}.fund.close')},0))"))
        F.ts(v("fund.interest"), "Interest on the opening balance", lambda x, k=k: f"={x.v(f'{k}.fund.open')}*{k}_Rate/12")
        F.ts(v("fund.line_fee"), "Line fee", lambda x, k=k: (
            f"=IF(AND({x.v(fc)}=1,{x.cnt}>={k}_Debt_Idx,{x.v(f'{k}.sale.sold')}=0),{k}_Limit*{k}_Line_Fee/12,0)"))
        F.ts(v("fund.estab_fee"), "Establishment fee", lambda x, k=k: (
            f"=IF(AND({x.v(fc)}=1,{x.cnt}={k}_Debt_Idx),{k}_Limit*{k}_Estab_Fee,0)"))
        F.ts(v("fund.fin_cost"), "Finance costs", lambda x, k=k: (
            f"={x.v(f'{k}.fund.interest')}+{x.v(f'{k}.fund.line_fee')}+{x.v(f'{k}.fund.estab_fee')}"), style="total")
        F.ts(v("fund.headroom"), "Headroom under the ceiling", total="none",
             fn=lambda x, k=k: f"=MAX(0,{x.v(f'{k}.fund.ceiling')}-{x.v(f'{k}.fund.open')})")
        F.ts(v("fund.fin_cap"), "Finance costs capitalised (within the ceiling)", lambda x, k=k: (
            f'=IF({k}_Cap_Fin="Yes",MIN({x.v(f"{k}.fund.fin_cost")},{x.v(f"{k}.fund.headroom")}),0)'))
        F.ts(v("fund.fin_cash"), "Finance costs paid in cash", lambda x, k=k: (
            f"={x.v(f'{k}.fund.fin_cost')}-{x.v(f'{k}.fund.fin_cap')}"))
        F.sub("Draws and equity")
        F.ts(v("fund.need"), "Cash needed before funding (costs and GST less sale proceeds)", lambda x, k=k: (
            f"={x.v(f'{k}.cost.paid')}+{x.v(f'{k}.gst.cash')}-{x.v(f'{k}.sale.proceeds')}"))
        F.ts(v("fund.gap"), "Movement to the ceiling", total="none", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{x.v(f'{k}.fund.ceiling')}-{x.v(f'{k}.fund.open')}-{x.v(f'{k}.fund.fin_cap')},0)"))
        F.ts(v("fund.draw"), "Facility drawn / (repaid)", lambda x, k=k: (
            f'=IF(OR({x.v(f"{k}.fund.gap")}<0,{k}_Release="Yes"),{x.v(f"{k}.fund.gap")},'
            f'MIN({x.v(f"{k}.fund.gap")},MAX(0,{x.v(f"{k}.fund.need")}+{x.v(f"{k}.fund.fin_cash")})))'))
        F.ts(v("fund.close"), "Closing facility balance", total="max", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fund.open')}+{x.v(f'{k}.fund.fin_cap')}+{x.v(f'{k}.fund.draw')}"))
        F.ts(v("fund.cash_open"), "Opening cash held", total="none", fn=lambda x, k=k: (
            "=0" if x.first else f"=IF({x.cnt}={L}+1,0,{x.prev(f'{k}.fund.cash')})"))
        F.ts(v("fund.cash"), "Cash held for GST payable", total="last", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,MAX(0,{x.v(f'{k}.gst.balance')}),0)"))
        F.ts(v("fund.equity"), "Equity contributed / (distributed)", style="total", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{x.v(f'{k}.fund.need')}+{x.v(f'{k}.fund.fin_cash')}-{x.v(f'{k}.fund.draw')}"
            f"+{x.v(f'{k}.fund.cash')}-{x.v(f'{k}.fund.cash_open')},0)"))
        F.ts(v("fund.contrib"), "Equity contributed", lambda x, k=k: f"=MAX(0,{x.v(f'{k}.fund.equity')})")
        F.ts(v("fund.distrib"), "Equity distributed", lambda x, k=k: f"=MAX(0,-{x.v(f'{k}.fund.equity')})")
        F.ts(v("fund.eq_bal"), "Net equity invested (includes opening equity)", total="max", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{k}_Open_Equity+{x.v(f'{k}.fund.equity')},0)" if x.first else
            f"=IF({x.v(fc)}=1,IF({x.cnt}={L}+1,{k}_Open_Equity,{x.prev(f'{k}.fund.eq_bal')})+{x.v(f'{k}.fund.equity')},0)"))
        F.sub("Covenant ratios")
        F.ts(v("fund.ltc"), "Loan to cost", style="pct", unit="%", total="max", fn=lambda x, k=k: (
            f"=IF({x.v(f'{k}.cost.cum')}>0,{x.v(f'{k}.fund.close')}/{x.v(f'{k}.cost.cum')},0)"))
        F.ts(v("fund.lvr"), "Loan to As Complete value", style="pct", unit="%", total="max", fn=lambda x, k=k: (
            f"=IF({k}_Sale_Price>0,{x.v(f'{k}.fund.close')}/{k}_Sale_Price,0)"))
        F.sub("Cash flows for returns")
        F.ts(v("fund.irr_eq"), "Equity cash flows (opening equity in the first forecast month)", lambda x, k=k: (
            f"=IF({x.v(fc)}=1,-{x.v(f'{k}.fund.equity')}-IF({x.cnt}={L}+1,{k}_Open_Equity,0),0)"))
        F.ts(v("fund.irr_proj"), "Project cash flows before finance (opening WIP in the first forecast month)",
             lambda x, k=k: (f"=IF({x.v(fc)}=1,-{x.v(f'{k}.fund.need')}-({x.v(f'{k}.fund.cash')}-{x.v(f'{k}.fund.cash_open')})"
                             f"-IF({x.cnt}={L}+1,{k}_Open_WIP,0),0)"))

    # ============================================================ Dev_FS
    FS = m.sheet("Dev_FS", "Development Financial Statements", timeline=True, after="Dev_Funding")
    FS.section("Portfolio")
    port(FS, "port.fs.revenue", "Revenue", "fs.revenue")
    port(FS, "port.fs.cos", "Cost of sales", "fs.cos")
    port(FS, "port.fs.fin_exp", "Finance costs expensed", "fs.fin_exp")
    port(FS, "port.fs.profit", "Net surplus / (deficit)", "fs.profit")
    port(FS, "port.fs.wip", "Work in progress", "fs.wip", total="last")
    port(FS, "port.fs.bs_check", "Balance sheet check", "fs.bs_check", total="none")
    for s in sites:
        k = s.code
        FS.section(("f", f'"{k}: "&{k}_Name'))
        FS.sub("Work in progress")
        FS.ts(f"{k}.fs.wip_open", "Opening WIP", total="none", fn=lambda x, k=k: (
            f"=IF({x.cnt}={L}+1,{k}_Open_WIP,0)" if x.first else
            f"=IF({x.cnt}={L}+1,{k}_Open_WIP,IF({x.v(fc)}=1,{x.prev(f'{k}.fs.wip')},0))"))
        FS.ts(f"{k}.fs.fin_wip", "Borrowing costs capitalised to WIP", lambda x, k=k: (
            f'=IF({k}_Cap_Borrow="Yes",{x.v(f"{k}.fund.fin_cost")},0)'))
        FS.ts(f"{k}.fs.wip_adds", "Additions (costs, unrecoverable GST, capitalised borrowing costs)", lambda x, k=k: (
            f"={x.v(f'{k}.cost.total')}+{x.v(f'{k}.cost.unrec_gst')}+{x.v(f'{k}.fs.fin_wip')}"))
        FS.ts(f"{k}.fs.cos", "Released to cost of sales", lambda x, k=k: (
            f"=IF(AND({x.v(fc)}=1,{x.v(f'{k}.sale.sold')}=1),{x.v(f'{k}.fs.wip_open')}+{x.v(f'{k}.fs.wip_adds')},0)"))
        FS.ts(f"{k}.fs.wip", "Closing WIP", total="last", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.wip_open')}+{x.v(f'{k}.fs.wip_adds')}-{x.v(f'{k}.fs.cos')}"))
        FS.sub("Income statement")
        FS.ts(f"{k}.fs.revenue", "Revenue", lambda x, k=k: f"={x.v(f'{k}.sale.revenue')}")
        FS.ts(f"{k}.fs.cos_pl", "Cost of sales", lambda x, k=k: f"=-{x.v(f'{k}.fs.cos')}")
        FS.ts(f"{k}.fs.fin_exp", "Finance costs expensed", lambda x, k=k: (
            f"={x.v(f'{k}.fund.fin_cost')}-{x.v(f'{k}.fs.fin_wip')}"))
        FS.ts(f"{k}.fs.profit", "Net surplus / (deficit)", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.revenue')}+{x.v(f'{k}.fs.cos_pl')}-{x.v(f'{k}.fs.fin_exp')}"))
        FS.sub("Balance sheet")
        FS.ts(f"{k}.fs.bs_cash", "Cash", total="last", fn=lambda x, k=k: f"={x.v(f'{k}.fund.cash')}")
        FS.ts(f"{k}.fs.bs_wip", "Work in progress", total="last", fn=lambda x, k=k: f"={x.v(f'{k}.fs.wip')}")
        FS.ts(f"{k}.fs.bs_assets", "Total assets", total="last", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.bs_cash')}+{x.v(f'{k}.fs.bs_wip')}"))
        FS.ts(f"{k}.fs.bs_debt", "Development facility", total="last", fn=lambda x, k=k: f"={x.v(f'{k}.fund.close')}")
        FS.ts(f"{k}.fs.bs_gst", "GST payable / (receivable)", total="last", fn=lambda x, k=k: f"={x.v(f'{k}.gst.balance')}")
        FS.ts(f"{k}.fs.bs_liab", "Total liabilities", total="last", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.bs_debt')}+{x.v(f'{k}.fs.bs_gst')}"))
        FS.ts(f"{k}.fs.bs_eq", "Partners' capital (net of distributions)", total="last", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fund.eq_bal')}"))
        FS.ts(f"{k}.fs.re", "Retained surplus", total="last", fn=lambda x, k=k: (
            f"=IF({x.v(fc)}=1,{x.v(f'{k}.fs.profit')},0)" if x.first else
            f"=IF({x.v(fc)}=1,IF({x.cnt}={L}+1,0,{x.prev(f'{k}.fs.re')})+{x.v(f'{k}.fs.profit')},0)"))
        FS.ts(f"{k}.fs.bs_equity", "Total equity", total="last", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.bs_eq')}+{x.v(f'{k}.fs.re')}"))
        FS.ts(f"{k}.fs.bs_check", "Balance check (assets less liabilities and equity)", style="check", total="none",
              fn=lambda x, k=k: f"=ROUND({x.v(f'{k}.fs.bs_assets')}-{x.v(f'{k}.fs.bs_liab')}-{x.v(f'{k}.fs.bs_equity')},2)")
        FS.sub("Cash flow statement")
        FS.ts(f"{k}.fs.cf_sale", "Sale proceeds received", lambda x, k=k: f"={x.v(f'{k}.sale.proceeds')}")
        FS.ts(f"{k}.fs.cf_costs", "Development costs paid", lambda x, k=k: f"=-{x.v(f'{k}.cost.paid')}")
        FS.ts(f"{k}.fs.cf_gst", "GST (paid to) / refunded by IRD", lambda x, k=k: f"=-{x.v(f'{k}.gst.cash')}")
        FS.ts(f"{k}.fs.cf_fin", "Finance costs paid", lambda x, k=k: f"=-{x.v(f'{k}.fund.fin_cash')}")
        FS.ts(f"{k}.fs.cf_ops", "Cash flow from development", style="total", fn=lambda x, k=k: (
            f"=SUM({x.v(f'{k}.fs.cf_sale')}:{x.v(f'{k}.fs.cf_fin')})"))
        FS.ts(f"{k}.fs.cf_draw", "Facility drawn", lambda x, k=k: f"=MAX(0,{x.v(f'{k}.fund.draw')})")
        FS.ts(f"{k}.fs.cf_repay", "Facility repaid", lambda x, k=k: f"=-MAX(0,-{x.v(f'{k}.fund.draw')})")
        FS.ts(f"{k}.fs.cf_contrib", "Equity contributed", lambda x, k=k: f"={x.v(f'{k}.fund.contrib')}")
        FS.ts(f"{k}.fs.cf_distrib", "Distributions to partners", lambda x, k=k: f"=-{x.v(f'{k}.fund.distrib')}")
        FS.ts(f"{k}.fs.cf_finance", "Cash flow from financing", style="total", fn=lambda x, k=k: (
            f"=SUM({x.v(f'{k}.fs.cf_draw')}:{x.v(f'{k}.fs.cf_distrib')})"))
        FS.ts(f"{k}.fs.cf_net", "Net change in cash", style="total", fn=lambda x, k=k: (
            f"={x.v(f'{k}.fs.cf_ops')}+{x.v(f'{k}.fs.cf_finance')}"))
        FS.ts(f"{k}.fs.cf_check", "Cash check (change in cash against the balance sheet)", style="check", total="none",
              fn=lambda x, k=k: f"=ROUND({x.v(f'{k}.fs.cf_net')}-({x.v(f'{k}.fund.cash')}-{x.v(f'{k}.fund.cash_open')}),2)")

    # ============================================================ Dev_Returns
    R = m.sheet("Dev_Returns", "Development Returns", timeline=False, after="Dev_FS",
                last_col=chr(ord("H") + len(sites)))
    R.extra_cols = [(8, 8 + len(sites), 14.0)]
    cols = {s.code: chr(ord("H") + i) for i, s in enumerate(sites)}
    pcol = chr(ord("H") + len(sites))
    R.section("Returns by site and portfolio")
    R.head("", cells={**{cols[s.code]: (("f", f"{s.code}_Name"), "head_r") for s in sites}, pcol: ("Portfolio", "head_r")})

    def metric(key, label, site_fn, port_fn, style="num"):
        cells = {cols[s.code]: (lambda x, s=s: "=" + site_fn(x, s.code), style) for s in sites}
        cells[pcol] = (lambda x: "=" + port_fn(x), style)
        R.table(key, label, cells)

    def sum_sites(key):
        return lambda x: "+".join(f"{x.cell(key, cols[s.code])}*{inc(s.code)}" for s in sites)

    R.sub("Development")
    metric("ret.units", "Units", lambda x, k: f"{k}_Units", sum_sites("ret.units"))
    metric("ret.revenue", "Sale revenue (excluding GST)", lambda x, k: x.tot(f"{k}.fs.revenue"), sum_sites("ret.revenue"))
    metric("ret.dev_cost", "Development cost (includes opening WIP)", lambda x, k: (
        f"{k}_Open_WIP+{x.tot(f'{k}.cost.total')}+{x.tot(f'{k}.cost.unrec_gst')}"), sum_sites("ret.dev_cost"))
    metric("ret.finance", "Finance costs (forecast months)", lambda x, k: x.tot(f"{k}.fund.fin_cost"),
           sum_sites("ret.finance"))
    metric("ret.profit", "Development surplus / (deficit)", lambda x, k: x.tot(f"{k}.fs.profit"), sum_sites("ret.profit"),
           style="total")
    metric("ret.margin", "Margin on revenue", lambda x, k: (
        f"IF({x.cell('ret.revenue', cols[k])}=0,0,{x.cell('ret.profit', cols[k])}/{x.cell('ret.revenue', cols[k])})"),
        lambda x: f"IF({x.cell('ret.revenue', pcol)}=0,0,{x.cell('ret.profit', pcol)}/{x.cell('ret.revenue', pcol)})",
        style="pct")
    metric("ret.poc", "Surplus on cost", lambda x, k: (
        f"IF(({x.tot(f'{k}.fs.cos')}+{x.tot(f'{k}.fs.fin_exp')})=0,0,{x.cell('ret.profit', cols[k])}/({x.tot(f'{k}.fs.cos')}+{x.tot(f'{k}.fs.fin_exp')}))"),
        lambda x: (f"IF(({x.tot('port.fs.cos')}+{x.tot('port.fs.fin_exp')})=0,0,{x.cell('ret.profit', pcol)}/"
                   f"({x.tot('port.fs.cos')}+{x.tot('port.fs.fin_exp')}))"), style="pct")
    R.sub("Funding")
    metric("ret.limit", "Facility limit", lambda x, k: f"{k}_Limit", sum_sites("ret.limit"))
    metric("ret.peak_debt", "Peak facility balance", lambda x, k: x.tot(f"{k}.fund.close"),
           lambda x: x.tot("port.fund.close"))
    metric("ret.peak_debt_month", "Month of peak facility balance", lambda x, k: (
        f"IF({x.cell('ret.peak_debt', cols[k])}=0,\"\",INDEX(Dev_Funding!$J$5:${x.last_letter}$5,MATCH({x.cell('ret.peak_debt', cols[k])},{x.rng(f'{k}.fund.close')},0)))"),
        lambda x: (f"IF({x.cell('ret.peak_debt', pcol)}=0,\"\",INDEX(Dev_Funding!$J$5:${x.last_letter}$5,"
                   f"MATCH({x.cell('ret.peak_debt', pcol)},{x.rng('port.fund.close')},0)))"), style="text_r")
    metric("ret.peak_equity", "Peak net equity invested", lambda x, k: x.tot(f"{k}.fund.eq_bal"),
           lambda x: x.tot("port.fund.eq_bal"))
    metric("ret.peak_ltc", "Peak loan to cost", lambda x, k: x.tot(f"{k}.fund.ltc"),
           lambda x: f"MAX({','.join(x.cell('ret.peak_ltc', cols[s.code]) for s in sites)})", style="pct")
    R.sub("Returns")
    metric("ret.eq_irr", "Equity IRR (per year, from the first forecast month)", lambda x, k: (
        f"IFERROR((1+IRR({x.rng(f'{k}.fund.irr_eq')},0.01))^12-1,0)"),
        lambda x: f"IFERROR((1+IRR({x.rng('port.fund.irr_eq')},0.01))^12-1,0)", style="pct")
    metric("ret.proj_irr", "Project IRR before finance (per year)", lambda x, k: (
        f"IFERROR((1+IRR({x.rng(f'{k}.fund.irr_proj')},0.01))^12-1,0)"),
        lambda x: f"IFERROR((1+IRR({x.rng('port.fund.irr_proj')},0.01))^12-1,0)", style="pct")
    metric("ret.multiple", "Equity multiple", lambda x, k: (
        f"IFERROR({x.tot(f'{k}.fund.distrib')}/({x.tot(f'{k}.fund.contrib')}+MAX(0,{k}_Open_Equity)),0)"),
        lambda x: ("IFERROR((" + "+".join(f"{x.tot(f'{s.code}.fund.distrib')}*{inc(s.code)}" for s in sites) + ")/(" +
                   "+".join(f"({x.tot(f'{s.code}.fund.contrib')}+MAX(0,{s.code}_Open_Equity))*{inc(s.code)}" for s in sites)
                   + "),0)"), style="multiple")
    R.sub("Exit and group view")
    metric("ret.hhlp_margin", "Margin to HHLP on the sale to the Fund", lambda x, k: x.tot(f"{k}.exit.hhlp_margin"),
           sum_sites("ret.hhlp_margin"))
    metric("ret.group_margin", "Group margin (surplus plus HHLP margin)", lambda x, k: (
        f"{x.cell('ret.profit', cols[k])}+{x.cell('ret.hhlp_margin', cols[k])}"), sum_sites("ret.group_margin"),
        style="total")
    metric("ret.gst_cost", "Of which GST on the internal sale", lambda x, k: x.tot(f"{k}.exit.gst_cost"),
           sum_sites("ret.gst_cost"))
    metric("ret.fund_price", "Te Piringa purchase price (BTR)", lambda x, k: x.tot(f"{k}.exit.fund_price"),
           sum_sites("ret.fund_price"))
    metric("ret.fund_debt", "Te Piringa debt at the fund LVR", lambda x, k: x.tot(f"{k}.exit.fund_debt"),
           sum_sites("ret.fund_debt"))
    metric("ret.fund_equity", "Te Piringa equity required", lambda x, k: x.tot(f"{k}.exit.fund_equity"),
           sum_sites("ret.fund_equity"))

    # ============================================================ Dev_Checks
    K = m.sheet("Dev_Checks", "Development Checks", timeline=False, after=r.checks_after,
                last_col=chr(ord("I") + len(sites) + 1))
    K.extra_cols = [(8, 8, 9.0), (9, 9 + len(sites) + 1, 12.0)]
    kcols = {s.code: chr(ord("I") + i) for i, s in enumerate(sites)}
    tcol = chr(ord("I") + len(sites))
    K.section("Summary")
    K.scalar("chk.errors", "Errors found (should be nil)", name="Dev_Err_Chk", style="check_flag",
             fn=lambda x: f"={x.cell('chk.err_total', tcol)}")
    K.scalar("chk.alerts", "Alerts raised", name="Dev_Alt_Chk", style="check_flag",
             fn=lambda x: f"={x.cell('chk.alt_total', tcol)}")
    K.scalar("chk.msg", "Status", style="text", fn=lambda x: (
        '=IF(Dev_Err_Chk>0,"Error: see the checks below",IF(Dev_Alt_Chk>0,Dev_Alt_Chk&" alert(s) to review","All checks pass"))'))
    K.section("Checks by site")
    K.head("Check", cells={"H": ("Type", "head_r"), **{kcols[s.code]: (("f", f"{s.code}_Name"), "head_r") for s in sites},
                           tcol: ("Total", "head_r")})
    last = None

    def check(key, label, kind, site_fn, port_fn=None):
        nonlocal last
        cells = {"H": (kind, "text_r")}
        for s in sites:
            cells[kcols[s.code]] = (lambda x, s=s: "=" + site_fn(x, s.code), "check")
        cells[tcol] = (lambda x, key=key: "=" + (port_fn(x) if port_fn else
                       f"SUM({x.cell(key, kcols[sites[0].code])}:{x.cell(key, kcols[sites[-1].code])})"), "check")
        K.table(key, label, cells)

    tol = "0.5"
    check("chk.budget", "Forecast cost plus opening WIP does not equal the budget", "Error",
          lambda x, k: f"IF(ABS({k}_Budget-{k}_Open_WIP-{x.tot(f'{k}.cost.total')})>{tol},1,0)")
    check("chk.limit", "Months with the facility above its limit", "Error",
          lambda x, k: f"SUMPRODUCT(--({x.rng(f'{k}.fund.close')}>{k}_Limit+{tol}))")
    check("chk.neg_debt", "Months with a negative facility balance", "Error",
          lambda x, k: f"SUMPRODUCT(--({x.rng(f'{k}.fund.close')}<-{tol}))")
    check("chk.bs", "Months where the balance sheet does not balance", "Error",
          lambda x, k: f"SUMPRODUCT(--(ABS({x.rng(f'{k}.fs.bs_check')})>{tol}))")
    check("chk.cf", "Months where the cash flow does not reconcile", "Error",
          lambda x, k: f"SUMPRODUCT(--(ABS({x.rng(f'{k}.fs.cf_check')})>{tol}))")
    check("chk.neg_cash", "Months with negative cash", "Error",
          lambda x, k: f"SUMPRODUCT(--({x.rng(f'{k}.fund.cash')}<-{tol}))")
    check("chk.port", "Portfolio facility balance differs from the included sites", "Error",
          lambda x, k: "0",
          lambda x: "SUMPRODUCT(--(ABS(" + x.rng("port.fund.close") + "-(" + "+".join(
              f"{x.rng(f'{s.code}.fund.close')}*{inc(s.code)}" for s in sites) + "))>" + tol + "))")
    check("chk.sale_window", "Settlement falls outside the forecast months", "Alert",
          lambda x, k: f"IF(OR({k}_Sale_Idx<={L},{k}_Sale_Idx>Ts_Term),1,0)")
    check("chk.debt_end", "Facility still outstanding at the end of the timeline", "Alert",
          lambda x, k: f"IF({x.last(f'{k}.fund.close')}>{tol},1,0)")
    check("chk.ltc", "Months above the maximum loan to cost", "Alert",
          lambda x, k: f"SUMPRODUCT(--({x.rng(f'{k}.fund.ltc')}>{k}_Max_LTC+0.000001))")
    check("chk.lvr", "Months above the maximum loan to value", "Alert",
          lambda x, k: f"SUMPRODUCT(--({x.rng(f'{k}.fund.lvr')}>{k}_Max_LVR+0.000001))")
    check("chk.late_cost", "Months with costs after the settlement month", "Alert",
          lambda x, k: f"SUMPRODUCT((Dev_Costs!$J$9:${x.last_letter}$9>{k}_Sale_Idx)*(ABS({x.rng(f'{k}.cost.total')})>{tol}))")
    check("chk.debt_after_sale", "Facility starts after settlement", "Alert",
          lambda x, k: f"IF({k}_Debt_Idx>={k}_Sale_Idx,1,0)")
    check("chk.gst_end", "GST still to settle with IRD at the end of the timeline", "Alert",
          lambda x, k: f"IF(ABS({x.last(f'{k}.gst.balance')})>{tol},1,0)")
    K.blank()

    def total(kind):
        def fn(x, c):
            _, r1 = x._where("chk.budget")
            _, r2 = x._where("chk.gst_end")
            return f'=SUMIF($H${r1}:$H${r2},"{kind}",{c}${r1}:{c}${r2})'
        return fn

    for key, label, kind in (("chk.err_total", "Total errors", "Error"), ("chk.alt_total", "Total alerts", "Alert")):
        fn = total(kind)
        K.table(key, label, {"H": (kind, "text_r"), **{
            c: (lambda x, c=c, fn=fn: fn(x, c), "check_flag") for c in list(kcols.values()) + [tcol]}})
    return m
