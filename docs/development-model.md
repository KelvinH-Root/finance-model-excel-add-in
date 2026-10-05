# HCP development model: design (version 0.1, 5 October 2026)

A monthly, multi-site development finance model built into Kelvin's template. It carries the logic of HCP's current models, fixes the faults found in them, and is written so every calculated row is reproduced and tested by `hfgmodels/dev/reference.py`.

## Sheets

| Sheet | Section | Holds |
|---|---|---|
| Dev_Sites | Financial Model | Model-wide inputs, then one input block per site: general, opening position, cost budget by line, sale and exit, funding, accounting, derived values |
| Dev_Costs | Financial Model | Phasing share per line, forecast cost by line, GST on costs, costs paid, cumulative cost |
| Dev_Sales | Financial Model | Settlement, revenue, output GST, proceeds; the exit through the SPV to the Fund |
| Dev_GST | Financial Model | Net GST, settlement with IRD after a lag, GST balance |
| Dev_Funding | Financial Model | Debt ceiling, finance costs, draws and repayments, equity, cash held for GST, ratios, cash flows for IRR |
| Dev_FS | Financial Model | WIP, income statement, balance sheet, cash flow statement, balance and cash checks |
| Dev_Returns | Financial Model | Returns by site and portfolio |
| Dev_Checks | Appendices | Every check in plain English, by site, with error and alert totals |

## Logic

Notation: month t, last actual month L (`DD_Ts_Last_Hist_Mth`). Forecast months are t > L. Amounts exclude GST unless stated.

**Cost phasing.** Each site has nine standard lines (land, due diligence and consents, development contributions, civils, vertical construction, contingency, professional fees, GP fee, sales and marketing). Each line has a budget, a profile (lump sum, flat, S-curve), a start date, a duration and a GST flag. The S-curve share in month k of n is `BETA.DIST(k/n, a, b) - BETA.DIST((k-1)/n, a, b)`; a = b gives a symmetric curve. The phased budget in forecast months is scaled so that opening WIP plus forecast cost equals the budget (the "cost to complete scaling factor"). That keeps any underspend before the cut-over in the forecast instead of dropping it.

**GST.** Input GST on taxable lines. A registered entity nets output against input GST and settles with IRD after a lag (default one month); an unregistered entity capitalises input GST into WIP. Land defaults to zero-rated. The sale is standard-rated by default because the buyer (SPV or NZ Housing Fund) may not be registered; it can be switched to zero-rated per site.

**Sale and exit.** Each site settles on a settlement date at its As Complete (BTS) value including GST. Revenue is the price less output GST. When "sold to an SPV" is Yes, the SPV pays the BTS value and sells to the Fund at the BTR value: the difference is the margin to HHLP, and the output GST on the internal sale is shown as a group cost. Te Piringa's debt and equity on the purchase use the fund LVR.

**Funding.** In each forecast month:
- Debt ceiling: nil once sold; from the facility start date, the lower of the limit and either the maximum LTC times cumulative cost ("LTC ceiling" method) or cumulative cost less the committed equity ("Equity first" method); before the facility start, the pre-development (land) loan.
- Interest on the opening balance at the all-in rate / 12; line fee on the limit from facility start until sale; establishment fee on the limit in the start month.
- Finance costs are capitalised only within the headroom under the ceiling; anything above is paid in cash by equity.
- The facility moves to the ceiling. With "release equity" on, it draws to the ceiling even when that returns equity (the catch-up draw at facility start). With it off, it draws only what is needed.
- Equity is the plug: cash needed (costs paid, GST paid less refunds, finance costs paid, less sale proceeds) less the draw. Cash equal to any GST payable is held back so output GST is paid from the sale proceeds rather than distributed and called back.
- Limit basis: override, LTC on budget, LVR on As Complete value, or LVR on BTR value.

**Statements.** WIP builds from costs, unrecoverable GST and (by policy switch) capitalised borrowing costs, and is released to cost of sales at settlement. The balance sheet (cash, WIP, facility, GST, partners' capital, retained surplus) balances every month; the cash flow statement reconciles to the change in cash.

**Returns.** Revenue, development cost, finance costs, surplus, margin on revenue, surplus on cost, peak debt and month, peak equity, peak LTC, equity IRR and project IRR (monthly, annualised, from the first forecast month with the opening position as the first flow), equity multiple, HHLP margin, group margin, GST on the internal sale, Te Piringa purchase, debt and equity.

## How HCP's current models work, and what changes

Reviewed 5 October 2026: Development Funding Model v4 (the newest), DevCo Model to CF (March 2026), DevCo Drawdown Forecast (September 2025), Funding Summary (April 2026), a buy and hold feasibility, and an AI-generated portfolio template. Detailed reviews are kept out of git because they contain real figures.

| HCP model behaviour | Version 0.1 |
|---|---|
| Costs are pasted per project from an upstream cashflow tool (profiles: manual, S-curve, flat, lump sum; escalation factors present but unused) | Phasing happens in the model from line budgets, profiles and dates. Escalation is not yet modelled |
| Opening WIP and debt at an as-at date; budget scheduled before the as-at date but not in WIP is dropped | Opening position at the last actual month; cost to complete is rescaled so budget = WIP + forecast (checked) |
| Sale in the month of the last cost, at As Complete / 1.15; source revenue and timing ignored | Explicit settlement date and price per site; GST treatment by switch |
| Two-step exit LP to SPV (BTS incl GST) to Fund (BTR, no GST); margin to HHLP; GST on the internal sale is a group cost | Same, by switch per site |
| Debt ceiling min(limit, 70% x cost to date) from the debt start date; land loan held flat before; debt and equity draw together; catch-up draw releases equity | Same as the "LTC ceiling" method with "release equity" on. "Equity first" is also available |
| Facility limit selector does not work (always uses BTR x LVR) | Working selector with four bases and an override |
| Capitalised interest can push debt above the limit | Capitalised only within the ceiling; the rest is paid by equity; checked |
| Upfront fee never fires | Establishment fee fires in the facility start month; line fee available |
| Sale-month surpluses double counted in equity and cash | One waterfall; cash only holds GST payable; checked |
| No GST on costs or refund timing | Input GST, settlement lag, unrecoverable GST for unregistered entities |
| No IRR, no checks | Equity and project IRR, multiples, and 14 checks |
| 3D sums across bookend sheets | Portfolio block on each sheet that adds sites by an include flag; checked against the sites |

## Not yet in version 0.1

- Te Paeroa group revolver (a group facility carrying other sites' land loans, drawn on need and repaid from released equity).
- Actuals feed from Home Hub for months up to the cut-over, and the Home Hub output contracts (`hh_out_plan`, `hh_out_intergroup`).
- Cost escalation, staged settlements, retentions, multiple tranches per facility.
- Scenario manager links (base, best, worst) and the Monte Carlo layer in Python in Excel.
- Wiring the template's Income, Balance and Cash summary pages to the development statements, and removing the template sheets the development model does not use (Syft PL group, Cash_Fcst).
- Hold period with IRRS income, which belongs in the NZ Housing Fund model (the next Excel model).

## Questions for Kelvin

See `collaboration.md` for the list raised on 5 October 2026.
