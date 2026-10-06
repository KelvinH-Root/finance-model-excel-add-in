# Report charts proof

Phase 0 proof that the summary and report modules can bring every chart with them, as native charts on a live model. Kelvin asked for all of the charts in Modano's example model; this rebuilds all 95 in HFG terms. It also proves saved versions: approved budgets and a reforecast for every month kept as values, and budget reports that compare against them (four more charts, C96 to C99).

    python prototypes/reports/build.py [out.xlsx]

The default output is `build/reports/reports_demo.xlsx` (not in git). Numbers are fictional: a small building and maintenance business over four financial years (FY2025 to FY2028), actuals to September 2026 and forecast after, with a budget and three scenarios.

## What is in the workbook

| Sheet | What it holds |
|---|---|
| Contents | Numbered contents (sections, lettered sheets, headings), links to every sheet, error and alert totals |
| 1 Dashboards (cover), Income summary, Balance summary, Cash summary, Budget summary, Version comparison | The four summary modules (5, 6, 5 and 6 charts) and the Version comparison (4 charts); Phase 1 |
| 2 Model (cover): Time, Assumptions, Scenarios, Seasonality, Inputs, Statements, Versions | Timeline, rates and opening balances, scenario factors, the revenue budget phased by seasonality, monthly inputs, the statements for the active scenario, the budget being built and each scenario, and the register of saved versions |
| 3 Reports (cover): Income report, Balance report, Cash report, Budget report, Scenario report | The five report modules (28, 7, 12, 6 and 20 charts); Phase 2 |
| 4 Appendices (cover): Chart register, Version store, Lookups, Checks | Every chart with its recipe and a link to its rows; every saved version's values; lists behind the drop-downs; error checks (must be nil) and alerts |

This is also the budget and actuals example model: every sheet links to the contents (A1) and the checks (A2), and each section has a cover sheet, written by `prototypes/models/navigation.py`.

## Seasonality

The revenue budget is phased the way the template's Seasonality sheet does it, as a module: each month's share of revenue in FY2025 and FY2026 (actual years), with an Include flag per year so an unusual year can be left out, a typed override per month, and an even spread when no history is included (an alert). The annual budget for each year times the month's share gives the monthly revenue budget; cost of sales follows at the budget ratio on Assumptions. Checks: the profile adds to 100% and the phased budget adds to the annual budgets.

## Saved versions: budgets and monthly reforecasts

Kelvin asked for a budget that is built and then saved, as Modano does, and for monthly reforecasts as the comparisons in budget versus actual models. Modano's example keeps one budget as typed values in a budget module (income statement only) and treats the live model as the reforecast, so last month's reforecast is gone once the model rolls forward. Here every saved version is kept (`versions.py`):

- **Versions** (register): one row per version: type (Budget, Reforecast or Other), label ("Budget FY2027", "Reforecast Sep 2026 (6+6)"), the year a budget is for, the month a reforecast is as at, status (an approved budget per year), lock, source, saved on, where its rows sit in the store, and a checksum taken when saved (each value's size times its column, so a changed or moved value shows).
- **Version store**: the values, one row per version and line, keyed `id|line`, for every month. A forecast version keeps every statement line (income statement, balance sheet, cash flow, categories); a budget keeps the income statement it was built with.
- **Compared with** on the Budget summary and Budget report: the approved budget for the year shown, last month's reforecast (the latest saved before the last actual month), the latest reforecast, the budget being built (not saved), or any saved version by name. Chart titles and legends say which.
- **Version comparison** (Dashboards): the month, year to date and full year against two comparisons for the 13 income statement lines, favourable when positive (for costs, spending less); and four charts: full-year outturn by version (the approved budget, each reforecast of the year, the current forecast), actual and forecast by month against both comparisons, what the budget and each of the 12 reforecasts before it expected for the month shown, and a walk from the comparison's profit after tax to the outturn.
- **Checks**: a stored value has changed since it was saved (error); two versions answer the same question, such as two approved budgets for a year (error); a comparison is not a listed choice (error); nothing is saved for a comparison (alert, blank); this month's reforecast is not saved yet (alert).

The demo's history is two approved budgets (FY2026, phased on FY2025 alone; FY2027) and 18 reforecasts, April 2025 to September 2026, each replaying the month's actuals and that month's forecast (fictional revisions on the trend). The package writer writes them from the reference; `save_live` is Save version through LibreOffice, as the add-in will do it through Office.js.

Each report sheet has its selections at the top (year shown, month shown, as in-cell drop-downs), a grid of charts, and below the grid the rows each chart reads. Every number in those rows is a formula on the statements (`INDEX` from the selection), so changing the year, the month, the active scenario or the last actual month redraws every chart without the add-in.

## The register

`register.yaml` lists the 99 charts (C01 to C95 from Modano's example, C96 to C99 the Version comparison): module, title, recipe and what the chart reads. The builder expands each recipe into formula rows and a native chart:

| Recipe | Chart |
|---|---|
| compare | One line over the year before, the year shown and the year after (or the 12 months before, to and after the month shown), as lines or columns, optionally cumulative |
| mix | A group stacked by month, with the group total for the periods before and after as lines |
| depth | A group's totals for the period, one column per member |
| pie | A group's make-up for the period |
| combo | Chosen lines as columns (clustered or stacked) with lines drawn over them |
| budget | Actual (solid) and forecast (hatched) by month, budget as a line |
| scenario | The line in each scenario, by month of the year shown or by year |
| bridge | A waterfall: net assets built up from the balance sheet, or opening to closing cash |
| movement | Change in each balance sheet line, or the totals, against a year earlier |

Groups such as revenue lines are category sets: an "each" series grows and shrinks with the model's categories, and `top: N` ranks members by the period's total with formulas (`LARGE` and `MATCH`, ties broken by order) and puts the rest in Other. Charts are named by their register id, so the add-in can find a module's charts when it inserts, replaces or removes the module.

## Mapping to Modano's example

The requirements spec's Chart register tab maps each HFG chart to the Modano chart it replaces. Kept the same: the chart types (pies stay pies, waterfalls stay waterfalls), the nine modules and how many charts each brings. Changed on purpose:

- Comparison periods line up with the months shown: the year before and after the year shown, or the 12 months before and after the month shown. Modano mixed "this financial year, last 12 months, next 12 months" on one axis, which only lines up when the month shown is a year end.
- The Scenario report's summary set shows each scenario by year (four points per line) rather than repeating the monthly set.
- Comparisons follow IBCS (prior grey, actual solid, forecast hatched, budget a line) and categories use the HF chart palette.
- Titles and labels are HFG's own wording.

## Tests

`tests/test_reports.py` (7 tests): the register holds C01 to C99 once with the module counts above; every chart in the workbook is named by its register id, reads only its own module's rows, has the expected series and chart type and shows #N/A as empty; with LibreOffice, the statements match the Python reference, there are no formula errors and no checks raised, the chart rows match the reference (ranking, bridges, movement, budget, scenarios, rolling windows, year to date), and moving the year, month and scenario moves every chart (the first year leaves the year-before series blank as #N/A and raises one alert, not an error); the contents, covers and A2 ticks read right; and seasonality phases the budget, leaving a year out changes the budget being built but not the saved budget, and a broken override raises the checks.

`tests/test_versions.py` (5 tests): the history is two budgets and a reforecast every month, the latest equal to the model as it stands; with LibreOffice, the store holds every version as saved with matching checksums; each Compared with choice reads the right version into the budget charts (a year with no approved budget is blank with an alert, a choice not in the list is an error); the Version comparison's month, year to date, full year, outturn by version, expectations for the month and walk match the reference, and the line and month shown move them; and Save version live: roll forward to October and save its reforecast (the alert clears and the stored values match the reference), saved values hold when an input changes, editing the store is caught by the checksum, a budget saved and then revised supersedes the first, and replaying August 2026 live stores exactly what the package writer did.

## Not proven here

- Excel itself. Rendered and recalculated in LibreOffice only; the hatched forecast bars, label formats and the "Show #N/A as an empty cell" setting need Excel desktop and the web to confirm.
- Inserting these modules live. `prototypes/assembly` proves that a module's chart arrives with it and follows category changes; this proof shows the full chart set and recipes. The engine combines the two.
