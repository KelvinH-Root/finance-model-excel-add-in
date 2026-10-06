# Report charts proof

Phase 0 proof that the summary and report modules can bring every chart with them, as native charts on a live model. Kelvin asked for all of the charts in Modano's example model; this rebuilds all 95 in HFG terms.

    python prototypes/reports/build.py [out.xlsx]

The default output is `build/reports/reports_demo.xlsx` (not in git). Numbers are fictional: a small building and maintenance business over four financial years (FY2025 to FY2028), actuals to September 2026 and forecast after, with a budget and three scenarios.

## What is in the workbook

| Sheet | What it holds |
|---|---|
| Contents | Links to every sheet, error and alert totals |
| Income summary, Balance summary, Cash summary, Budget summary | The four summary modules (5, 6, 5 and 6 charts) |
| Income report, Balance report, Cash report, Budget report, Scenario report | The five report modules (28, 7, 12, 6 and 20 charts) |
| Time | Timeline, last actual month (`Ts_Last_Actual`) |
| Assumptions | Tax rate, working capital days, opening balance sheet |
| Scenarios | Scenario factors (forecast months only) and the active scenario (`DD_Scenario`) |
| Inputs | Monthly inputs by category and the budget |
| Statements | Income statement, balance sheet and cash flow for the active scenario, the budget and each scenario |
| Chart register | Every chart with its recipe, what it reads and a link to its rows |
| Lookups | Lists behind the drop-downs |
| Checks | Error checks (must be nil) and alerts |

Each report sheet has its selections at the top (year shown, month shown, as in-cell drop-downs), a grid of charts, and below the grid the rows each chart reads. Every number in those rows is a formula on the statements (`INDEX` from the selection), so changing the year, the month, the active scenario or the last actual month redraws every chart without the add-in.

## The register

`register.yaml` lists the 95 charts (C01 to C95): module, title, recipe and what the chart reads. The builder expands each recipe into formula rows and a native chart:

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

`tests/test_reports.py` (5 tests): the register holds C01 to C95 once with the module counts above; every chart in the workbook is named by its register id, reads only its own module's rows, has the expected series and chart type and shows #N/A as empty; with LibreOffice, the statements match the Python reference, there are no formula errors and no checks raised, the chart rows match the reference (ranking, bridges, movement, budget, scenarios, rolling windows, year to date), and moving the year, month and scenario moves every chart (the first year leaves the year-before series blank as #N/A and raises one alert, not an error).

## Not proven here

- Excel itself. Rendered and recalculated in LibreOffice only; the hatched forecast bars, label formats and the "Show #N/A as an empty cell" setting need Excel desktop and the web to confirm.
- Inserting these modules live. `prototypes/assembly` proves that a module's chart arrives with it and follows category changes; this proof shows the full chart set and recipes. The engine combines the two.
