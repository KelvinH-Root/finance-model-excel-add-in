# Charts and controls proof (Phase 0)

Shows what the Build tab's Charts and Controls commands will produce: native Excel charts and controls driven by formulas, written by the package writer. No pictures, no Python, no add-in needed to calculate or redraw.

```
python prototypes/charts/build.py          # build/charts/charts_demo.xlsx
pytest tests/test_charts.py                # 3 tests; the LibreOffice one skips without it
```

## What is in the workbook

| Sheet | Shows |
|---|---|
| Contents | Links to every sheet and the checks total |
| Data | Prior year, actual then forecast, and budget by month for three lines; a classic combo box (form control) that picks the line for the Z chart, linked to `DD_Chart_Line` with its list in `LU_Chart_Lines`; the last actual month |
| Z chart | The template's Z-Chart design, live: actual bars black, forecast hatched, budget grey behind and to the left; year to date actual then forecast (dashed), year to date budget and moving annual total, labelled on the lines instead of a legend, with values each quarter; a divider between actual and forecast, AC and FC under the months, and year to date and full year variance boxes on the right (green good, red bad, aware of cost lines) |
| Waterfall | P&L walk and budget to actual, built as stacked columns: totals grey, increases green, decreases red, signed labels; the colours follow the numbers |
| IBCS | Prior year grey, plan outlined, actual solid, forecast hatched; variance bars green when good and red when bad, aware of cost lines; the comparison (plan or prior year) chosen with an in-cell drop-down |
| Lookups | The lists behind both drop-downs, and which lines are costs (a fall against budget is good) |
| Checks | Selections valid, the Z chart ties to the data and its MAT meets the full year, both bridges close |

The Z chart reproduces the template's Z-Chart picture element by element, and its numbers: actuals and forecast as typed there, and a budget that passes through every labelled point of the template's budget line (6, 27, 38, 46, 66, 76). Two of the template's typed numbers cannot come from live formulas, so they differ on purpose:

- its budget bars add to 74 while its budget line ends at 76; here the bars add up to the line;
- its moving annual total ends at 55 while the year totals 82; a moving annual total at the year end is the year's total, so here the two lines meet at March, which is what gives a Z chart its Z.

Labels round half up as Excel does, so a value of 60.5 shows as 61 where the picture showed 60.

## How it is built

- Charts are openpyxl chart objects with series fills, pattern fills, dashes, secondary axes and label formats set in the chart XML. Office.js can create most of this live (the probe's chart tests check which parts), but not pattern fills, so live inserts outline forecast bars instead of hatching them.
- The Z chart's budget bars sit on the primary axes with an empty second series (overlap 50%, gap 17%), which makes them 0.6 of a month wide and moves them 0.15 of a month to the left; actual and forecast bars sit on the secondary axes (overlap 100%, gap 100%), 0.5 of a month wide and centred, and Excel draws them in front. Hidden series with the same maximum keep both axes on one scale.
- Everything placed between or beyond the months (the divider, AC and FC, the dotted reference lines, the variance boxes and their labels) is a scatter series on the category axis, where x = 1 is the first month and 6.5 falls between September and October. The helper table on the Z chart sheet works out each position from the last actual month, and turns a helper to `#N/A` when it should not show (no divider with a full year of actuals, no YTD box before the first actual month, the red box only when a variance is bad).
- Labels that are words (MAT, AC+FC, BUD, AC, FC, YTD Δ, +6) are series names linked to cells, shown as data labels, so they change with the numbers. The variance boxes are thick lines with flat ends sized for the chart's width, with a thin outline drawn after them.
- Every chart has Excel's "Show #N/A as an empty cell" switched on, so `#N/A` draws nothing and shows no label.
- Waterfalls use a hidden base series rather than Excel's built-in waterfall type, because Office.js cannot mark totals on the built-in type.
- `#N/A` in the chart helper rows is deliberate: it stops a line or bar being drawn for that month.
- The combo box is written as Excel writes one: a control properties part, a VML shape, a hidden drawing shape and the sheet's controls element. Office.js cannot create form controls, so the package writer adds them when a model is built or rebuilt; commands that run inside an open workbook use in-cell drop-downs and checkboxes on the same `DD_` and `CB_` names.

## Not yet verified in Excel

LibreOffice renders every chart and calculates every formula, and the tests change the selections, the last actual month and the line (including a cost line) and check the results. LibreOffice draws three things differently from Excel: it puts bars from two axes side by side instead of overlapping them, it does not follow a chart title linked to a cell, and it draws later scatter series underneath earlier ones. Excel itself should open the file cleanly, show the combo box, overlap the bars and draw the hatching; open `build/charts/charts_demo.xlsx` in Excel desktop and on the web to confirm. If the chart is resized, the variance boxes keep their width in points; Refresh charts will reset them.
