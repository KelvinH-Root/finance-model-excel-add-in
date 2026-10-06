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
| Z chart | Monthly bars (actual solid, forecast hatched, budget grey behind), year to date actual then forecast (dashed), year to date budget, moving annual total, and the year to date and full year variances |
| Waterfall | P&L walk and budget to actual, built as stacked columns: totals grey, increases green, decreases red, signed labels; the colours follow the numbers |
| IBCS | Prior year grey, plan outlined, actual solid, forecast hatched; variance bars green when good and red when bad, aware of cost lines; the comparison (plan or prior year) chosen with an in-cell drop-down |
| Lookups | The lists behind both drop-downs |
| Checks | Selections valid, the Z chart ties to the data, both bridges close |

The Z chart reproduces the numbers in the template's Z-Chart sheet (year to date 44 against 38 budget at September; full year 82 against 76), which today is a picture drawn by Python from typed values.

## How it is built

- Charts are openpyxl chart objects with series fills, pattern fills, dashes, secondary axes and label formats set in the chart XML. Office.js can create most of this live (the probe's chart tests check which parts), but not pattern fills, so live inserts outline forecast bars instead of hatching them.
- The Z chart puts budget bars on one axis group and actual and forecast bars on another so they overlap like the template's picture; a hidden series with the same maximum keeps both axes on one scale.
- Waterfalls use a hidden base series rather than Excel's built-in waterfall type, because Office.js cannot mark totals on the built-in type.
- `#N/A` in the chart helper rows is deliberate: it stops a line or bar being drawn for that month.
- The combo box is written as Excel writes one: a control properties part, a VML shape, a hidden drawing shape and the sheet's controls element. Office.js cannot create form controls, so the package writer adds them when a model is built or rebuilt; commands that run inside an open workbook use in-cell drop-downs and checkboxes on the same `DD_` and `CB_` names.

## Not yet verified in Excel

LibreOffice renders every chart and calculates every formula, and the tests change both selections and check the results. Excel itself should open the file cleanly, show the combo box and draw the hatching; open `build/charts/charts_demo.xlsx` in Excel desktop and on the web to confirm.
