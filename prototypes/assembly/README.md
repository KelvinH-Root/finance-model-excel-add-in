# Assembly proof (Phase 0)

A small, model-agnostic engine that proves the add-in's core idea: insert a module into a model that is already built, and it links itself to the rest of the model. Modano's example model (SEM_COA6) works this way; this proof shows the HFG design does too. It is not the add-in and not the TypeScript engine; the spec's Phase 1 engine follows the same design.

```
python prototypes/assembly/demo.py      # builds build/assembly/base.xlsx, live.xlsx, fresh.xlsx and the chart files
pytest tests/test_assembly.py           # 20 tests; the LibreOffice ones skip without it
```

## What the demo shows

1. Build a model from seven module instances: Financial statements, Checks, two revenue lines, a cost line, Debtors and a debt facility.
2. Open the built workbook, read its metadata (a custom XML part), insert a third revenue line and a second facility.
3. Preview the change in plain words before touching anything:

```
Contents: 2 new rows (rows 13 and 20).
Revenue: 4 new rows (rows 17 to 21).
Working capital: 4 new rows (rows 18 to 22).
Funding: 11 new rows (rows 19 to 30).
Statements: 6 new rows (rows 11, 16, 22, 27, 35 and 38).
Statements: 6 rows rewired (Total revenue, Total interest, Total receipts, Net financing, Total assets, Total debt).
Checks: 1 new row (row 10).
New links: bs.debt x1, bs.debtors x1, cf.financing x1, cf.receipts x1, check.error x1, is.interest x1, is.revenue x2.
```

4. Apply the plan to the open workbook (56 operations: insert rows, write cells, add names), the way the add-in's live writer will through Office.js.
5. Build the same model from scratch and compare. Every formula, value and name matches, there are no formula errors, the balance sheet balances, and the numbers match an independent Python calculation. The same holds after removing a revenue line and a facility.

6. Insert an Income summary into the built model. It is a module that carries a chart, as Modano's summary and report modules do: its sheet, rows and chart arrive together. Then insert a fourth revenue line, and it joins the chart:

```
Dashboard: 1 new row (row 14).
Dashboard: 1 row rewired (Total revenue).
Dashboard: chart Revenue by line, six months re-pointed (4 column series and 1 line, was 3 column series and 1 line).
```

The chart in the live file matches the chart built from scratch: the same series, labels, categories, stacking and anchor cell. Removing a revenue line takes its series out the same way.

Nobody told Debtors about the new revenue line: it mirrors every module that sends `is.revenue`, so it grew a block, and that block's receipts and closing debtors flowed on into the cash flow and balance sheet. That second step is the outward propagation the engine repeats until nothing new appears.

## How it works

| Step | Here |
|---|---|
| 1. Compatibility | `Model.insert` refuses a second copy of a single module or an unknown setting |
| 2. Instance naming | Category modules are numbered (Revenue line 3); numbers are never reused |
| 3. Placement | Each module sits on its area's sheet, areas in `library/areas.yaml` order, instances in insertion order |
| 4. Cells | Formulas are written with structural references (`[row]`, `[row@prev]`, `[range:row]`, `$setting`, `{src}`, `{src_range}`, `[sum:link]`, `{p}` for the period number, `{periods}`) and rendered to A1 references only at the end |
| 5. Links in | A module's `inputs` collect every sender of a link: `each` adds one row per sender, `total` one row for all |
| 6. Outward | `mirror` modules add a block per sender; their outputs feed further modules; resolution repeats to a fixed point |
| 7. Register | Settings become named inputs (`Rev3_Base`), check rows roll up to Checks, the contents lists the module under its sheet and section |
| 8. Records | Every link resolved is kept as a record in the workbook's metadata; required links with no sender and outputs nobody takes are reported |

## Contents, section covers and links

Laid out as Modano's. `library/areas.yaml` groups the areas into sections (Dashboards, Financial Model, Appendices), and the layout is:

- **Contents**, first: sections numbered 1, 2, 3 linking to their covers, sheets lettered a., b., c. within each section, and the modules on each sheet marked "-", then the error and alert totals.
- **A cover sheet before each section**: the section title, "Section N.", a link to the contents, links to the sheets either side, and notes. A cover appears with its section's first sheet and goes with its last.
- **A1 and A2 on every other sheet**: a link to the contents and a link to the checks, which shows a tick while the error checks are clear.
- **Links are HYPERLINK formulas to names** (`HL_Home`, `HL_Err_Chk`, `HL_Sheet_<sheet>`, `HL_Toc_<module>`), and each entry reads the sheet's or module's own label cell (markers `«S|sheet»` and `«B|row»`), so the plan treats the contents like any other rows.

So inserting the Income summary into a built model previews and applies like this, and the result matches a fresh build:

```
Dashboards: new section cover (Dashboards).
Dashboard: new sheet.
Contents: 3 new rows (rows 9 to 11).
Contents: 2 rows rewired (section 2 Financial Model, section 3 Appendices).
Model: 2 rows rewired (section number, link to the previous sheet).
Appendices: 1 row rewired (section number).
```

The live writer stands in for Office.js, which does not style here, so new rows on the contents carry no fonts in the live file; the package writer styles them.

## Charts that come with a module

A module can declare charts. Each series points at rows by key, and `each` gives one series per collected row, so a chart grows and shrinks with the categories feeding it:

```yaml
charts:
  - key: revenue_mix
    title: Revenue by line, six months
    categories: month
    series:
      - {each: rev}               # one stacked column per revenue line
      - {row: total, line: true}
```

The Income summary's chart shows a six-month window that starts at a "first month shown" setting, like Modano's First_Disp_Per drop-downs. The window rows read the full timeline with `INDEX({src_range}, $first+{p}-1)`, so the chart's ranges never move when the timeline is extended; changing the setting moves the window with no add-in involved. A check flags a window that runs past the timeline.

The plan treats charts like rows: `add_chart` when a module with charts arrives, `set_chart` when the rows a chart shows change, `delete_chart` when a module goes (or nothing, if its sheet goes with it). Ranges that only shift because rows were inserted or deleted are left to Excel, which moves them itself. In Office.js these are `sheet.charts.add`, `chart.series.add(...).setValues(range)` with `setXAxisValues`, and `chart.delete()` (ExcelApi 1.7). The live writer stand-in rebuilds the series through LibreOffice's chart API without formatting; the add-in applies the module's chart style as well.

The change plan is the difference between the layout before and after. Rows carry stable ids, so the plan knows which rows are new, which went, and which existing formulas now point at a different set of rows (those are rewritten; everything else is left to Excel's own reference shifting). Inputs someone typed are never overwritten by a structural change.

## Module definition

```yaml
id: demo.debtors
title: Debtors
area: Working capital
mirror: is.revenue           # one block for every module that sends is.revenue
rows:
  - {key: revenue, label: Revenue, unit: "$", formula: "={src}"}
  - {key: receipts, label: Receipts (one month in arrears), unit: "$", first: "=0", formula: "=[revenue@prev]"}
  - {key: closing, label: Closing debtors, unit: "$", total: last, first: "=[revenue]-[receipts]", formula: "=[closing@prev]+[revenue]-[receipts]"}
outputs:
  - {link: cf.receipts, row: receipts}
  - {link: bs.debtors, row: closing}
```

Other keys: `as_category: true` (many instances, each a category for its consumers), `code` (name prefix), `settings`, `inputs` (`link`, `mode` each or total, `required`), `collect` rows (with an optional `key`, `formula` and `span` to make a second, shaped set of rows for the same link), `section` rows, `check: error | alert` rows, `name` on a row, `span` (periods written), `charts`.

## Files

| Path | Holds |
|---|---|
| `assemble.py` | Library, model, link resolution, layout, change plan, package writer, metadata part |
| `live.py` | Applies a plan to an existing workbook through LibreOffice (stand-in for Office.js), and reads snapshots |
| `demo.py` | The demo above |
| `library/` | Demo module definitions with fictional numbers; `assurance.yaml` (framework: assurance) brings the Group assumptions and Input register sheets and the key outputs (see `prototypes/assurance/`) |

Workbooks are written to `build/` and never committed.
