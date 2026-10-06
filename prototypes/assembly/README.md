# Assembly proof (Phase 0)

A small, model-agnostic engine that proves the add-in's core idea: insert a module into a model that is already built, and it links itself to the rest of the model. Modano's example model (SEM_COA6) works this way; this proof shows the HFG design does too. It is not the add-in and not the TypeScript engine; the spec's Phase 1 engine follows the same design.

```
python prototypes/assembly/demo.py      # builds build/assembly/base.xlsx, live.xlsx and fresh.xlsx
pytest tests/test_assembly.py           # 11 tests; the LibreOffice ones skip without it
```

## What the demo shows

1. Build a model from seven module instances: Financial statements, Checks, two revenue lines, a cost line, Debtors and a debt facility.
2. Open the built workbook, read its metadata (a custom XML part), insert a third revenue line and a second facility.
3. Preview the change in plain words before touching anything:

```
Contents: 2 new rows (rows 10 and 14).
Revenue: 4 new rows (rows 17 to 21).
Working capital: 4 new rows (rows 18 to 22).
Funding: 11 new rows (rows 19 to 30).
Statements: 6 new rows (rows 11, 16, 22, 27, 35 and 38).
Statements: 6 rows rewired (Total revenue, Total interest, Total receipts, Net financing, Total assets, Total debt).
Checks: 1 new row (row 10).
New links: bs.debt x1, bs.debtors x1, cf.financing x1, cf.receipts x1, check.error x1, is.interest x1, is.revenue x2.
```

4. Apply the plan to the open workbook (54 operations: insert rows, write cells, add names), the way the add-in's live writer will through Office.js.
5. Build the same model from scratch and compare. Every formula, value and name matches, there are no formula errors, the balance sheet balances, and the numbers match an independent Python calculation. The same holds after removing a revenue line and a facility.

Nobody told Debtors about the new revenue line: it mirrors every module that sends `is.revenue`, so it grew a block, and that block's receipts and closing debtors flowed on into the cash flow and balance sheet. That second step is the outward propagation the engine repeats until nothing new appears.

## How it works

| Step | Here |
|---|---|
| 1. Compatibility | `Model.insert` refuses a second copy of a single module or an unknown setting |
| 2. Instance naming | Category modules are numbered (Revenue line 3); numbers are never reused |
| 3. Placement | Each module sits on its area's sheet, areas in `library/areas.yaml` order, instances in insertion order |
| 4. Cells | Formulas are written with structural references (`[row]`, `[row@prev]`, `$setting`, `{src}`, `[sum:link]`) and rendered to A1 references only at the end |
| 5. Links in | A module's `inputs` collect every sender of a link: `each` adds one row per sender, `total` one row for all |
| 6. Outward | `mirror` modules add a block per sender; their outputs feed further modules; resolution repeats to a fixed point |
| 7. Register | Settings become named inputs (`Rev3_Base`), check rows roll up to Checks, the Contents sheet lists every module |
| 8. Records | Every link resolved is kept as a record in the workbook's metadata; required links with no sender and outputs nobody takes are reported |

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

Other keys: `as_category: true` (many instances, each a category for its consumers), `code` (name prefix), `settings`, `inputs` (`link`, `mode` each or total, `required`), `collect` rows, `section` rows, `check: error | alert` rows, `name` on a row.

## Files

| Path | Holds |
|---|---|
| `assemble.py` | Library, model, link resolution, layout, change plan, package writer, metadata part |
| `live.py` | Applies a plan to an existing workbook through LibreOffice (stand-in for Office.js), and reads snapshots |
| `demo.py` | The demo above |
| `library/` | Demo module definitions with fictional numbers |

Workbooks are written to `build/` and never committed.
