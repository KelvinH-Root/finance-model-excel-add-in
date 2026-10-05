# hfg-models

Home Foundation Group's model library: one frame standard, a link dictionary, reusable modules and a Python build script that assembles Excel models into Kelvin's preferred template. The HCP development finance model is the first model built with it.

Owner: Kelvin Herbst (Group Finance Manager). Build team: Kelvin, Claude and Codex. See `collaboration.md` before starting any work.

## What is here

| Path | What it holds |
|---|---|
| `hfgmodels/xlsx/` | Package-level editing of the template (adds sheets, names and styles without disturbing form controls, charts or Python in Excel parts), a worksheet XML writer and a style book |
| `hfgmodels/frame.py` | The frame standard in code: column layout, header rows, timeline block, cell formats taken from the template |
| `hfgmodels/layout.py` | Row layout engine: modules declare rows by key and refer to each other by key across sheets |
| `hfgmodels/dev/` | The development model: recipe loader, sheet definitions, reference calculation, assembler, comparison |
| `hfgmodels/verify.py` | Headless LibreOffice recalculation used by tests and for cached values |
| `models/hcp-development/recipe.yaml` | Build recipe for the HCP development model, with fictional demo sites |
| `library/links.yaml` | Link dictionary (first cut): the named items modules and Home Hub exchange |
| `docs/` | Frame standard, development model design, link dictionary |
| `tests/` | Reference properties, package editing, and a full build compared with the reference |

## Rules that keep this repository safe

- No workbooks are committed. The template (`Budget_Template.xlsx`), Modano's example files and every HCP model stay in Kelvin's files. `.gitignore` blocks `*.xlsx`, `*.xlsm`, `*.xlsb` and `build/`.
- No real financial data is committed. Recipes carry fictional demo data. Real inputs come from Home Hub (later) or are typed into the built workbook.
- Clean room: HFG modules are written here from first principles. Nothing is copied from Modano's library, formula templates, names or icons.

## Build the development model

Requirements: Python 3.11 or later, LibreOffice (for recalculation and tests).

```
pip install -e ".[test]"
export HFG_TEMPLATE=/path/to/Budget_Template.xlsx      # Windows PowerShell: $env:HFG_TEMPLATE="C:\path\Budget_Template.xlsx"
python -m hfgmodels build-dev --verify
```

The workbook lands in `build/hcp_development_model.xlsx`. `--verify` recalculates it in LibreOffice and compares every calculated row with the reference calculation in `hfgmodels/dev/reference.py`. Use `--no-values` to skip LibreOffice; Excel recalculates on open either way.

## Tests

```
pytest -q
```

Tests that need the template skip themselves when `HFG_TEMPLATE` is not set, so CI runs the rest.

## Status (5 October 2026)

Version 0.1 of the development model: multi-site inputs, cost phasing (S-curve, flat, lump sum) with cost to complete scaled from an opening WIP, GST with a settlement lag and a registration switch, the two-step exit (LP to SPV to Fund) with the margin to HHLP, facility funding with an LTC ceiling or equity first, capitalised interest within the ceiling, statements that balance, returns and a checks sheet. See `docs/development-model.md` for the design, how it maps to HCP's current models and what comes next.
