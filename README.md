# hfg-models

Home of the HFG Excel add-in: a Modano-style tool that builds and maintains HFG's Excel models from a shared library of modules, without being tied to any one model. The repo holds the frame standard, the link dictionary, the module library, the engine and its tests.

Owner: Kelvin Herbst (Group Finance Manager). Build team: Kelvin, Claude and Codex. Read `collaboration.md` before starting any work.

## Where things stand (6 October 2026)

- **Goal:** parity with Modano's add-in, built as an internal tool and not tied to any one model.
- **Requirements spec (Phase 0):** a Claude Doc in Kelvin's Technical Roadmap project, "HFG Excel add-in: requirements spec". Its tabs hold the spec itself, a parity matrix of 143 Modano features from its user guide and content libraries (none known to be blocked; 13 met a different way) plus three HFG additions, a catalogue of 19 HFG model types built from 50 modules, and a register of NZ rules the modules must handle. No add-in code until it is agreed.
- **Probe (Phase 0):** `addin-probe/` is a throwaway add-in that tests the Office.js behaviour the spec could not settle from documentation. Kelvin runs it on Windows, Mac and the web and sends back the results sheet.
- **Proven so far:** editing the template at package level (form controls, charts and Python in Excel parts survive), a row layout engine that links modules by key across sheets, and a harness that recalculates in LibreOffice and compares every row with a reference calculation.
- **Test example:** `examples/development/` is a fictional multi-site development model used only to prove the engine. It is not a product and not HCP's model.

## What is here

| Path | What it holds |
|---|---|
| `hfgmodels/xlsx/` | Package-level template editing, worksheet XML writer, style book |
| `hfgmodels/frame.py` | The frame standard in code: columns, header rows, timeline block, template cell formats |
| `hfgmodels/layout.py` | Row layout engine: rows declared by key, formulas that refer to other rows by key |
| `hfgmodels/verify.py` | Headless LibreOffice recalculation |
| `examples/development/` | Test example: recipe, sheet definitions, reference calculation, comparison |
| `addin-probe/` | Phase 0 probe add-in: manifest, task pane, probes, Node tests |
| `library/links.yaml` | Link dictionary, first cut |
| `docs/frame-standard.md` | The frame standard |
| `tests/` | Package editing, reference properties, and the full example build compared with its reference |

## Rules that keep this repository safe

- No workbooks are committed. The template, Modano's example files and every HCP model stay in Kelvin's files. `.gitignore` blocks Excel files and `build/`.
- No real financial data is committed. Examples use fictional data.
- Clean room: modules are HFG's own. Nothing is copied from Modano's library, formula templates, names or icons.
- Nothing in `hfgmodels/` is specific to one model. Model-specific logic lives in examples or, later, in library modules.

## Run the test example

Requirements: Python 3.11 or later; LibreOffice for recalculation.

```
pip install -e ".[test]"
export HFG_TEMPLATE=/path/to/Budget_Template.xlsx      # PowerShell: $env:HFG_TEMPLATE="C:\path\Budget_Template.xlsx"
python -m hfgmodels build-example --verify
pytest -q
```

Tests that need the template skip themselves when `HFG_TEMPLATE` is not set.

The probe has its own instructions in `addin-probe/README.md` (`npm install`, `npm run certs`, `npm start`; `npm test` for its Node tests).
