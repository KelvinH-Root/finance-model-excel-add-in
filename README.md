# hfg-models

Home of the HFG Excel add-in: a Modano-style tool that builds and maintains HFG's Excel models from a shared library of modules, without being tied to any one model. The repo holds the frame standard, the link dictionary, the module library, the engine and its tests.

Owner: Kelvin Herbst (Group Finance Manager). Build team: Kelvin, Claude and Codex. Read `collaboration.md` before starting any work.

## Where things stand (6 October 2026)

- **Goal:** parity with Modano's add-in, built as an internal tool and not tied to any one model.
- **Requirements spec (Phase 0):** a Claude Doc in Kelvin's Technical Roadmap project, "HFG Excel add-in: requirements spec". Its tabs hold the spec itself, a parity matrix of 143 Modano features from its user guide and content libraries (none known to be blocked; 10 met a different way) plus four HFG additions, a catalogue of 19 HFG model types built from 55 modules and shaped on professional models, and a register of NZ rules the modules must handle. No add-in code until it is agreed.
- **Probe (Phase 0):** `addin-probe/` is a throwaway add-in that tests the Office.js behaviour the spec could not settle from documentation, and carries the designed ribbon (the HFG Model tab and a runtime HFG Build tab) so Kelvin can click through it. Kelvin runs it on Windows, Mac and the web and sends back the results sheet.
- **Proven so far:** editing the template at package level (form controls, charts and Python in Excel parts survive), a row layout engine that links modules by key across sheets, and a harness that recalculates in LibreOffice and compares every row with a reference calculation.
- **Phase 0 proofs:** `prototypes/assembly/` inserts modules into a built workbook and they link themselves in, charts included (the same result as a fresh build, cell for cell and series for series); `prototypes/montecarlo/` is a native, seeded Monte Carlo over a one-way cash flow model with scenario events, matching its Python reference trial by trial; `prototypes/charts/` writes native, formula-driven Z, waterfall and IBCS charts and a classic combo box, with the Z chart matching the template's Z-Chart design; `prototypes/reports/` rebuilds all 95 charts that come with Modano's summary and report modules as native charts on a live three-statement model, each from a recipe in a chart register, and is also the budget and actuals example, with saved budgets and monthly reforecasts that the budget reports compare against; `prototypes/models/` writes Modano-style contents, section covers and links into a built workbook and finishes a development example from the template, with screenshots of both; `prototypes/consolidation/` consolidates a fictional group shaped like HFG's (five groups, intergroup margin into WIP and property, at-cost on-charges through netting accounts, costs moved into a development LP's WIP, a portfolio sale, NCI at nodes) and matches an independent group-view reference in every group and year; `prototypes/impacts/` proves the Impacts command: the live impact of changing an input (change, recalculate, read, restore) and Impacts sheets for each kind of transaction a model holds; `prototypes/assurance/` proves model assurance: key outputs checked before and after every structural command with a change log, the input register, group assumptions, model compare, release profiles and GST due dates. All are model-agnostic proofs for the spec, not the add-in.
- **Phase 1 engine (started 7 October 2026):** `engine/` is the engine the add-in will run, in TypeScript. Its core (library, link resolution, layout, rendering, change plan) reproduces the assembly proof exactly on eleven scenarios. HFG's standard frame (the Look and wiring standard: Settings and timeline block, contents with logo and status line, covers, cell hyperlinks, spacing, styles on each entity's theme) and the package writer followed the same day, and the New model wizard (`addin/`) builds a model in it from the probe's ribbon.
- **Test example:** `examples/development/` is a fictional multi-site development model used only to prove the engine. It is not a product and not HCP's model.

## What is here

| Path | What it holds |
|---|---|
| `hfgmodels/xlsx/` | Package-level template editing, worksheet XML writer, style book |
| `hfgmodels/frame.py` | The frame standard in code: columns, header rows, timeline block, template cell formats |
| `hfgmodels/layout.py` | Row layout engine: rows declared by key, formulas that refer to other rows by key |
| `hfgmodels/verify.py` | Headless LibreOffice recalculation |
| `examples/development/` | Test example: recipe, sheet definitions, reference calculation, comparison |
| `engine/` | Phase 1 engine in TypeScript: module library, link resolution, layout, rendering and the change plan, tested cell for cell against the assembly proof |
| `addin/` | Phase 1 add-in: the New model wizard, Insert module and the live writer (built into the probe as `addin.bundle.js` until the add-in has its own manifest) |
| `addin-probe/` | Phase 0 probe add-in: manifest, task pane, probes, Node tests |
| `prototypes/assembly/` | Phase 0 proof: module definitions, link resolution to a fixed point, change plans, live apply, charts that come with a module |
| `prototypes/montecarlo/` | Phase 0 proof: seeded Monte Carlo with one data table, scenario events, self-checking workbook |
| `prototypes/charts/` | Phase 0 proof: native Z chart, waterfalls and IBCS charts driven by formulas, a form control and an in-cell drop-down |
| `prototypes/reports/` | Phase 0 proof: the nine summary and report modules with all 95 of their charts, from a chart register and nine recipes; saved versions (budgets and monthly reforecasts) and the Version comparison; the budget and actuals example |
| `prototypes/models/` | Phase 0 proof: contents, section covers and links as Modano lays them out, kept current; the development example and screenshots of both examples |
| `prototypes/consolidation/` | Phase 0 proof: group consolidation from entity trial balances and the intercompany register; eliminations by type, unrealised margin by group and site, investments, NCI at nodes, statements by group, checks |
| `prototypes/impacts/` | Phase 0 proof: the Impacts command's live round trip on an assembled model, with the link chain, and Impacts sheets drawn from a model's own accounts and entities |
| `prototypes/assurance/` | Phase 0 proof: key outputs, the change log and the before-and-after check; input register and group assumptions; model compare; release profiles; GST return periods and due dates |
| `library/links.yaml` | Link dictionary, first cut |
| `docs/frame-standard.md` | The frame standard |
| `tests/` | Package editing, reference properties, the full example build compared with its reference, and the proofs |

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

The proofs have their own READMEs (`python prototypes/assembly/demo.py`, `python prototypes/montecarlo/simulation.py`, `python prototypes/charts/build.py`, `python prototypes/reports/build.py`, `python prototypes/models/development.py`, `python prototypes/consolidation/build.py`, `python prototypes/assurance/demo.py`). The engine has its own instructions in `engine/README.md` (`npm install`, `npm test`). The probe has its own instructions in `addin-probe/README.md` (`npm install`, `npm run certs`, `npm start`; `npm test` for its Node tests).
