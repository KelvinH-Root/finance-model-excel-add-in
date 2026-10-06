# HFG Excel add-in (Phase 1)

The add-in's task pane features, built on the engine (`../engine`): New model, Insert module and the live writer. TypeScript, bundled with esbuild for the browser; the Node tests run the sources directly (Node 22.18 or later).

```
npm ci             # here and in ../engine
npm test           # wizard tests
npm run typecheck
npm run build      # writes ../addin-probe/src/addin.bundle.js
```

## New model wizard

`src/wizard/core.ts` holds the choices, their checks and the workbook they build; `src/wizard/view.ts` draws the steps in the task pane. Steps: Entity (brand, logo and theme, the name on the contents), Model (title, what it starts from, the Prepared by line, notes), Timeline (first month, months, financial year end, last month of actuals, denomination), Display (status phrases in the model name line), Review (the contents and timeline in words). Create builds the workbook with the engine's package writer in HFG's standard frame and opens it with `Excel.createWorkbook` (ExcelApi 1.8); outside Excel it offers the file to download.

What a model starts from: a blank frame, or the assembly demo (fictional data). Catalogue model types join as their modules reach the library; Scenarios joins the wizard with the Scenarios module.

Until the add-in has its own manifest, the features run inside the Phase 0 probe: `npm run build` writes `addin-probe/src/addin.bundle.js` (generated, committed so the probe keeps its no-build-step rule), and the probe's HFG Model > New model and Modules > Insert draw them. Rebuild the bundle whenever the engine, the library or the add-in changes; CI checks it is current.

## Insert module and the live writer

`src/insert/` reads the model from the open workbook's metadata part, lists the library's modules by area (a module that goes in once is offered only while it is absent), takes the module's inputs, previews the change plan in plain words and applies it. `src/live/apply.ts` is the live writer: each plan operation becomes Office.js calls (rows inserted and deleted so Excel moves every reference, cells written, named styles set, row heights, hyperlinks to `Go_` names with screen tips, checks that turn bold red, validations, names, charts) and the metadata part is replaced. Formats arrive as named styles, so the writer types no colour but the check red; the outline is cleared and regrouped because Office.js cannot read a row's level. Requirement sets: ExcelApi 1.7, 1.8 and 1.10.

`test/live.test.ts` runs the live writer against a stand-in for Excel working on the engine's workbook image and compares the result with a fresh build of the same model (inserts, a removal, a new section, a blank model growing into the demo). Charts made live carry their series names as text until the model is next rebuilt; one-step undo of a whole insert needs ExcelApi 1.20 and is not wired yet.

`tests/test_wizard_browser.py` (repo root) steps through the bundle in Chromium: the wizard creates a model that recalculates in LibreOffice, and Insert previews a module on the demo model.
