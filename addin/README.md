# HFG Excel add-in (Phase 1)

The add-in's task pane features, built on the engine (`../engine`). TypeScript, bundled with esbuild for the browser; the Node tests run the sources directly (Node 22.18 or later).

```
npm ci             # here and in ../engine
npm test           # wizard tests
npm run typecheck
npm run build      # writes ../addin-probe/src/wizard.bundle.js
```

## New model wizard

`src/wizard/core.ts` holds the choices, their checks and the workbook they build; `src/wizard/view.ts` draws the steps in the task pane. Steps: Entity (brand, logo and theme, the name on the contents), Model (title, what it starts from, the Prepared by line, notes), Timeline (first month, months, financial year end, last month of actuals, denomination), Display (status phrases in the model name line), Review (the contents and timeline in words). Create builds the workbook with the engine's package writer in HFG's standard frame and opens it with `Excel.createWorkbook` (ExcelApi 1.8); outside Excel it offers the file to download.

What a model starts from: a blank frame, or the assembly demo (fictional data). Catalogue model types join as their modules reach the library; Scenarios joins the wizard with the Scenarios module.

Until the add-in has its own manifest, the wizard runs inside the Phase 0 probe: `npm run build` writes `addin-probe/src/wizard.bundle.js` (generated, committed so the probe keeps its no-build-step rule), and the probe's HFG Model > New model draws it. Rebuild the bundle whenever the engine, the library or the wizard changes; CI checks it is current.

`tests/test_wizard_browser.py` (repo root) steps through the bundle in Chromium, creates a model and recalculates it in LibreOffice.
