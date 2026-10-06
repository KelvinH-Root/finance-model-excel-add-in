# HFG add-in probe

A throwaway Excel add-in that tests, in real Excel, the Office.js behaviour the requirements spec could not settle from documentation: write speed, row inserts inside blocks, metadata after save and reopen, what sheet insertion carries, PDF export, ribbon and right-click updates, the designed ribbon, native charts, one-step undo and Entra sign-in. It is Phase 0 of the spec, not the HFG add-in.

Everything it adds to a workbook is named `zProbe_`, and **Clean up** removes it. Use a blank workbook or a copy, never a live model. The results hold timings and yes or no answers, no financial data.

## What you need

- Node.js 20 or later.
- Excel from Microsoft 365 on Windows or Mac. For the web test, a blank workbook saved in OneDrive or SharePoint.
- For the sign-in test only: an Entra app registration (see below).

## Run it on Windows or Mac

```
cd addin-probe
npm install
npm run certs      # once: installs and trusts a localhost certificate
npm start          # starts https://localhost:3000 and opens Excel with the probe loaded
```

In Excel, open the **HFG Model** tab and choose **Open probe** in the **Probe** group at its right-hand end. When you are finished, `npm run stop` removes the sideloaded add-in.

## Run it in Excel on the web

1. `npm run serve` on the same computer as the browser.
2. Open a blank workbook stored in OneDrive or SharePoint.
3. Choose Home, Add-ins, then the option to upload your own add-in (Microsoft moves the label between releases; it is usually More Add-ins, My Add-ins, Upload My Add-in), and pick `manifest.xml`.

## Work through the pane

1. **Run automatic tests.** About a minute; Excel may pause during the write and insert tests.
2. **Tests that need you.** Each card says what to do and asks you to confirm what you saw:
   - insert sheets from a copy of the template (pick sheets with drop-downs, check boxes and charts), then copy one;
   - read the workbook as a file, then open a copy;
   - export to PDF and check the print setup;
   - write three cells in one undo group and press Ctrl+Z once;
   - grey out and enable the ribbon button and the right-click item; click **Mark cell**, **Menu test, Mark from menu**, the right-click **HFG, Mark this cell (probe)**, and press Ctrl+Alt+Shift+J (Cmd+Option+Shift+J on Mac);
   - the designed ribbon: click buttons, menu items and right-click items on the **HFG Model** tab and check each opens its view at the top of the pane; show and hide the runtime **HFG Build** tab; hide and show the Analysis group (needs RibbonApi 1.3);
   - open a dialog, and sign in if you have an app registration.
3. **Metadata after save and reopen.** Save, close and reopen the workbook (on another platform if you can), open the probe and choose **Check the metadata**.
4. **Write results to a sheet**, then **Download results** or save the workbook and send the `zProbe_Results` sheet. Run it on each platform you use.
5. **Clean up.**

## Sign-in test (optional)

Register an application in Entra ID (App registrations, New registration):

- Supported account types: this organisation only.
- Platform: Single-page application, with two redirect URIs: `brk-multihub://localhost:3000` (Excel desktop) and `https://localhost:3000/src/taskpane.html` (Excel on the web).
- API permissions: Microsoft Graph, delegated, `User.Read`.

Paste the Application (client) ID and the Directory (tenant) ID into the probe and choose **Sign in**. On the web, sign-in only works for a workbook stored in OneDrive or SharePoint.

## The designed ribbon

The **HFG Model** tab and the runtime **HFG Build** tab are the spec's ribbon, generated from one registry, `src/commands.js`. Each command has a label, a glyph for its icon, a tip and the phase that builds it; how each compares with the reference add-in is kept outside the add-in, in `docs/command-parity.json`. In the probe every command opens a view in the pane that says what it will do; none of them changes the workbook yet.

An add-in can define only one tab in its manifest, so the Build tab is created at runtime with `Office.ribbon.requestCreateControls` (RibbonApi 1.2). Its groups are Manage, Structure, Styles, Content, Charts, Review and Finish.

After changing the registry, regenerate the manifest, shortcuts and icons:

```
node tools/build-ribbon.mjs     # manifest.xml and shortcuts.json
python3 tools/make-icons.py     # assets/cmd, three sizes per command
npm test && npm run validate
```

**Group** on the Model group (Ctrl+Alt+Shift+G for Group structure) is the one menu with designed views rather than placeholders, drawn on a fictional sample group generated from the consolidation proof (`python3 tools/group-sample.py` writes `src/group-sample.js`): Group structure shows the entities as a tree (share held, owned by the top, outside investors, planned or actual, the groups each rolls into) and each group's roll-up for the year shown; Add entity takes code, name, parent, share held, member from and capital, previews where it goes and what the change plan writes, then adds it to the sample in the pane; Change ownership previews a new parent or share held from a date. Nothing is written to the workbook. Read this workbook reads the entity register (the Ent_ names) from an open consolidation workbook.

**Explorer** (Model group; also Build > Manage and Modules > Links) and **Impacts** (Analysis group, and right-click > Show impacts) have designed views too, on the assembly proof's demo model (`python3 tools/model-sample.py` writes `src/model-sample.js`; `src/model-calc.js` is the demo's own calculation, standing in for Excel). The Explorer shows the model in one pane: the tree of sections, sheets and modules, and Composition, Links (a diagram of what the module takes from and sends to; click to move along), Properties and Checks tabs. Impacts shows Impact of a change (every statement line that moves, the ties and the chain of links) and Impacts sheets (each item's effect by entity, with eliminations and group for intergroup items, in either the demo's or the consolidation example's accounts).

The speed check probe (`speed`, ExcelApi 1.14) runs on whatever workbook is open: it calculates each sheet in turn and times it, and scans the formulas for volatile functions, whole-column and whole-row references and links to other workbooks. Open BUD25 or the template first to see what Adopt workbook would report.

The chart probes (`chart-z`, `chart-ibcs`) build a Z chart and an IBCS column chart with Office.js and record which formatting steps Excel accepted. The fully formatted versions, with hatching and a form control, are in `prototypes/charts/`.

## What it does not test

Performance on a full-size model, co-authoring, and the Home Hub endpoint itself. Those come in Phase 1 and 2 with the real engine.

## Files

| Path | Holds |
|---|---|
| `manifest.xml` | XML manifest: the HFG Model tab, the probe group, right-click menu, shared runtime (generated) |
| `shortcuts.json` | Keyboard shortcuts (generated) |
| `src/commands.js` | Command registry: both tabs, the right-click menu and the shortcuts |
| `src/group.js`, `src/group-sample.js` | Group structure, roll-up, Add entity and Change ownership views, and their sample group (generated) |
| `src/explorer.js`, `src/impacts.js`, `src/model-calc.js`, `src/view-kit.js`, `src/model-sample.js` | Explorer and Impacts views, the demo model's calculation, shared view helpers, and their sample (generated) |
| `src/probe.js` | Every probe, the results store, the ribbon command handlers and the Build tab definition |
| `tools/` | `build-ribbon.mjs` writes the manifest and shortcuts from the registry; `make-icons.py` draws the icons; `group-sample.py` and `model-sample.py` write the samples |
| `assets/cmd/` | Command icons at 16, 32 and 80 pixels (generated) |
| `src/ui.js`, `src/taskpane.html`, `src/taskpane.css` | The task pane |
| `src/zip.js` | Reads sheet names from an .xlsx so you can choose what to insert |
| `serve.mjs` | HTTPS static server on port 3000 |
| `test/` | Node tests: manifest and handler wiring, requirement checks, zip reading, the models behind the Group, Explorer and Impacts views |

`npm run validate` checks the manifest with Microsoft's validation service; `npm test` runs the Node tests.
