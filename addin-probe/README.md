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

The **HFG Model** tab and the runtime **HFG Build** tab are the spec's ribbon, generated from one registry, `src/commands.js`. Each command has a label, a glyph for its icon, a tip, the Modano command it matches and the phase that builds it. In the probe every command opens a view in the pane that says what it will do; none of them changes the workbook yet.

An add-in can define only one tab in its manifest, so the Build tab is created at runtime with `Office.ribbon.requestCreateControls` (RibbonApi 1.2). Its groups are Manage, Structure, Styles, Content, Charts, Review and Finish.

After changing the registry, regenerate the manifest, shortcuts and icons:

```
node tools/build-ribbon.mjs     # manifest.xml and shortcuts.json
python3 tools/make-icons.py     # assets/cmd, three sizes per command
npm test && npm run validate
```

The chart probes (`chart-z`, `chart-ibcs`) build a Z chart and an IBCS column chart with Office.js and record which formatting steps Excel accepted. The fully formatted versions, with hatching and a form control, are in `prototypes/charts/`.

## What it does not test

Performance on a full-size model, co-authoring, and the Home Hub endpoint itself. Those come in Phase 1 and 2 with the real engine.

## Files

| Path | Holds |
|---|---|
| `manifest.xml` | XML manifest: the HFG Model tab, the probe group, right-click menu, shared runtime (generated) |
| `shortcuts.json` | Keyboard shortcuts (generated) |
| `src/commands.js` | Command registry: both tabs, the right-click menu and the shortcuts |
| `src/probe.js` | Every probe, the results store, the ribbon command handlers and the Build tab definition |
| `tools/` | `build-ribbon.mjs` writes the manifest and shortcuts from the registry; `make-icons.py` draws the icons |
| `assets/cmd/` | Command icons at 16, 32 and 80 pixels (generated) |
| `src/ui.js`, `src/taskpane.html`, `src/taskpane.css` | The task pane |
| `src/zip.js` | Reads sheet names from an .xlsx so you can choose what to insert |
| `serve.mjs` | HTTPS static server on port 3000 |
| `test/` | Node tests: manifest and handler wiring, requirement checks, zip reading |

`npm run validate` checks the manifest with Microsoft's validation service; `npm test` runs the Node tests.
