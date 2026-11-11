# Independent project score

11 November 2026. Reviewer: Codex. Reviewed commit: `aef9e7765595c71189322830f96a54096c492452`.

**42.8 / 100: working alpha under Claude's rubric.**

The reusable engine, model library and reference calculations are the strongest parts. They build substantial financial models and support a tested core editing loop. The product around them is much less complete: most designed commands are placeholders, several convincing screens operate only on fictional samples, and the repository does not record the Excel results needed to substantiate everyday use, recovery or platform support.

This is an assessment of the add-in against its whole stated purpose, including modularity and Modano parity. It is not a score for the quality of one example workbook or the quality of the plan.

## Method and limits

I used the twelve rows, weights, maturity scale, half points and bands in Claude's 16:58 handoff in `collaboration.md`. Each contribution is `weight x points / 5`. I did not obtain or read Claude's score. The original rubric and its weights have not been changed.

Scores for a row with mixed progress reflect the material capabilities in that row, rather than the most advanced individual feature. Built engine functionality receives credit even when its separate ribbon command is unfinished. A sample screen does not receive credit as a workbook command.

I inspected the source, tests, command registry, parity mapping, manifests, build configuration and current board. I ran the four existing test suites, both TypeScript checks and a fresh bundle build. I also exercised the current dialogs in a browser, created a fictional consolidation workbook through the wizard, generated all three populated recipes, and inspected four rendered areas of the full model. LibreOffice recalculation and browser use are not evidence from Excel itself. No row exceeds 3.

The full requirements document and its complete feature matrix live outside this repository in a Claude document. The local command mapping does not reproduce that entire matrix. The command inventory below is therefore a measured command count, not an independently certified percentage of all Modano features. The catalogue targets of 19 model types and 63 modules are the rubric's supplied baseline.

The review changed documentation only. Existing product code and real workbooks were not edited. Generated fictional workbooks, logs, helper scripts and screenshots are in the ignored `build/review-scoring-2026-11-11/` folder.

## Scores

| Row | Weight | Points / 5 | Contribution | Evidence and reason |
|---|---:|---:|---:|---|
| 1. Module assembly and self-linking | 12 | 3.0 | 7.2 | Insert, add, add above, duplicate, rename, remove and delete have real service and planner implementations. Passing tests compare live writes with fresh assembly and check that category choices follow moves and renames. Replace, an explicit Mirror command and manual link/unlink are unfinished. The live equivalence tests have explicit chart, data-table and control limitations. |
| 2. Command coverage against the parity list | 12 | 1.5 | 3.6 | Of 114 actionable keys, 18 have substantive implementations, two have substantial partial implementations, six are sample-only and 88 are placeholders. That is 20 with meaningful workbook functionality, 17.5%, including partial commands and aliases. Only one of the Build tab's 58 actionable keys is implemented. |
| 3. Content library and model types | 8 | 2.5 | 4.0 | The production library loads 43 module definitions from 44 YAML files. New model offers three populated recipes: full financial model, budget against actuals and group consolidation, plus a blank frame and an assembly demo. This shows reuse across models, but does not deliver the supplied catalogue of 19 model types or the full funds, development and property breadth. Definition counts are not a like-for-like semantic mapping to 63 catalogue modules. |
| 4. Quality of the models produced | 8 | 3.0 | 4.8 | The frame, contents, navigation names, entity themes, checks, controls and report charts are implemented. Fresh builds contain 32 sheets and 106 chart parts in the full model; 29 sheets and 32 charts in the budget model. Numerical references pass after LibreOffice recalculation. Four rendered sheet areas show coherent professional layouts. Excel appearance, repair-free opening and live control/chart fidelity are not recorded as confirmed. |
| 5. Analysis and reporting | 6 | 2.5 | 3.0 | Scenario sheets and reporting, saved budgets/reforecasts, approve/lock and version comparison are built and tested. Sensitivity and simulation commands remain placeholders; Monte Carlo remains a Python proof and Impacts uses sample models. Manage versions opens a working versions screen but lacks the promised rename, delete, push and load operations. |
| 6. Data, roll forward and consolidation | 6 | 2.0 | 2.4 | The monthly consolidation engine is substantial: reference tests cover every group and month, eliminations, margin, NCI, statements and cash-flow checks. Pull, Push, Import, Map accounts and Roll forward have no workbook implementations. Group structure, Add entity and Change ownership still use the probe's sample. This cannot yet serve as a connected monthly planning workflow. |
| 7. Review, audit, finalise and share | 6 | 1.5 | 1.8 | Model Checks and Explorer are useful, while the assurance proof covers comparison, input records, releases and changes. Trace, formula-error review, consistency, inspector, work in progress, review notes, model compare, finalise, report pack and release-copy commands are unimplemented. Version comparison is not general workbook comparison. |
| 8. UX: surfaces, discoverability and consistency | 10 | 2.5 | 5.0 | The registry consistently routes to selection actions, dialogs, the task pane or sheets. The current wizard, searchable Insert picker, Explorer and group chart are legible; light and dark styles work in the browser. However, 88 placeholders and six sample-only commands dominate the available actions. Complete keyboard and screen-reader behaviour is not demonstrated; for example the Explorer tree lacks arrow-key navigation and the Saved by input has no accessible name. |
| 9. UX: safety of change | 8 | 2.0 | 3.2 | Preview/apply, required-module refusal and preservation of drop-down choices are real. The before-and-after key-output guard and change log are proven in Python but are not called by the live service's apply path. Add, Add above and Duplicate apply immediately under the agreed selection workflow. Grouped undo and partial-write recovery have no recorded Excel proof, although the UI promises Ctrl+Z. Unlocking an approved version does not collect the reason promised in the registry. |
| 10. Proven in Excel, platforms and speed | 10 | 1.0 | 2.0 | There is a detailed probe and requirement-set handling, but no completed platform results or real-model timing record in the inspected repository and board. Earlier requests to test controls, charts, dialogs and undo remain requests, not recorded passes. Neither these browser checks nor the LibreOffice suite raises this row to Excel confirmation. |
| 11. Engineering quality | 8 | 2.5 | 4.0 | 292 tests passed across four suites, with four skip reports; both TypeScript checks and the bundle build pass, and the rebuilt bundle is byte-identical to the committed bundle. Golden fixtures, numerical references and separate engine/add-in layers provide meaningful evidence. CI omits the LibreOffice and browser dependencies for important checks, the browser tests reference removed globals, no lint script is configured, and the product still uses the probe's manifest. |
| 12. Onboarding, help, install and deployment | 6 | 1.5 | 1.8 | Windows start/stop scripts, sideloading instructions and developer READMEs exist. They still install and start the probe with Node, localhost hosting and certificate setup. Product help is a short probe page; some READMEs describe an earlier interface. No deployable product manifest, staff rollout or verified non-developer installation is present. |
| **Total** | **100** | | **42.8 / 100** | **Working alpha, 41 to 60 in the supplied bands.** |

## Command inventory

The source is `addin-probe/src/commands.js`, cross-checked against `addin/src/probe-entry.ts`, `addin/src/service.ts`, `addin-probe/src/probe.js`, `addin-probe/src/hfg-dialog.js` and `docs/command-parity.json`. The registry has 124 entries, including ten menu containers. Excluding those containers leaves 114 actionable keys. All 114 have local parity-map entries.

An implemented key means it performs a meaningful operation on the current workbook or its model. It does not mean every detail of its advertised behaviour is complete. Partial keys perform real work but omit major advertised operations. Sample-only keys work on the probe's fictional state and do not operate on the open model. The classification of every registry entry is retained in [the command inventory](CODEX_COMMAND_COVERAGE_2026-11-11.json).

| Phase | Implemented | Partial | Sample only | Placeholder | Actionable total |
|---|---:|---:|---:|---:|---:|
| 1 | 18 | 2 | 0 | 35 | 55 |
| 2 | 0 | 0 | 6 | 43 | 49 |
| 3 | 0 | 0 | 0 | 10 | 10 |
| **Total** | **18** | **2** | **6** | **88** | **114** |

| Surface in registry | Implemented | Partial | Sample only | Placeholder | Actionable total |
|---|---:|---:|---:|---:|---:|
| HFG Model tab | 17 | 2 | 5 | 26 | 50 |
| HFG Build tab | 1 | 0 | 0 | 57 | 58 |
| Right-click-only keys | 0 | 0 | 1 | 5 | 6 |

The right-click menu also reuses keys already counted on a ribbon tab. The six above are the additional context-only keys, not the size of the whole right-click menu. Aliases such as the two Explorer entries remain in this command count because the rubric asks about visible command coverage. They do not establish additional unique capabilities.

The eighteen implemented keys are New model, Explorer, Checks, Properties, Insert, Duplicate, Rename, Delete, Add category, Add above, Remove category, Show the set, Save version, Approve budget, Lock or unlock, Compare versions, Builder mode and Build's Explorer. Compare versions currently goes to an existing comparison sheet; it does not insert a missing module or set up both choices. Show the set is likewise a sheet navigator.

The two substantial partial implementations are Modules > Links, which reads links but cannot link or unlink them, and Manage versions, which exposes existing save/approve/lock functionality but not the promised management operations. The six sample-only keys are Group structure, Add entity, Change ownership, Impact of a change, Impacts sheets and the separate context-only Impacts key.

The 20/114 figure is 17.5%. Counting only the eighteen implementation entries gives 15.8%. Phase 1 alone is 20/55, 36.4%, including partial implementations. These percentages are useful coverage evidence, not substitutes for the maturity scale or the total rubric score.

## Library and model evidence

| Populated recipe | Months | Module instances | Sheets | Layout rows | Native chart parts |
|---|---:|---:|---:|---:|---:|
| Full financial model | 36 | 57 | 32 | 4,512 | 106 |
| Budget against actuals | 36 | 61 | 29 | 3,366 | 32 |
| Group consolidation | 36 | 2 | 16 | 2,311 | 0 |

These are builds of the current committed recipes, not real customer models. The consolidation instance is a large module that creates multiple sheets, which illustrates why a simple module count is not a measure of functional breadth. The budget build reports ten unused `wc.stock` output warnings; the other two report none. Passing numerical checks do not establish that every policy or real-world case is covered.

The library genuinely separates definitions, recipes and a reusable engine. It is not merely a hard-coded version of the development example. That architectural achievement deserves credit even though user-facing content creation, replacement, review and library management commands are incomplete.

I inspected the rendered Contents, Income summary, Financials and Budget summary areas of a fresh full-model build. They show a linked contents structure, quiet colours, coherent hierarchy, dense monthly tables, actual/forecast distinctions, totals and variance presentation. The previews are cropped sheet areas, not a complete print or chart audit. LibreOffice does not display the classic controls as Excel does and substitutes fonts. The inspection therefore supports row 4's engineering maturity, not an Excel-confirmed score of 4.

## Checks run

| Check | Result |
|---|---|
| `engine`: `npm test` | 110 passed; none skipped. |
| `addin`: `npm test` | 36 passed; none skipped. |
| `addin-probe`: `npm test` | 25 passed; one real-template test skipped because `HFG_TEMPLATE` was not supplied. |
| Python: existing WSL environment, `python -m pytest -q -ra` | 121 passed; three skips reported; 330.60 seconds. Two require the template. The browser module was skipped because Python Playwright is absent. |
| `engine`: `npm run typecheck` | Passed. |
| `addin`: `npm run typecheck` | Passed. |
| `addin`: fresh build to a separate review path | Passed. SHA-256 equals the committed bundle: `FE091C4BB475656A3B5DCE6E0C84DED7C2163BA3FF86296106D383EFBD8B3442`. |
| Fresh builds of all three populated recipes | Passed. Counts above come from assembly and the actual XLSX package contents. |
| Current UI in Chromium | Wizard validation, recipe selection and fictional consolidation creation worked. Insert generated a preview. Explorer, light/dark Insert, versions, Replace's placeholder and the sample group screen were inspected. This was outside Excel. |
| Four model sheet previews | Exported from recalculated LibreOffice and rendered with Windows Poppler; visually inspected. |
| Microsoft external manifest validation | Not completed. The initial attempt could not reach the service; automatic approval review rejected sending the manifest externally without explicit permission. Local XML/wiring tests passed as part of the probe suite. |
| Lint | No lint script or lint job is configured in these packages. No lint pass is claimed. |

Windows JavaScript checks used Node 26.13.0 and npm 11.20.0. The Python suite used the project's existing Ubuntu 24.04 environment and `~/hfg-venv`. No dependencies or certificates were installed, no existing workbook was opened in Excel, and nothing was deployed or pushed.

The build hashes were checked directly. Test skips are not treated as passes. The dormant browser tests are a separate issue from the missing dependency: `tests/test_wizard_browser.py` calls `window.HfgWizard` and `window.HfgInsert`, while the current bundle exports `HfgAddin` and `HfgAddinViews`. I confirmed the old globals are undefined in the running browser page. I did not report these tests as executed failures.

Important test limits are visible in `addin/test/live.test.ts`: some equivalence cases exclude chart or data-table operations, and the full-model case compares selected sheets without names because live controls differ from package controls. These are useful focused tests, not proof of complete live/package equivalence in Excel. CI's Python job installs neither LibreOffice/UNO nor Python Playwright, so the repository's most valuable recalculation and UI checks are not all enforced there.

## What most improves actual usability

1. **Prove the existing core in Windows Excel and record it.** Create and save/reopen each model; insert, duplicate, rename and remove categories; check formulas, classic/live controls, charts, metadata and dialog messaging. Record workbook size, Excel build, timings, error/alert counts and repair prompts. Then repeat on the other target platforms. This supplies evidence currently missing across several rows.
2. **Close the live safety gap.** Integrate the key-output comparison and change log already proven in Python. Establish recovery after a write fails, and align the undo wording with demonstrated host behaviour. Resolve how the agreed immediate selection commands meet the rubric's preview requirement, rather than silently changing that workflow.
3. **Finish a connected monthly workflow.** Deliver the Home Hub connection, Pull, mapping contract, Roll forward, Save version and comparison end to end on a fictional organisation. The consolidation engine cannot compensate for the missing data workflow.
4. **Complete core modular editing and review.** Replace, mirror and link/unlink, then Trace, error review, finalise and a usable report handover. These complete existing promises before adding more visible placeholder actions.
5. **Make installation and verification part of the product.** Give the add-in its own manifest and deployment route, replace probe-only help, repair the dormant browser tests, and enforce the recalculation and browser checks in CI.

## Where the rubric needs correction

The score above uses the rubric unchanged so Kelvin can compare the independent results. I recommend the following changes for its next use; Kelvin retains the decision.

**Coverage and evidence maturity are mixed.** A narrow, tested implementation can earn 3, while a broad prototype earns 2. Rows contain several features with different maturity. Report feature coverage and verification maturity separately, retain the row evidence, and state explicitly how mixed rows aggregate. The 42.8 score is neither 42.8% feature parity nor a probability that the product is safe.

**Menu entries and module files are not equal units.** A sheet navigator, an Explorer alias, a data integration and a whole consolidation engine are all counted as commands or modules at different points. The parity map also includes HFG additions and one-to-many mappings to reference features. Keep the transparent registry count, but add a stable, deduplicated feature checklist and a catalogue crosswalk before interpreting counts as parity.

**Readiness needs minimum conditions, not only a weighted average.** The 61 to 80 and 81 to 100 bands sound like operational approval and parity. High scores elsewhere can offset a missing release or safety capability. Those labels should also require a successful staff installation, an Excel-confirmed core workflow, recovery evidence and no unresolved critical financial defects.

**The Excel cap is too broad for some rows.** It is appropriate for behaviour, usability and speed. Independent numerical checks, clean-room provenance and repeatable builds should have their own verifiable engineering standards rather than being capped solely by whether Excel has been used. This is a proposed future rubric change; I kept the cap for this score.

**The safety row conflicts with the agreed selection workflow.** The rubric asks for a preview before every structural command, but Kelvin's 13 October workflow intentionally makes Add, Add above and Duplicate immediate selection actions. Define which changes need a preview and what compensating undo/check/recovery contract makes immediate actions acceptable. Do not penalise compliance with one agreed design without acknowledging the conflict.

**The external benchmark must identify the product and version.** The public [Modano architecture document, version 15.0](https://app.modano.com/files/Modano-Software%20Architecture.pdf) describes a Windows COM add-in. Its current [support page](https://www.modano.com/support) also describes a browser-based platform. That does not establish the platform support of every Modano edition. The agreed add-in checklist should stay the baseline until the target edition and version are explicit. The [public library catalogue](https://app.modano.com/content_add_ons) corroborates the financial model, consolidation, funds, multi-business-unit and property breadth, but does not prove a complete feature-by-feature comparison. [Operis's own feature descriptions](https://www.operisanalysiskit.com/) support using formula tracing, comparison and model review as reference behaviours for row 7. No proprietary code, commercial module files or paywalled product contents were read for this review.

## Evidence locations

The committed companion [command inventory](CODEX_COMMAND_COVERAGE_2026-11-11.json) gives the key, tab, phase, classification, reason and local parity mapping for every registry entry. The source locations most relevant to the findings are:

- `addin/src/probe-entry.ts`: working view map and selection actions.
- `addin/src/service.ts`: current-model reads, previews, direct selection writes and the live apply path.
- `addin/src/live/apply.ts`: Office.js writes and separate metadata replacement.
- `addin/src/modules/core.ts`: category changes and choice preservation.
- `addin/src/versions/core.ts` and `addin/src/versions/view.ts`: version capabilities and unlock behaviour.
- `addin-probe/src/probe.js`: routing and sheet navigators.
- `addin-probe/src/hfg-dialog.js`: sample screens and explicit placeholders.
- `engine/src/resolve.ts`, `engine/src/plan.ts` and `engine/src/consolidation.ts`: modular links, change plans and consolidation.
- `prototypes/assurance/guard.py`: before-and-after comparison and change-log proof.
- `tests/test_full_model.py`, `tests/test_budget_model.py`, `tests/test_consolidation_model.py` and `engine/test/golden.test.ts`: numerical and structural references.
- `addin/test/live.test.ts`, `tests/test_wizard_browser.py` and `.github/workflows/tests.yml`: important verification limits.

Ignored local evidence includes `command-inventory.mjs`, `command-coverage.json`, `model-inventory.mjs`, `model-coverage.json`, the four suite logs, typecheck/build logs, browser screenshots, the three fictional XLSX builds and `model-shots/review-1.png` to `review-4.png`. These files are review outputs, not production data or release artifacts.
