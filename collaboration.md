# Collaboration board: hfg-models

Shared message board for Kelvin (owner, final say), Claude and Codex.
Read before responding. Append only, newest at the bottom.
Types: DECISION, QUESTION, BLOCKER, HANDOFF, DISAGREE, NOTE.

---

## [2026-10-05 16:11 NZ] KELVIN | NOTE
Board opened. Scope and current state are in README.md and docs/development-model.md. Background lives in the Technical Roadmap project docs (hcp-development-model-template, hfg-model-library, hfg-excel-add-in, numeric-finance-ops-tool).

## [2026-10-05 16:11 NZ] CLAUDE | DECISION
Decisions carried in from the Technical Roadmap project (all Kelvin's, 2 to 5 October 2026).
- The template is the mould: build from a copy of Budget_Template.xlsx so form controls, named styles and the scenario manager survive. The HCP palette (olive #575F46) replaces the theme colours, so every sheet follows.
- No VBA. Python (build script, Python in Excel for Monte Carlo) and Office Scripts (scenario buttons); an Office.js add-in later.
- Actuals come from Home Hub, not Xero directly.
- Home Hub Planning is the main planning tool (Runway-style, AI built in). Excel models are an extra layer that pulls from and pushes to it. Push rule: every plan line has one source; a model writes only to lines it owns, anything else goes back as a proposed scenario accepted in Home Hub.
- The NZ Housing Fund model is the second Excel model, where the library and assembler get extracted.
- No workbooks and no real financial data in git.

## [2026-10-05 16:11 NZ] CLAUDE | DECISION
Build approach for the development model, version 0.1.
- The build edits the template at package level (zip parts), not with openpyxl, because openpyxl drops form controls, charts, Python in Excel results and web extensions. New sheets are written as XML using the template's own cell formats.
- Modules declare rows by key; formulas refer to rows by key across sheets (hfgmodels/layout.py).
- hfgmodels/dev/reference.py is the specification. Tests rebuild the workbook, recalculate it in LibreOffice and compare every row. At this entry: 0 mismatches, 0 model errors, 0 alerts on the demo recipe.
- Demo recipe uses three fictional sites. HCP's real inputs are not committed.

## [2026-10-05 16:11 NZ] CLAUDE | NOTE
HCP's six models were reviewed (detailed reviews kept outside git because they hold real figures). Version 0.1 reproduces v4's funding pattern (LTC ceiling, land loan before facility start, catch-up draw, two-step exit with margin to HHLP) and fixes: the broken facility limit selector, interest capitalised above the limit, dropped pre-cut-over budget, the dormant upfront fee, double-counted sale-month cash, missing GST on costs, no checks. See docs/development-model.md.

## [2026-10-05 16:11 NZ] CLAUDE | QUESTION
Questions for Kelvin on HCP's funding logic. The model has a switch for each, so the answers set defaults rather than block work.
1. Funding pattern: v4 draws debt and equity together up to 70% of cost to date; the September 2025 drawdown forecast used equity first, then debt. Which is standard for each lender?
2. Catch-up draw: when the facility starts, v4 draws to the ceiling and releases equity already spent. Intended?
3. Once the limit binds, v4 keeps capitalising interest above the limit. Version 0.1 caps at the ceiling and equity pays the rest. Agree?
4. Sale price basis: As Complete (BTS) value including GST, with the LP sale standard-rated because the SPV or Fund is not GST registered. Is the SPV GST registered? If it is, the GST on the internal sale is claimable and the group cost disappears.
5. Facility limit basis per site: override, LTC on budget, LVR on As Complete, or LVR on BTR?
6. Upfront fees (1.5% of the limit in v4's inputs) never fire in v4. Should they be charged?
7. Te Paeroa: keep it as a group revolver carrying other sites' land loans? It is not in version 0.1.
8. Cut-over: v4 is labelled 31 March but its actuals end 28 February. Which month should the opening position sit at?
9. Should the model take line budgets, profiles and escalation directly from the upstream HCP cashflow tool (the project sheets), or replace that tool?
10. Te Piringa purchase debt: v4 uses 75% of BTR; the March model sizes it as the lower of 80% of BTS and net rent / 1.3 / 5%. Which should the NZ Housing Fund model use?

## [2026-10-05 16:11 NZ] CLAUDE | HANDOFF
Next work, in order: answers to the questions above set defaults; Te Paeroa group revolver; scenario manager links and the Monte Carlo layer; wiring the template's summary pages and removing unused template sheets; Home Hub contracts (hh_out_plan, hh_out_intergroup). Codex can take any item; append a DECISION entry before starting one.

## [2026-10-05 17:30 NZ] KELVIN | DECISION
The product is an Excel add-in like Modano, not a version of the HCP model. HCP's models are inspiration for what the add-in must support; it must not be built around a specific model.

## [2026-10-05 17:30 NZ] KELVIN | DECISION
Three follow-on choices. The development model code stays as a test example only (moved to examples/development). A requirements spec comes before any add-in code. The repo keeps the name hfg-models.

## [2026-10-05 17:30 NZ] CLAUDE | NOTE
Requirements spec drafted as a Claude Doc, "HFG Excel add-in: requirements spec", in the Technical Roadmap project: purpose, principles, concepts, capabilities, module definition format, model metadata, engine options (recommendation: one TypeScript engine shared by the add-in and a command-line build), Office.js limits, requirements drawn from HCP's models, first module library, phases with gates, open decisions. The earlier HANDOFF entry and the ten HCP funding questions are parked until development modules are built.

## [2026-10-06 11:25 NZ] KELVIN | DECISION
The add-in must reach feature parity with Modano's add-in, shaped as an internal tool. The learnings from the reference models (the Modano-built template and BUD25, the EXL Cloud demo PDFs) belong in the spec, and the catalogue of model types should be researched.

## [2026-10-06 11:25 NZ] CLAUDE | NOTE
Spec updated (same Claude Doc, now with tabs). Spec tab: parity at a glance, learnings from the reference models, two writers in one TypeScript engine (package writer for new and adopted models, form controls, data tables and theme; live writer for everyday edits through Office.js), Office.js limits checked against ExcelApi 1.21, ribbon and task pane layout, polish standard with measurable targets, how sure we are, phases 0 to 4, open decisions. Modano parity tab: 129 features from Modano's user guide contents: 96 same, 16 better, 12 different, 5 not needed, none blocked; 53 land in Phase 1, 98 by Phase 2, all 124 needed by Phase 3. Model catalogue tab: 19 HFG model types from 50 modules, ranked by reuse; first wave development feasibility and finance, hold model, NZ Housing Fund. NZ rules register tab: 49 rules with sources. Finding: the template and BUD25 are released Modano workbooks; public evidence does not show the EXL Cloud demos are Modano-built.

## [2026-10-06 11:25 NZ] CLAUDE | HANDOFF
addin-probe/ is ready for Kelvin: npm install, npm run certs, npm start (Windows or Mac), or upload manifest.xml in Excel on the web with npm run serve running. Work through the pane, write results to a sheet and send it back, then Clean up. Manifest passes Microsoft's validation service; 12 Node tests pass. Phase 1 waits on the spec's first four open decisions (engine, module format, platform floor, probe run).

## [2026-10-06 13:10 NZ] CLAUDE | NOTE
Two independent reviews ran after the spec update. The spec review found about 30 Modano user guide pages with no parity row, a contradictory platform floor, the clean room claim clashing with the Modano-derived template, and no exit criteria; all fixed in the doc. The parity tab now has 143 Modano rows (110 same, 14 better, 13 different, 6 not needed; 59 by Phase 1, 111 by Phase 2, 137 by Phase 3) plus 3 HFG additions, which supersedes the figures in the 11:25 note. New in the spec: platform floor ExcelApi 1.20 recommended, metadata drift and co-authoring lock, exit criteria per phase, template licence as an open decision. The API review of the probe found six bugs (now fixed in commit 6145231), including insertWorksheetsFromBase64 returning ids and Range.address having no $ signs.

## [2026-10-06 13:05 NZ] KELVIN | DECISION
The add-in needs Modano's modular architecture: a module inserted while a model is being built links itself to the rest of the model, as in Modano's example model. The catalogue should be shaped on professional models, and the Monte Carlo and cash flow modelling in the EXL PDFs belong in the design.

## [2026-10-06 13:05 NZ] CLAUDE | NOTE
Two Phase 0 proofs added, both model-agnostic and fictional. prototypes/assembly: module definitions in YAML (single, category or mirror; inputs by each, total or mirror; required or optional), link resolution repeated until nothing new appears, a change plan diffed from two layouts, applied to a built workbook through LibreOffice as the live writer would through Office.js. Inserting a revenue line and a facility, and removing them, matches a fresh build in every formula, value and name (tests/test_assembly.py, 11 tests). prototypes/montecarlo: seeded Monte Carlo with one native data table (HDR generator, Latin hypercube, Cholesky copula, five distributions), a scenario events register (bullet, repeat, spread, start and end, percentage, shift, scale) and a self-checking workbook; every trial matches the Python reference (tests/test_montecarlo.py, 8 tests). hfgmodels.verify.recalculate now rebuilds wide data tables, because LibreOffice imports only their first two columns. The probe gained calculation mode and Monte Carlo data table probes (34 in all). The spec has new sections (How modules link themselves in, Simulation, Cash flow forecasting) and the catalogue tab a professional models table and four new modules (54 in all).

## [2026-10-06 13:05 NZ] CLAUDE | HANDOFF
For Kelvin: build the demo with python prototypes/montecarlo/simulation.py, open build/montecarlo/mc_demo.xlsx in Excel (desktop and web), and run the probe's Monte Carlo card; confirm the HDR generator's terms (new open decision). For any agent: the proofs are evidence for the spec; Phase 1's TypeScript engine reimplements the assembly design, it does not extend the Python prototype.

## [2026-10-06 14:09 NZ] CLAUDE | NOTE
Ribbon, drivers, controls, framework sheets and charts added to the spec; the probe carries the designed ribbon and a charts proof is in prototypes/charts.
- Ribbon: HFG Model tab in the manifest (seven groups mapped to Modano's main tab, 32 commands), an HFG right-click submenu (12 items) and five Ctrl+Alt+Shift shortcuts. HFG Build is a contextual tab created at runtime (RibbonApi 1.2), because a manifest allows only one custom tab; its groups are Manage, Structure, Styles, Content, Charts, Review and Finish (34 commands, 24 menu items). One registry, addin-probe/src/commands.js, generates the manifest, shortcuts, icons and the Build tab (node tools/build-ribbon.mjs, python3 tools/make-icons.py). Probe now 0.2.0 with 39 probes, including the designed ribbon, the Build tab, hiding a group (RibbonApi 1.3) and live Z and IBCS charts.
- Drivers and controls: method drop-down per category with a scenario layer, working capital allocation drop-downs to cash or debtor and creditor profiles, delay days as validated inputs (as in Modano's example). All 51 of Modano's example controls are mapped. The package writer writes classic form controls; live commands use in-cell drop-downs and checkboxes on the same DD_ and CB_ names. Parity now 113 same, 14 better, 10 different, 6 not needed, plus four HFG additions (native IBCS, Z and waterfall charts is the fourth).
- Framework sheets (Time, Scenarios, Seasonality, Lookups, Checks, Contents, reports) exist in the template but are not yet wired to modules: Time, Lookups, Checks and Contents in Phase 1; Scenarios, Seasonality, working capital and reports in Phase 2.
- prototypes/charts (3 tests): native Z chart matching the template's Z-Chart numbers, two waterfalls, IBCS column and variance charts, a classic combo box and an in-cell drop-down, all formula driven. Rendered and recalculated in LibreOffice; not yet opened in Excel.

## [2026-10-06 14:09 NZ] CLAUDE | HANDOFF
For Kelvin: open the charts demo (python prototypes/charts/build.py, or the copy sent in chat) in Excel desktop and on the web and check the combo box, the hatched forecast bars and the signed labels; when running the probe, click through both ribbon tabs and the right-click menu. For any agent: the registry in commands.js is the single source for commands; change it, then regenerate rather than editing manifest.xml by hand.

## [2026-10-06 14:50 NZ] KELVIN | DECISION
The Scenarios and Seasonality modules are wired in Phase 1, not Phase 2. Seasonality applies only to the models where it is relevant, not to every model.

## [2026-10-06 14:50 NZ] CLAUDE | NOTE
Spec, catalogue and parity tabs updated for the decision above, and the Z chart rebuilt to match Kelvin's.
- Seasonality is its own optional module (the catalogue now has 55), linked into driver lines through an optional link, with profiles by method: even, history (up to six prior years, any year can be left out, for example one holding the 2020 lockdown), working days, or typed; weights add to 100% with a check, and no history falls back to an even spread with an alert instead of the template's #DIV/0!. On by default for M13 entity budgets built in Excel, M10 HCL business and M15 13-week cash (by week); optional for M04 land (earthworks season) and M05 to M07 (monthly years only); not used by the development, annual or Home Hub models, whose timing comes from phasing profiles, the working-day calendar, payment timing or Planning. The spec's Drivers, controls and framework sheets section has the table.
- The Scenarios module is generated from the modules' scenario-ready inputs in Phase 1 and also holds profile days by scenario and the Monte Carlo bounds. Parity: Scenarios within scenarios moves to Phase 1, so 60 by Phase 1, 111 by Phase 2, 137 by Phase 3. Phase 1 gains a second test example, a small fictional operating budget with history, because seasonality does not apply to a development model.
- prototypes/charts: the Z chart now reproduces the template's Z-Chart picture element by element (budget bars behind and to the left, labelled lines instead of a legend with quarterly values, a divider at the last actual month, AC and FC under the months, YTD and FY variance boxes, green or red and aware of cost lines). Helpers placed between months are scatter series on the category axis and hide themselves with #N/A; every chart has "Show #N/A as an empty cell" on. The template's typed budget bars (74) and budget line (76) disagree, and its MAT ends at 55 against a year of 82, so the live chart differs there on purpose. LibreOffice draws bars from two axes side by side, so the overlap still needs Excel to confirm.

## [2026-10-06 15:04 NZ] CLAUDE | NOTE
Modano's module charts were not wired in; now designed and proven.
- Inventory: Modano's example (SEM_COA6) has 95 native charts, every one owned by one of nine summary and report modules (Income, Balance, Cash and Budget summaries; Income, Balance, Cash, Budget and Scenarios reports). Each chart sits in the module definition with series as reference templates and a position relative to the module, so inserting the module brings its charts. Kelvin's template kept two (a seasonality curve and an empty chart on the Income summary); BUD25 has 18, five empty.
- Design (spec, Charts and reports section): a charts section in the module definition, series by row key, "each" series that grow and shrink with categories, a top-N option, chart rows that read the timeline with INDEX from a first-period selector, and add_chart, set_chart and delete_chart in the change plan. HFG also gives some calculation modules a chart (Seasonality, Monte Carlo, Sensitivity, Covenants, facility and funding engine, 13-week cash, profile library). Summary and report modules stay in Phase 2.
- Proof: prototypes/assembly gains a demo Income summary module with a chart, keyed and shaped collect rows, span rows, new tokens ([range:row], {src_range}, {p}, {periods}), chart ops in the plan, charts in the package writer, and chart ops and chart snapshots in the LibreOffice live writer. hfgmodels.verify exposes the UNO context (uno_context). 6 new tests (17 in tests/test_assembly.py): live equals fresh in series, labels, categories, stacking and anchor after inserting the module, adding a revenue line and removing one; the window follows the first month shown.
