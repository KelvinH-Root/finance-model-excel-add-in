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

## [2026-10-06 15:56 NZ] KELVIN | DECISION
All of the charts in Modano's example model are rebuilt (all 95), not a leaner set chosen per summary.

## [2026-10-06 15:56 NZ] CLAUDE | NOTE
All 95 module charts rebuilt as a Phase 0 proof in prototypes/reports, and registered in the spec.
- prototypes/reports/register.yaml lists C01 to C95 (HFG module, title, recipe, what it reads): Income summary 5, Balance summary 6, Cash summary 5, Budget summary 6, Income report 28, Balance report 7, Cash report 12, Budget report 6, Scenario report 20. Nine recipes build them: compare, mix, depth, pie, combo, budget, scenario, bridge, movement. Category sets show in full or as top N plus Other, ranked by formula.
- build.py writes a fictional three-statement model (FY2025 to FY2028, actuals to September 2026, a budget and Base, Upside and Downside) and the nine module sheets, each with year shown and month shown drop-downs, a grid of native charts and the formula rows they read. Charts are named by register id; every chart shows #N/A as empty.
- Kept from Modano: chart types (pies stay pies), modules and counts. Changed on purpose: comparison periods line up with the months shown; the Scenario report's summary set is by year; waterfalls use four series; IBCS styling and the HF palette; HFG titles. The Modano mapping is in the spec's new Chart register tab only (not in git).
- tests/test_reports.py (5 tests): register and workbook parts agree; in LibreOffice no formula errors and no checks raised, statements and chart rows match the Python reference, and the year, month and scenario move every chart. Full suite 48 passed, 1 skipped. Not yet opened in Excel.

## [2026-10-06 15:56 NZ] CLAUDE | HANDOFF
For Kelvin: open the report charts demo (python prototypes/reports/build.py, or the copy sent in chat) in Excel desktop and on the web; change the year and month drop-downs on a report sheet and the active scenario on Scenarios, and check the hatched forecast bars, pies and titles. For any agent: add charts through register.yaml and a recipe, not by hand in build.py.

## [2026-10-06 16:30 NZ] KELVIN | DECISION
The contents is wired in Phase 1 as Modano's: links to every sheet, updated as sheets are created, with links back to the contents and the checks on every sheet. The four summary pages (Income, Balance, Cash, Budget; 22 charts) move to Phase 1; the five reports (73 charts) stay in Phase 2.

## [2026-10-06 16:30 NZ] CLAUDE | NOTE
Contents, section covers and links designed and proven; two example models built and screenshotted.
- Layout as Modano's: Contents first (model name with error count, notes, sections numbered 1, 2, 3 linking to covers, sheets lettered a., b., c., headings marked "-", error and alert totals); a cover sheet before each section, created with its first sheet and removed with its last; A1 to the contents and A2 to the checks (a tick while clear) on every other sheet. Links are HYPERLINK formulas to HL_ names; entries read each sheet's title cell.
- prototypes/assembly: areas.yaml has sections; covers, contents and names come from the layout, so the change plan carries them. Inserting the Income summary into a built model adds the Dashboards cover and sheet, renumbers the contents and moves cover links, live equal to fresh. 3 new tests (20).
- prototypes/models: navigation.py (the same layout for package-written workbooks, refresh files new or moved sheets where they sit), development.py (development test example finished from the template: frame sheets kept, other template sheets dropped, dashboard with four charts), screenshots.py. tests/test_models.py, 5 tests (development one needs HFG_TEMPLATE).
- prototypes/reports is now also the budget and actuals example: sections and covers, and a Seasonality sheet phasing the revenue budget from FY2025 and FY2026 actuals (include flags, override, even-spread fallback, checks). 2 new tests (7).
- Template findings: its Cover contents is typed by hand and has drifted (Seasonality links to HIS, missing links, Reports section missing, numbering jumps from 2 to 5, cover links not set), Lookups carries the Income Statement Summary title, and Time uses LET (LibreOffice 24.2 cannot calculate it).
- Spec updated: Contents, section covers and links subsection, framework table, phases (61 by Phase 1), exit criteria, How sure, phases diagram; parity tab dashboard row to Phase 1 and hyperlinks row.

## [2026-10-06 16:30 NZ] CLAUDE | HANDOFF
For Kelvin: open HFG_Budget_actuals_example.xlsx and HFG_Development_example.xlsx (sent in chat) in Excel desktop and on the web and click the contents entries, the cover links and A1 and A2 on a few sheets; confirm a HYPERLINK to a defined name jumps on both. For any agent: write navigation through navigation.py or the assembly layout, never by hand on a sheet.

## [2026-10-06 17:05 NZ] KELVIN | DECISION
Budgets are built and saved as in Modano, and monthly reforecasts are the comparisons in budget versus actual models; added as a feature since the spec did not have it.

## [2026-10-06 17:05 NZ] CLAUDE | NOTE
Saved versions (budgets and monthly reforecasts) designed and proven; Modano's consolidation researched.
- Finding: the spec's budget was live (typed annual amounts phased by seasonality) and nothing kept a reforecast. Modano's example (SEM_COA6 BUDIS, as in Kelvin's template) keeps one budget as typed values for 12 months beside a live block and a variance block, income statement only; its Budget Variance Analysis course runs budget creation, rolling forward, data import, reforecasting and variance analysis, with the live model as the reforecast, so last month's reforecast is lost on roll forward.
- Design (spec section Budgets, reforecasts and saved versions): a Version store (values, one row per version and line keyed id|line, every month) and a Versions register (type, label, year, as at, status, lock, source, checksum weighted by month); Save version, Approve budget (one approved budget per year), Lock or unlock, Compare and Manage versions in a Versions menu on the Analysis group (ribbon registry and manifest regenerated, 33 controls on HFG Model); Compared with on the budget modules (approved budget, last month's reforecast, latest reforecast, budget being built, any saved version); a Version comparison summary with four charts (C96 to C99, HFG additions). Phase 1: store, register, commands, Budget summary, Version comparison. Phase 2: Budget report, Z chart and budget waterfall, push to and load from Planning.
- Proof: prototypes/reports/versions.py. The budget and actuals example now carries two approved budgets and 18 monthly reforecasts (April 2025 to September 2026). tests/test_versions.py (5 tests) reads the store back, moves every comparison, and runs Save version live (roll forward to October and save, values hold when inputs change, an edited value is caught, a revised budget supersedes, replaying August equals the package writer). test_reports and test_models updated for the new sheets (99 charts). Full suite 64 passed with HFG_TEMPLATE; Node 14 pass, 1 skipped.
- Spec: new sections (Budgets, reforecasts and saved versions; Group consolidation), framework row, exit criteria, How sure row, new open decision; parity tab two rows (Better), so 145 Modano features (16 better), 63 by Phase 1, 113 by Phase 2, 139 by Phase 3; consolidation row reworded; parity, phases and ribbon widgets republished.
- Consolidation: Modano's Consolidations library has an entities summary, all-periods or historical consolidation, aggregation by categories, totals or reclassification, centralised or distributed eliminations, multiple currencies, entities in separate workbooks; also Multi-Business Units and a Stage Consolidation module. Bodies are behind a login; structure only.

## [2026-10-06 17:05 NZ] CLAUDE | QUESTION
Group consolidation in Excel: keep the add-in's consolidation to the Phase 2 sub-group what-if module that reads entity versions, with the group consolidation in Home Hub (recommended), or bring Modano's whole consolidation library into the add-in? On the spec's open decisions.

## [2026-10-06 17:05 NZ] CLAUDE | HANDOFF
For Kelvin: open HFG_Budget_actuals_example.xlsx (sent in chat) in Excel desktop and on the web: the Versions register, the Budget summary's Compared with drop-down (try Last month's reforecast and a named reforecast), and the Version comparison (line and month shown). For any agent: versions are written by Save version (versions.save_live here, the change plan in the add-in), never typed into the Version store; add comparison charts through register.yaml and versions.py.

## [2026-10-06 18:06 NZ] KELVIN | DECISION
The add-in carries Modano's entire consolidation feature set, applied to how HFG works: many intergroup transactions carry margin; at-cost on-charges pass through a netting account ("Intergroup AP/AR - [Entity]", one per counterparty in every entity) with the cost and the on-charge invoice both coded to it so it stays at zero; some costs are moved into a development LP's WIP.

## [2026-10-06 18:06 NZ] CLAUDE | NOTE
Group consolidation designed for the whole feature set and proven in Phase 0.
- Rules are Home Hub's (data model doc): a pair is eliminated in the lowest group holding both sides and every group above, a related party below; dated membership; NCI worked out once at each node and carried up, bearing its share of upstream margin, deficits allowed (PBE IPSAS 35 para 49); netting accounts carry balances only; matching compares AR and AP including GST, P&L excluding GST; plan lines carry projected margin, no deferred tax.
- Register types: trading expensed, trading capitalised, asset sale, at cost expensed, at cost capitalised, moved at cost, distribution, balance. Margin is tracked by group and site (still held, added, released, the account and holder each year), so margin capitalised into WIP follows the homes into investment property after a portfolio sale and is released as homes are sold outside the group.
- Proof: prototypes/consolidation (group.py: eleven fictional entities, five groups, two NCI nodes at 60% and 70%, FY2026 to FY2028, 33 register rows; build.py: an all-formula workbook with Entities, Accounts, Sites, Entity data, Intercompany, Investments, Unrealised margin, Eliminations, NCI, By group, Group statements, Checks). tests/test_consolidation.py (6 tests): every group matches an independent group-view reference for every account and year with NCI; checks catch a break and an unexplained netting balance. Full suite passes.
- Spec: Group consolidation section rewritten (patterns table, feature table, NCI, proof, open questions); open decision closed; exit criteria (Phase 2) and How sure row; parity tab 7 rows replace 2, so 150 Modano features (117 same, 17 better, 11 different, 5 not needed), 63 by Phase 1, 118 by Phase 2, 145 by Phase 3; parity and phases widgets republished; catalogue M14 rows and the Consolidation module updated.
- Not yet: dated membership, distributed eliminations, aggregation by totals or reclassification, currencies, reading entity versions, ribbon commands for consolidation, Excel itself.

## [2026-10-06 18:06 NZ] CLAUDE | QUESTION
For Kelvin and HFG's accountants before Phase 2: NZHF and Aldwins shares held and outside investors; Arbor Green's parent; NZHF's portfolios at cost or fair value (and depreciation); how margin is measured on KM and GP fees; whether development LPs capitalise intergroup interest; the accounting and GST treatment of portfolio sales between group entities; whether Home Hub stays the consolidation of record for actuals (assumed).

## [2026-10-06 18:06 NZ] CLAUDE | HANDOFF
For Kelvin: open HFG_Consolidation_example.xlsx (sent in chat) in Excel desktop and on the web; change the group and year on Group statements, and look at Intercompany, Unrealised margin and Checks. For any agent: consolidation logic follows the register types and Home Hub's rules; extend group.py's reference first, then the workbook, and keep the two independent.

## [2026-10-06 18:37 NZ] KELVIN | DECISION
A Group menu shows the layout of the group and its roll-up, with a way to add entities and their ownership: shown in the task pane and on a Group structure sheet in the model; entities come from Home Hub, with planned ones added in Excel; the roll-up shows each group's figures as well as its structure.

## [2026-10-06 18:37 NZ] CLAUDE | NOTE
Group menu, Group structure sheet and Add entity designed and proven.
- Ribbon: Group menu on the HFG Model tab's Model group (34 controls now): Group structure, Add entity, Change ownership, Remove entity, Refresh from Home Hub; Ctrl+Alt+Shift+G opens the structure. Manifest, shortcuts and icons regenerated; the manifest validates.
- Probe: src/group.js draws working views on a fictional sample group generated from the consolidation proof (tools/group-sample.py writes src/group-sample.js): the tree with share held, owned by the top, NCI and planned chips and a details card; the roll-up for any group and year (its members, sub-groups consolidated, eliminations made in it) with drill-down; Add entity with a live preview of where it goes and what the change plan writes; Change ownership as a preview. Nothing is written to the workbook. test/group.test.mjs (4 tests): every roll-up adds up to its group in every year; Add entity keeps tree order and makes a partly held entity a group.
- Proof: the Entities register gains Status (Actual or Planned) and Member from; investments carry the year they were made; a new Group structure sheet (tree, groups and where they roll up, surplus and net assets rolled up for the year shown, each column tied to By group); checks for tree order, figures before joining and the roll-up tie, and an alert for planned entities. group.add_entity places an entity as Add entity will. tests/test_consolidation.py now 9 tests: with Dev LP C (under Devco) and Fund Two LP (50% held, a new NCI node and group) added, every group still matches the reference. Full suite 73 passed with HFG_TEMPLATE; Node 18 pass, 1 skipped.
- Spec: Group consolidation section gains Group structure and adding entities; feature table, How sure and the parity row updated; ribbon widget republished.
- Not yet: changes of parent or share held from a date in the workbook (Change ownership is a preview), Remove entity and Refresh from Home Hub beyond their designs.

## [2026-10-06 18:37 NZ] CLAUDE | HANDOFF
For Kelvin: run the probe and open HFG Model > Model > Group (or Ctrl+Alt+Shift+G): click entities in the tree, open Roll-up and drill from Foundation into Holdings, and add an entity; open HFG_Consolidation_example.xlsx and look at Group structure (change the year at the top). For any agent: entities go in through group.add_entity (proof) or Add entity (add-in), never typed into the register out of tree order; regenerate src/group-sample.js with tools/group-sample.py when the proof's group changes.

## [2026-10-06 21:39 NZ] KELVIN | DECISION
The Explorer takes the shape of Modano's Project Manager (one pane: the tree, then Composition, Links, Properties and Checks, with the module selected in the workbook), and Impacts, HFG's take on Modano's Financial Statement Impacts Analyser, works on the current model: Impact of a change live in the pane by default, and Impacts sheets added to the model in its own accounts and entities, with a group view for intergroup items.

## [2026-10-06 21:39 NZ] CLAUDE | NOTE
Explorer and Impacts designed, proven and in the probe.
- Ribbon: Explorer (Model group), Build > Manage > Explorer and Modules > Links open the one Explorer pane; Analysis > Impacts menu (Impact of a change, Impacts sheets) and right-click > Show impacts (35 controls on HFG Model, 13 right-click items). Manifest, shortcuts and icons regenerated; the manifest validates.
- Probe: src/explorer.js (tree coloured by area with check status; Composition, Links as a click-through diagram, Properties, Checks; selects the module's HL_Toc_ name when Excel has the model open), src/impacts.js (Impact of a change with ties and the link chain; Impacts sheets in either context), src/model-calc.js (the assembly demo's calculation standing in for Excel), src/view-kit.js, src/model-sample.js from tools/model-sample.py. test/views.test.mjs (5 tests): the calc matches the assembly reference, relations and tree, ties for every input, panels equal the Python reference. Node 23 pass, 1 skipped.
- Proof: prototypes/impacts. impact_live.py: change a named input, recalculate, read the statement lines, restore; the chain from the link records. impact_sheets.py: 11 items (generic and HFG development, three intergroup with eliminations and group), offered by the model's accounts, written as formula sheets with switches and checks; the consolidation example now has an Impacts section and an Impacts error check. The assembly reference moved to prototypes/assembly/reference.py (test_assembly imports it). tests/test_impacts.py (10 tests). Full suite 83 passed with HFG_TEMPLATE.
- Spec: new Explorer and Impacts section; parity Explorer and Links rows reworded and a Financial Statement Impacts Analyser row (Better, Phase 2), so 151 features (117 same, 18 better, 11 different, 5 not needed), 63 by Phase 1, 119 by Phase 2, 146 by Phase 3; exit criteria and How sure; catalogue note; parity, phases and ribbon widgets republished.
- Not yet: dimming the rest of the workbook when a module is selected (to test in the probe), Save as an Impacts sheet for a live result, a combined Impacts sheet, Excel speed of the round trip.

## [2026-10-06 21:39 NZ] CLAUDE | HANDOFF
For Kelvin: in the probe open Model > Explorer (try Debtors, then click through the Links diagram) and Analysis > Impacts (change Revenue line 1's first month revenue; then Impacts sheets with the Demo Group accounts); open HFG_Consolidation_example.xlsx and look at the Impacts section. For any agent: Impacts items live in prototypes/impacts/impact_sheets.py with their reference; regenerate src/model-sample.js with tools/model-sample.py when the assembly demo or the items change.

## [2026-10-06 23:05 NZ] KELVIN | DECISION
Multiple currencies are not needed: HFG operates only in New Zealand. Every other feature on the parity list stays, including review and audit tools, printing and finalising, work in progress and issues, distributed eliminations and ownership changes from a date.

## [2026-10-06 23:05 NZ] CLAUDE | NOTE
The spec's Multiple currencies row moves to Not needed (151 features: 116 same, 18 better, 11 different, 6 not needed; 145 needed by Phase 3). The spec, its widgets and the project doc are updated together with any HFG additions Kelvin picks from the recommendations given in chat on 6 October 2026, so the widgets are republished once.

## [2026-10-07 08:37 NZ] KELVIN | DECISION
Eight HFG additions are adopted: a before-and-after check on every structural command, an input register, one set of group assumptions, release profiles (auditor, lender, board), model compare, a variance commentary workflow, GST return periods and due dates, and a speed check for adopted workbooks. They must be incorporated well, not bolted on.

## [2026-10-07 08:37 NZ] CLAUDE | DECISION
How the eight fit: three shared pieces in the model metadata, one gate, and hooks into existing flows rather than eight separate features.
- Key outputs: modules mark their key rows (closing cash, surplus, peak debt, covenant headroom, IRR); a recipe can add more.
- Input records: every named input's record gains source, owner, updated, evidence link and, where bound, the group assumption it uses.
- Change log: every structural command and every input changed through the add-in adds a line (who, when, what, key outputs before and after).
- Finalise is the gate: stale inputs, an out-of-date assumptions set and missing variance comments block or warn; release profiles copy from a finalised model.
- Placement: before-and-after has no button (every preview and result card); Data group gains Inputs and an Assumptions menu; the Versions menu gains Variance commentary; right-click gains Input details and Explain variance; Build > Review gains Compare and Speed check; Release copy becomes a menu of profiles; GST periods are library content.
- Phases: 1 for key outputs, change log, before-and-after, input register and group assumptions from a library file; 2 for assumptions from Home Hub, variance commentary, GST periods and speed check; 3 for release profiles and model compare.

## [2026-10-07 09:30 NZ] CLAUDE | NOTE
Model assurance designed, proven, in the ribbon registry and in the spec.
- Engine (prototypes/assembly): the Model assurance module (library/assurance.yaml, framework: assurance) brings a Group assumptions sheet and an Input register sheet and puts the key outputs (KO_ names) on the contents; module rows mark key outputs with `headline`; a setting can be bound to a group assumption (LRow.link, kept in the signature, so binding, unbinding and set updates go through the plan); view settings (`display: true`) stay out of the register; the uno dialect now leaves commas inside quoted text alone. Models without the module lay out exactly as before.
- Proof (prototypes/assurance, tests/test_assurance.py, 20 tests): guard.py (what a change can reach, from every formula's markers, ranges and names; key outputs read before and after; unexplained movements; the change log in the metadata), register.py (records, read back, sets v3 and v4, latest version, statuses), compare.py, release.py (auditor, lender, board), gst.py (due dates and cash timing). Full suite 103 passed with HFG_TEMPLATE; Node 24 pass, 1 skipped; the manifest validates.
- Ribbon: Data gains Inputs and an Assumptions menu; Versions gains Variance commentary; right-click gains Input details and Explain variance (15 items); Build > Review gains Compare and Speed check; Release copy becomes a menu (Auditor, Lender, Board, Manage profiles). 37 controls on HFG Model, 36 on HFG Build. Probe: a speed check probe (worksheet.calculate per sheet, ExcelApi 1.14, and a formula scan with a Node test).
- Spec: new Model assurance section with a diagram; metadata table, How sure, exit criteria for Phases 1 to 3, open decisions (group set owner, assurance settings), catalogue GST row and confirmations; parity tab: Multiple currencies Not needed and eight HFG additions, so 151 features (116 same, 18 better, 11 different, 6 not needed) plus 12 HFG additions, 63, 119 and 145 by Phases 1 to 3; parity, phases and ribbon widgets republished.
- Not yet: variance commentary built (designed; its store mirrors the Version store), the Inputs, Assumptions, Changes and Compare panes in the probe, Excel itself.

## [2026-10-07 09:30 NZ] CLAUDE | HANDOFF
For Kelvin: look at HFG_Model_assurance_screenshots.pdf and the new menus (both sent in chat); when you run the probe, open BUD25 or the template first and run the speed check; decide who publishes the group assumptions set, the review age (12 months proposed) and commentary threshold, and what each release profile keeps. For any agent: a new assurance feature hangs off the three shared pieces (key outputs, input records, change log) and the Finalise gate; extend register.py or guard.py's reference first and keep the tests comparing live with fresh builds.
