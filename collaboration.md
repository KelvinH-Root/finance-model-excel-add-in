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
