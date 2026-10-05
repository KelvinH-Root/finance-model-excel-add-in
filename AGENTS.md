# Instructions for coding agents (Codex, Claude)

## Collaboration board

Three parties work on this project: Kelvin (owner, final say), Claude, and Codex.

`collaboration.md` at the repo root is the shared message board.

Read `collaboration.md` before forming any response, at the start of every turn.
It holds the current state of the project and may have changed since your last
turn, including entries written by Claude. Do not rely on your recollection of
it from earlier in the session.

Append an entry when you make or change a decision, hit something that redirects
the work, disagree with an entry already on the board, or finish work Claude or
Kelvin will pick up. Do not append for routine chatter.

Never edit or delete another author's entries. Append only, newest at the bottom.

Entry format:

    ## [YYYY-MM-DD HH:MM NZ] CODEX | DECISION
    One line summary.
    Detail, reasoning, and anything the next reader needs.

Types: DECISION, QUESTION, BLOCKER, HANDOFF, DISAGREE, NOTE.

Where you disagree with Claude, write a DISAGREE entry giving your position and
the reasoning, and leave the call to Kelvin. Kelvin has the final say and records
it as a DECISION. Do not replace Claude's approach without an entry explaining why.

No em dashes in anything you write.

## Repository rules

- Never commit workbooks (`.xlsx`, `.xlsm`, `.xlsb`) or real financial data. Recipes use fictional demo data only.
- Clean room: do not copy content, formula templates, module names or icons from Modano or other commercial products. Use them for ideas only.
- The product is the add-in. Agreed scope lives in the requirements spec (a Claude Doc in Kelvin's Technical Roadmap project); do not start add-in features before it is agreed.
- Nothing in `hfgmodels/` may be specific to one model. Model-specific logic belongs in `examples/` or in library modules.
- Every formula the engine writes must be reproduced by a reference calculation. For the test example that is `examples/development/reference.py`. Change both together, then run `pytest`. With the template available (`HFG_TEMPLATE`), `python -m hfgmodels build-example --verify` must report no mismatches and no model errors.
- Follow the frame standard in `docs/frame-standard.md`: labels B to G, units H, totals I, timeline from J, rows 1 to 3 header, rows 5 to 15 timeline block, row heights 15, named ranges for every scalar input and derived value, check rows for anything that can go wrong.
- Never edit an existing template part unless the change is deliberate and documented (for example the Time sheet term). Add new parts instead.
- Use only functions that work in Excel 365 and LibreOffice. Prefix newer functions as Excel stores them (`_xlfn.BETA.DIST`). Avoid volatile functions (INDIRECT, OFFSET, TODAY) in model logic.
- Writing style for anything Kelvin reads: New Zealand spelling, plain and direct, dates as d Month yyyy, no em dashes.
- Fonts and formats come from the template's cell formats (Segoe UI, 9pt body, 10pt headings, #404040 text). Do not hard-code new colours; derive from the template's theme.
