# HFG model engine

The engine the add-in and the command-line build share: the module library, link resolution, layout, rendering of formulas in Excel's and LibreOffice's dialects, and the change plan between two layouts. TypeScript, run directly by Node (22.18 or later strips the types), with `tsc` for type checks only.

```
npm install
npm test            # golden and behaviour tests
npm run typecheck
npm run golden      # rewrite test/fixtures from the Python proof (needs Python and PyYAML)
```

## How it is checked

It starts as a faithful port of the Phase 0 assembly proof (`prototypes/assembly/assemble.py`, everything except the openpyxl writer and the metadata part). `tools/golden.py` runs eleven scenarios through the proof (a base build, inserts of categories, a facility and a dashboard, a removal, a changed setting, a model with Model assurance, a newer group assumptions set noted and adopted, a rate unbound, an insert into an assured model) and writes, for each, the models before and after, the layout, every rendered cell in both dialects and the change plan. `test/golden.test.ts` must reproduce all of it exactly.

When the proof changes, regenerate the fixtures and change the engine to match in the same commit. Once the engine moves past the proof (the Look and wiring frame, the package writer, the live writer), its own tests take over and the fixtures are frozen.

## Files

| File | What it holds |
|---|---|
| `src/frame.ts` | Frame columns and rows, name helpers, `AssemblyError` |
| `src/library.ts` | Module definitions and areas from YAML |
| `src/model.ts` | A model's instances, settings, bindings to group assumptions and assurance state |
| `src/resolve.ts` | Blocks and link producers, to a fixed point |
| `src/assemble.ts` | Rows, marker formulas, names, link records, charts and key outputs |
| `src/navigate.ts` | Sections, covers, contents and navigation names |
| `src/assurance.ts` | Group assumptions and Input register sheets |
| `src/render.ts` | Markers to cell formulas, row and frame cells, chart ranges |
| `src/plan.ts` | The change plan and its preview |
