# HFG model engine

The engine the add-in and the command-line build share: the module library, link resolution, layout, rendering of formulas in Excel's and LibreOffice's dialects, and the change plan between two layouts. TypeScript, run directly by Node (22.18 or later strips the types), with `tsc` for type checks only.

```
npm install
npm test            # golden and behaviour tests
npm run typecheck
npm run golden      # rewrite test/fixtures from the Python proof (needs Python and PyYAML)
node tools/build.ts build/demo.xlsx --brand HF     # the demo model in HFG's standard frame
python tools/screenshots.py build/demo.xlsx        # screenshots of it (LibreOffice)
```

## How it is checked

It starts as a faithful port of the Phase 0 assembly proof (`prototypes/assembly/assemble.py`, everything except the openpyxl writer and the metadata part). `tools/golden.py` runs eleven scenarios through the proof (a base build, inserts of categories, a facility and a dashboard, a removal, a changed setting, a model with Model assurance, a newer group assumptions set noted and adopted, a rate unbound, an insert into an assured model) and writes, for each, the models before and after, the layout, every rendered cell in both dialects and the change plan. `test/golden.test.ts` must reproduce all of it exactly.

When the proof changes, regenerate the fixtures and change the engine to match in the same commit.

## Two frames

A model without New model's choices (`Model.info`) is laid out in the proof's frame, so the golden tests compare like with like. A model with them is laid out in HFG's standard frame (docs/frame-standard.md, the spec's Look and wiring tab): a Settings sheet that holds the model, timeline and display choices and works out the timeline block every calculation sheet reads; rows 1 to 3 header (sheet title, the model name line with error and alert counts, the entity); the timeline block in rows 5 to 15, content from row 17; the contents with the entity's name and logo, notes and a grouped table of contents; section covers; cell hyperlinks to `Go_` names with screen tips; spacer rows of 3, 6 and 9 points; a dashed rule under a list's last item, plain list totals and bold major results; checks that turn bold red.

Names use HFG's own prefixes: `Go_` navigation, `Tl_` timeline settings, `Opt_` switches, `Model_` the model's own lines, `Chk_` check totals, with `GA_`, `Reg_` and `KO_` from Model assurance and module codes for inputs (`Rev1_Base`).

## Package writer

`buildWorkbook(layout, model, { logo })` writes the whole .xlsx: sheets, the HFG style catalogue (`HFG Heading 1`, `HFG Input Number` and so on, every colour a theme slot), the entity's theme (`theme.ts`: accent 1 the signature colour, section bar text chosen by contrast), defined names, outline levels, row heights, frozen panes, hyperlinks, conditional formats, validations, module charts, the logo and the model metadata part. It runs in Node and in the browser (no Node modules outside `src/node/`). `tests/test_engine_writer.py` checks the proof frame against the Python writer cell for cell, name for name and series for series, and recalculates every entity's build in LibreOffice against the reference calculation.

Not yet: formats in the change plan for the live writer, classic form controls (switches are validated TRUE or FALSE cells for now), the conditional formats that follow time series inputs, page setup beyond landscape A4.

## Files

| File | What it holds |
|---|---|
| `src/frame.ts` | Frame columns and rows, the proof and standard frames, name helpers, `AssemblyError` |
| `src/library.ts` | Module definitions and areas, from a bundle (`src/node/library.ts` reads the YAML folder) |
| `src/model.ts` | A model's instances, settings, bindings to group assumptions and assurance state |
| `src/resolve.ts` | Blocks and link producers, to a fixed point |
| `src/assemble.ts` | Rows, marker formulas, names, link records, charts and key outputs |
| `src/navigate.ts` | Sections, covers, contents and navigation names |
| `src/assurance.ts` | Group assumptions and Input register sheets |
| `src/render.ts` | Markers to cell formulas, row and frame cells, chart ranges |
| `src/plan.ts` | The change plan and its preview |
| `src/standard.ts` | HFG's standard frame: Settings, timeline block, contents, covers, links, the model name line |
| `src/theme.ts` | Entity themes and the contrast rule for section bars |
| `src/styles.ts` | The HFG style catalogue and the styles.xml it writes |
| `src/xlsx/` | Package writer: dressing rows into cells and formats, worksheets, charts, the package |
| `src/node/` | Node only: reading the library folder and the entity logos (`assets/logos`) |
