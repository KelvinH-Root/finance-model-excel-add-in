# Model assurance proof (Phase 0)

The eight additions Kelvin adopted on 7 October 2026, built as one system (seven are new; GST timing adapts the GST payments Modano's tax module already makes to New Zealand): three shared pieces in each model's metadata, one gate at Finalise, and hooks into commands the add-in already has. Multiple currencies were dropped the same day (HFG operates only in New Zealand).

```
pytest tests/test_assurance.py            # 20 tests; the LibreOffice ones skip without it
python prototypes/assurance/demo.py       # takes the demo through the commands and prints what the add-in shows
```

## The shared pieces

| Piece | Where it lives | Used by |
|---|---|---|
| Key outputs | Module rows marked `headline` with a measure (sum, last, min, max); KO_ names at the foot of the contents | Before-and-after check, model compare, release copies |
| Input records | `model.assurance["inputs"]` in the metadata; shown and typed on the Input register sheet, read back before every change | Input register, group assumptions, model compare, release copies |
| Change log | `model.assurance["log"]` in the metadata | Before-and-after check, model compare (the commands in between), the Explorer's Changes tab |

The engine (`prototypes/assembly/assemble.py`) writes the Group assumptions and Input register sheets, and the key outputs on the contents, when the Model assurance module (`library/assurance.yaml`, `framework: assurance`) is in a model. Its four alerts feed the model's own Checks. A setting bound to a group assumption is written as a link (`=GA_LendingRate`, or a formula such as `=(1+GA_Cpi)^(1/12)-1`) and kept in the row's signature, so binding, unbinding and set updates go through the ordinary change plan. View settings (`display: true`, such as the month a chart starts on) stay out of the register.

## Files

| File | What it proves |
|---|---|
| `guard.py` | The before-and-after check: what a change can reach is worked out from the layout (every formula's markers, ranges and defined names), the key outputs are read before and after the plan is applied, and anything that moved outside that set, or at all for a command that should not move numbers, is unexplained. Each command is logged. |
| `register.py` | Input records, reading the register back, group assumption sets (`sets/hfg-group-v3.yaml`, `v4`), the latest version known, the status each register row should show, and the settings a model resolves to for the reference. |
| `compare.py` | Model compare by module: modules, rows added, removed or rewired, inputs changed (read from the files), bindings, records, the set, key outputs and the change log in between. |
| `release.py` | Release profiles (auditor, lender, board): only the sheets the results need, a contents with the key outputs, the register as values without internal columns, values only where set, no navigation, subtitles or metadata. |
| `gst.py` | What New Zealand adds to the GST payments in Modano's tax module: two-monthly filing, the due dates (the 28th of the next month, 15 January for November, 7 May for March) and refunds after a lag, as formulas with a Python reference. |
| `demo.py` | Writes the workbooks to `build/assurance` (kept out of git). |

## Tests

`tests/test_assurance.py`: every input appears once in the register and view settings are left out; what a change can reach is right for a revenue line, a facility, a summary and a set update; the GST rule; bound settings resolve. In LibreOffice, a chain of commands runs the way the add-in will: a summary insert moves nothing, a revenue insert moves only what it reaches, a planted fault is caught, the log is kept, every result equals a fresh build, a source typed on the register survives, statuses, counts and alerts match Python, a newer set raises an alert until the update, the statements match the reference with the group set, model compare lists what changed, release copies keep only what the recipient needs and agree with the model, and GST timing matches the reference for five filing patterns.

## Not proven here

Variance commentary (designed; its store mirrors the Version store proven in `prototypes/reports`), the Inputs, Assumptions, Changes and Compare panes, the speed check's long-chain measure (the probe times sheets and scans formulas in Excel), and Excel itself.
