# Impacts proof (Phase 0)

The add-in's Impacts command (Analysis > Impacts, and right-click > Show impacts): HFG's own take on Modano's Financial Statement Impacts Analyser. Modano's analyser is a separate workbook of hand-built examples; this one works on the model in front of you. Kelvin agreed the design on 6 October 2026.

```
pytest tests/test_impacts.py        # 10 tests; the LibreOffice ones skip without it
python prototypes/consolidation/build.py   # the consolidation example, with its Impacts section
```

## Two modes

**Impact of a change (live), `impact_live.py`.** Pick an input, a new value and a month. The add-in writes the value, recalculates, reads every statement line, and puts the value back, all in one step, so the model is never left changed. It shows each line that moved, the ties (surplus carried to the accumulated surplus, net cash flow to cash, the balance sheet still balancing) and the chain of links that carried the change, read from the link records the engine keeps (Revenue line 1 sends revenue to Debtors, Debtors sends receipts and closing debtors to the statements).

Proven on a model built by the assembly engine (prototypes/assembly), through LibreOffice: for five inputs (revenue, cost, facility amount, growth and rate) every statement total moves exactly as the assembly reference says with that input changed, the ties hold, and every formula and value in the workbook is the same afterwards as before.

**Impacts sheets (Add to model), `impact_sheets.py`.** Writes an Impacts section into the model: one formula sheet per kind of transaction the model holds, chosen by the accounts it has and labelled with its own account names and entities. Each sheet has its inputs and switches (GST registered, paid in the period, capitalised), the entries (debit positive, by entity, with each cash entry's cash flow class), and the effect on the income statement, balance sheet and cash flow, with the surplus carried to retained surplus, net cash flow equal to the change in cash and a balance check. Intergroup items show the seller, the buyer, the eliminations and the group.

| Item | Needs | HFG point it shows |
|---|---|---|
| Sale on credit, operating cost, facility drawn and repaid, interest paid, equity injection | revenue, costs, loans, capital | The generic set, for any model |
| Land purchase | WIP | Land for a development is inventory: an operating cash flow |
| Construction cost into WIP | WIP, GST | GST claimed back when registered, capitalised when not |
| Interest capitalised | WIP, loans | Into WIP, not the income statement; added to the loan it is not a cash flow |
| Homes sold to the Fund | WIP, investment property, sales | The sale and margin come out in the group; homes at group cost; no group cash flow |
| Construction claim with margin | WIP, construction revenue and costs | The builder's margin comes out of the LP's WIP in the group |
| On-charge at cost | the netting account | The netting account stays at nil; the receiver expenses or capitalises |

In the consolidation example the section holds 11 sheets in the group chart's accounts (1200 Development work in progress, 2050 Intergroup AP/AR clearing) and its entities (9001 BuildCo Ltd, 9006 Dev LP A, 9008 Housing Fund LP); the assembly demo is offered the 4 generic items its accounts support.

## Tests

`tests/test_impacts.py`: the items offered follow the model's accounts; every item balances and ties for every combination of its switches, and intergroup items earn nothing inside the group; in LibreOffice every Impacts sheet matches the Python reference for its defaults, each switch and a doubled input, and typing over an entry raises the check; the live round trip matches the reference and leaves the workbook as it was; the chain names the links.

## Not proven here

Excel itself (the live round trip runs through LibreOffice; the probe's one-step undo test covers part of Office.js's side), speed on a large model, Save as an Impacts sheet for a live result (values with their inputs and a stale flag, as Freeze results), and a combined sheet adding every item.
