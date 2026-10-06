# Group consolidation proof

Phase 0 proof for the spec's consolidation feature set. Kelvin (6 October 2026): the add-in carries the whole set, applied to how HFG works: many intergroup transactions carry margin, at-cost on-charges pass through the "Intergroup AP/AR - [Entity]" netting accounts, and some costs are moved into a development LP's work in progress.

```
python prototypes/consolidation/build.py [out.xlsx]      # default build/consolidation/consolidation_demo.xlsx (not in git)
python prototypes/consolidation/group.py                 # the reference, printed by group
pytest tests/test_consolidation.py                       # 6 tests; the LibreOffice ones skip without it
```

## The group

Eleven fictional entities in HFG's shape, three years (FY2026 to FY2028), $000:

| Code | Entity | Plays the part of |
|---|---|---|
| 9000 | Foundation | Parent; pays shared costs and on-charges them at cost |
| 9001 | BuildCo Ltd | The builder: construction claims with a margin |
| 9002 | DevManager Ltd | Development management fees with a margin |
| 9003 | PropManager Ltd | Property management fee plus GST to the Fund |
| 9004 | Holdings LP | Sub-group head; due diligence WIP, loans to LPs with interest |
| 9005 | Devco LP | Sub-group head for the development LPs |
| 9006, 9007 | Dev LP A, Dev LP B | Build sites A and B; A sells its portfolio to the Fund |
| 9008 | Housing Fund LP | 60% held, 40% outside investors; buys portfolios, sells five homes outside the group |
| 9009 | Fund Dev LP | Builds site F for the Fund |
| 9010 | Partner LP | 70% held; sells site P to the Fund |

Groups that consolidate: Foundation, Holdings, Devco, Fund, and Partner LP as an NCI node.

## The workbook

Inputs are what Home Hub (or each entity's saved version) hands over: Entities (tree, shares held, GST, NCI nodes), Accounts (group chart), Sites (who holds each asset and how many homes remain), Entity data (each trial balance by group account and year) and Intercompany (the register from matching, both sides of every pair). Everything else is formulas:

- **Groups.** Each entity's groups are worked out from the tree; each register row is tagged with the lowest group holding both sides, eliminated there and in every group above, a related party below.
- **Eliminations** (three lines a row from the register, by type): trading the buyer expenses comes out in full; trading the buyer capitalises comes out at the seller's cost; a portfolio sale comes out against cost of sale; at-cost pass-throughs and moves at cost eliminate balances only; distributions against distribution income; balance pairs against each other, with any difference to 2900 and flagged.
- **Unrealised margin** by group and site: margin counts in a group when both sides and the asset's holder are in it; it comes out of the asset where it sits at the year end (WIP, then investment property after a portfolio sale), is carried forward through retained surplus, and is released to cost of sale as homes are sold outside the group or the asset leaves the group.
- **Investments** against capital, the outside investors' share to NCI. **NCI** is worked out once at each node from the node's own consolidation and carried up; upstream margin (from a partly owned seller) takes the node's NCI share.
- **By group:** combined, eliminations and consolidated trial balance for every group and year, with a summary that must be nil. **Group statements:** income statement, balance sheet, the entities in the group and its related parties, for any group and year.
- **Checks:** trial balances nil; every pair has a group; both sides agree; no intercompany difference left; each netting account balance is an on-charge in the register; investments are the parent's share of capital; eliminations and every consolidated trial balance balance; NCI ties to the nodes; the top group carries no investments. Alerts: an on-charge accrued at the cut-off, GST a group entity cannot claim (a real group cost).

## How HFG's cases come out

| Case | Foundation group | Holdings group | Devco group |
|---|---|---|---|
| Shared costs on-charged at cost through the netting accounts, one share into Dev LP A's WIP | Balances only; the cost sits once (expense or WIP) | Related party | Related party |
| BuildCo claims and DevManager fees capitalised into WIP | Margin out of WIP, then out of the Fund's property after the sale | Related party, at price | Related party |
| Interest on the Holdings loan capitalised into Dev LP B's WIP | Out of WIP | Out of WIP | Related party, at price |
| Due diligence moved from Holdings to Dev LP B at cost | Balances only | Balances only | Related party |
| Dev LP A's portfolio sold to the Fund | Sale and margin out; the Fund's property at group cost | Same | A real sale (the Fund is outside) |
| Partner LP's portfolio sold to the Fund | As above; Partner LP's outside investors bear 30% of the margin while the homes are held | Same | Not in the group |
| Five of site A's homes sold outside the group | A quarter of every margin on site A released | A quarter of the sale margin released | |
| PropManager's fee plus GST to the Fund (not GST registered) | Fee against expense; the GST stays as a group cost | Related party | |

## Tests

`tests/test_consolidation.py` (6 tests). The reference (`group.py`) re-records every event as each group sees it, tracks each site's group cost, and works out NCI from the nodes' net assets; it never reads the register. With LibreOffice, every group's consolidated accounts match it for every asset, liability, revenue and expense account and every year, with NCI and its share of surplus; the HFG cases above come out as stated; the statements follow the group and year shown; and a break between two sides, or a netting account balance with no on-charge, raises its check.

## Not proven here

- Dated group membership (an LP moving into the Fund from a date), multiple currencies, distributed eliminations, aggregation by totals or reclassification, and consolidating entity workbooks through their saved versions: in the spec, not in this proof.
- Plans: intergroup plan lines with projected margin, eliminated per scenario.
- Excel itself: recalculated in LibreOffice only.
- Accounting policy: margin here stays out of investment property while the homes are held (cost model). If NZHF measures its portfolios at fair value, or depreciates buildings, the release changes; that is for HFG's accountants and auditors.
