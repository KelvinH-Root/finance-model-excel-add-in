"""The group consolidation demo: the fictional group in story.py as a recipe the engine builds.

    python -m examples.consolidation.recipe      # writes library/hfg/recipes/consolidation_model.json

Everything here is fictional demo data ($000), shaped like HFG's group: a foundation at the top, a
holdings sub-group with development LPs, a housing fund and a partner LP that outside investors
part own. The recipe carries what Home Hub will hand over for a real group: the entity tree, each
entity's opening balance and monthly movements by group account, the intercompany register and
the intergroup balance pairs (the matching's output), the investments and the sites. The timeline
is the model's own (its settings), not part of the model type.
"""

from __future__ import annotations

import json
from pathlib import Path

from examples.consolidation import story

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "library" / "hfg" / "recipes" / "consolidation_model.json"
LAST_ACTUAL = 18     # September 2026: the months after it are the plan


def r9(x: float) -> float:
    v = round(x, 9)
    return int(v) if v == int(v) else v


def group_data(b: story.Book) -> dict:
    groups = dict(story.GROUPS)
    roles = {e[0]: e[5] for e in story.ENTITIES}
    entities = [{"code": c, "name": story.NAME[c], "parent": story.PARENT[c], "held": story.HELD[c], "gst": story.GST_REG[c],
                 "status": "Actual", "from": 1, **({"group": groups[c]} if c in groups else {}), "role": roles[c]}
                for c in story.CODES]
    sites = [{"code": k, "name": s["name"], "units": s["units"], "holder": s["holder"], "account": s["account"], "held": s["held"]}
             for k, s in story.SITES.items()]
    investments = [{"parent": p, "sub": s, "cost": r9(c), "period": t} for (p, s, c, t) in b.investments]
    transactions = [{"period": x["period"], "type": x["type"], "seller": x["seller"], "buyer": x["buyer"], "site": x["site"],
                     "sellerAcct": x["seller_acct"], "buyerAcct": x["buyer_acct"], "costAcct": x["cost_acct"],
                     "sellerAmt": r9(x["seller_amt"]), "buyerAmt": r9(x["buyer_amt"]), "gst": r9(x["gst"]),
                     "margin": r9(x["margin"]), "desc": x["desc"]} for x in b.register]
    balances = [{"seller": x["seller"], "buyer": x["buyer"], "sellerAcct": x["seller_acct"], "buyerAcct": x["buyer_acct"],
                 "sellerOpening": 0, "buyerOpening": 0, "sellerBal": [r9(v) for v in x["seller_bal"]],
                 "buyerBal": [r9(v) for v in x["buyer_bal"]], "desc": x["desc"]} for x in b.balances]
    accounts = sorted({a for (_, _, a, _, _, _) in b.lines})
    figures = []
    for (e, a), (opening, values) in story.figures(b, accounts).items():
        if any(abs(v) > 1e-9 for v in values) or abs(opening) > 1e-9:
            figures.append({"entity": e, "account": a, "opening": r9(opening), "values": [r9(v) for v in values]})
    return {"entities": entities, "sites": sites, "investments": investments, "transactions": transactions,
            "balances": balances, "figures": figures}


def recipe() -> dict:
    b = story.book()
    return {
        "title": "Group consolidation",
        "entity": {"name": "Home Foundation Group (fictional demo)", "brand": "HF"},
        "preparedBy": "Prepared by Group Finance",
        "notes": [
            "Fictional demo data ($000) shaped like the group: a foundation, a holdings sub-group with development LPs, "
            "a housing fund and a partner LP that outside investors part own.",
            "Entity figures, the intercompany register and the intergroup balances stand in for what Home Hub hands over; "
            "every elimination and every group's statements are formulas.",
            "Choose the group shown on Group statements; By group holds every group's consolidation month by month.",
        ],
        "timeline": {"start": f"{story.START[0]}-{story.START[1]:02d}", "fyEndMonth": story.FY_END, "lastActual": LAST_ACTUAL,
                     "denomination": "$000", "budget": {"first": 25, "months": 12}},
        "periods": story.T,
        "instances": [
            {"module": "fm.consolidation", "settings": {"group": 1}, "data": {"group": group_data(b)}},
            {"module": "fm.checks"},
        ],
    }


if __name__ == "__main__":
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(recipe(), indent=1) + "\n")
    print(f"wrote {OUT.relative_to(ROOT)}")
