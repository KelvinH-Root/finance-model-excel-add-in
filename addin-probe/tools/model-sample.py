"""Write src/model-sample.js: the model the probe's Explorer and Impacts views show.

    python tools/model-sample.py

The model is the assembly proof's demo (prototypes/assembly), built with its Income summary:
sections, sheets, modules, their rows, settings and checks, and the link records the engine
keeps. The Impacts items and their expected effects come from prototypes/impacts, for the demo
model's own accounts and for the consolidation example's group chart and entities. Fictional
data only; the Node tests compare the pane's calculations with the references written here.
"""
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
PROTO = HERE.parent / "prototypes"
for sub in ("impacts", "assembly", "consolidation"):
    sys.path.insert(0, str(PROTO / sub))
from assemble import Library, assemble  # noqa: E402
from demo import base_model  # noqa: E402
from reference import reference  # noqa: E402
import impact_live as LV  # noqa: E402
import impact_sheets as IS  # noqa: E402
import group as G  # noqa: E402

lib = Library.load()
model = base_model(lib)
model.insert("demo.dashboard", first=1)
layout = assemble(model)
pos = layout.positions()
area_of = {m["id"]: m["area"] for m in lib.modules.values()}

sections = []
for sec in lib.sections:
    sheets = [a for a in sec["areas"] if any(s == a for s, _ in layout.sheets)]
    if sheets:
        sections.append({"title": sec["title"], "cover": sec["cover"], "sheets": sheets})

modules = []
for inst in model.instances:
    uid = inst.uid
    mod = lib.modules[inst.module]
    blocks = [b for b in layout.blocks if b == uid or b.startswith(uid + "/")]
    comps = {}
    checks = []
    for sheet, rows in layout.sheets:
        for r in rows:
            if r.id == uid or r.id.startswith(uid + "/"):
                if r.kind in ("blank",):
                    continue
                comps.setdefault(sheet, []).append({"id": r.id, "label": r.label, "kind": r.kind, "row": pos[r.id][1],
                                                    "style": r.style, "name": r.name})
                if r.style == "check":
                    spec = next((x for x in mod.get("rows", []) if x.get("key") == r.id.rsplit("/", 1)[1]), {})
                    checks.append({"label": r.label, "kind": spec.get("check", "error"), "id": r.id})
    settings = [{"name": n, "key": rid.rsplit("/", 1)[1].removeprefix("set."), "label": next(r.label for _, rows in layout.sheets
                 for r in rows if r.id == rid), "value": inst.settings[rid.rsplit("/", 1)[1].removeprefix("set.")]}
                for n, rid in layout.names.items() if rid.startswith(uid + "/set.")]
    modules.append({"uid": uid, "module": inst.module, "title": model.title(inst), "area": area_of[inst.module],
                    "kind": lib.kind(inst.module), "number": inst.number, "blocks": [{"id": b, "title": layout.blocks[b]} for b in blocks],
                    "settings": settings, "components": [{"sheet": s, "rows": rs} for s, rs in comps.items()], "checks": checks,
                    "note": mod.get("note", "")})

links = []
for rec in layout.records:
    src, dst = LV.block_of(layout, rec["from"]), LV.block_of(layout, rec["to"])
    e = {"from": src, "link": rec["link"], "to": dst, "mode": rec["mode"]}
    if e not in links:
        links.append(e)

lines = [{"id": ln.id, "label": ln.label, "section": ln.section, "total": ln.total} for ln in LV.statement_lines(layout)]
base = reference(model)
cases = []
for name, value in (("Rev1_Base", 110.0), ("Fac1_Amount", 1500.0)):
    uid, key = LV.setting_of(layout, name)
    m2 = model.copy()
    next(i for i in m2.instances if i.uid == uid).settings[key] = value
    cases.append({"name": name, "value": value, "reference": reference(m2)})

model_sample = {"name": "Assembly demo", "periods": model.periods, "areas": lib.areas, "sections": sections,
                "modules": modules, "blocks": layout.blocks, "links": links, "lines": lines,
                "reference": {"base": base, "cases": cases}}


def item_json(item, ctx):
    return {"key": item.key, "title": item.title, "purpose": item.purpose, "sheet": item.sheet,
            "inputs": [{"name": n, "label": lab, "value": IS.defaults(item, ctx)[n]} for n, lab, _ in item.inputs],
            "switches": [{"name": n, "label": lab, "value": d} for n, lab, d in item.switches],
            "legs": [{"entity": IS.entity_label(ctx, item, s) if s != IS.ELIM else IS.ELIM, "role": r, "expr": e, "cash": c}
                     for s, r, e, c in item.legs],
            "columns": IS.columns(item, ctx), "group": item.group, "notes": item.notes, "lines": IS.lines(item, ctx),
            "expected": IS.panels(item, ctx)}


contexts = []
for label, ctx in (("This model (assembly demo)", IS.from_assembly(layout, model)),
                   ("Demo Group (consolidation example)", IS.from_consolidation(G))):
    contexts.append({"label": label, "model": ctx.model, "accounts": {k: list(v) for k, v in ctx.accounts.items()},
                     "items": [item_json(i, ctx) for i in IS.offered(ctx)]})
impacts_sample = {"contexts": contexts, "fixed": IS.FIXED, "cashClasses": list(IS.CASH_CLASSES)}

out = HERE / "src" / "model-sample.js"
out.write_text("/* Generated by tools/model-sample.py from prototypes/assembly and prototypes/impacts: do not edit by hand. */\n"
               "(function (root) {\n  root.HfgModelSample = " + json.dumps(model_sample, ensure_ascii=False) + ";\n"
               "  root.HfgImpactsSample = " + json.dumps(impacts_sample, ensure_ascii=False) + ";\n"
               "})(typeof globalThis !== 'undefined' ? globalThis : this);\n")
print(out, out.stat().st_size, "bytes")
