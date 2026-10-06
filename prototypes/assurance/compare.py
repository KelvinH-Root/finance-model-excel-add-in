"""Model compare: what changed between two versions of a model, by module rather than by cell.

Both files carry their metadata, so the engine rebuilds both layouts and compares like with like
even after rows have moved: modules added and removed, rows added, removed or rewired in each
module, inputs whose values changed (read from the files, since people type over inputs), inputs
moved onto or off the group assumptions, input records edited, the group assumptions set, the key
outputs, and the change log entries made in between when one file descends from the other.
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "assembly"))
sys.path.insert(0, str(HERE))
from assemble import (GROUP_SHEET, REGISTER_SHEET, Library, Model, assemble, read_metadata)  # noqa: E402

FRAME = {"contents": "Contents and navigation", "cover": "Contents and navigation", "ga": GROUP_SHEET,
         "register": REGISTER_SHEET}


def _owner(rid: str, model: Model, titles: dict[str, str]) -> str:
    head = rid.split("/", 1)[0]
    if head in FRAME:
        return FRAME[head]
    if head in titles:
        return titles[head]
    return head


def read_names(path: Path, names: list[str]) -> dict[str, float]:
    from hfgmodels.verify import _prop, libreoffice
    from live import _url
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(path), "_blank", 0, (_prop("Hidden", True), _prop("ReadOnly", True)))
        try:
            doc.calculateAll()
            out = {}
            for n in names:
                if doc.NamedRanges.hasByName(n):
                    out[n] = doc.NamedRanges.getByName(n).getReferredCells().getCellByPosition(0, 0).getValue()
            return out
        finally:
            doc.close(True)


def compare(path_a: Path, path_b: Path, lib: Library, values: bool = True) -> dict:
    meta_a, meta_b = read_metadata(path_a), read_metadata(path_b)
    ma, mb = Model.from_dict(lib, meta_a["model"]), Model.from_dict(lib, meta_b["model"])
    la, lb = assemble(ma), assemble(mb)
    titles = {i.uid: ma.title(i) for i in ma.instances} | {i.uid: mb.title(i) for i in mb.instances}
    ua, ub = [i.uid for i in ma.instances], [i.uid for i in mb.instances]

    out: dict = {"a": Path(path_a).name, "b": Path(path_b).name,
                 "modules_added": [titles[u] for u in ub if u not in ua],
                 "modules_removed": [titles[u] for u in ua if u not in ub]}

    ra = {r.id: r for _, rows in la.sheets for r in rows if r.kind != "blank"}
    rb = {r.id: r for _, rows in lb.sheets for r in rows if r.kind != "blank"}
    blocks: dict[str, dict] = {}
    for rid in list(ra) + [x for x in rb if x not in ra]:
        if rid in ra and rid in rb:
            if ra[rid].signature() == rb[rid].signature():
                continue
            what = "rewired"
        else:
            what = "removed" if rid in ra else "added"
        owner = _owner(rid, ma if rid in ra else mb, titles)
        b = blocks.setdefault(owner, {"added": 0, "removed": 0, "rewired": []})
        if what == "rewired":
            b["rewired"].append(rb[rid].label)
        else:
            b[what] += 1
    # Modules added or removed are listed once above; their own rows are not repeated here.
    out["rows"] = {k: v for k, v in blocks.items() if k not in out["modules_added"] + out["modules_removed"]}

    def setting_state(m: Model) -> dict[str, object]:
        lay = assemble(m)
        st = {}
        for _, rows in lay.sheets:
            for r in rows:
                if r.kind == "setting" and r.name and not r.id.startswith("register/"):
                    st[r.name] = ("group", r.link) if r.link else ("local", None)
        return st

    sa, sb = setting_state(ma), setting_state(mb)
    labels = {r.name: f"{_owner(r.id, mb, titles)}: {r.label}" for r in rb.values() if r.kind == "setting" and r.name} | \
             {r.name: f"{_owner(r.id, ma, titles)}: {r.label}" for r in ra.values() if r.kind == "setting" and r.name and r.name not in rb}
    out["bindings"] = [{"name": n, "input": labels[n], "was": sa[n][0], "now": sb[n][0]}
                       for n in sa if n in sb and sa[n][0] != sb[n][0]]

    ka = {h["label"]: h["name"] for h in la.headlines}
    kb = {h["label"]: h["name"] for h in lb.headlines}
    if values:
        names_a = [n for n in sa] + list(ka.values())
        names_b = [n for n in sb] + list(kb.values())
        va, vb = read_names(path_a, names_a), read_names(path_b, names_b)
        out["inputs"] = [{"name": n, "input": labels[n], "a": va[n], "b": vb[n]}
                         for n in sa if n in sb and n in va and n in vb and abs(va[n] - vb[n]) > 1e-12]
        out["key_outputs"] = [{"output": k, "a": va.get(ka[k]), "b": vb.get(kb[k])}
                              for k in kb if k in ka]
    recs_a, recs_b = ma.assurance.get("inputs", {}), mb.assurance.get("inputs", {})
    out["records"] = [{"name": n, "input": labels.get(n, n), "a": recs_a.get(n, {}), "b": recs_b.get(n, {})}
                      for n in sorted(set(recs_a) | set(recs_b)) if recs_a.get(n, {}) != recs_b.get(n, {})]
    set_a, set_b = ma.assurance.get("set") or {}, mb.assurance.get("set") or {}
    out["assumptions"] = None if set_a.get("version") == set_b.get("version") else {
        "a": set_a.get("version"), "b": set_b.get("version"),
        "items": [{"item": i["label"], "a": {x["key"]: x["value"] for x in set_a.get("items", [])}.get(i["key"]), "b": i["value"]}
                  for i in set_b.get("items", [])
                  if {x["key"]: x["value"] for x in set_a.get("items", [])}.get(i["key"]) != i["value"]]}
    log_a, log_b = ma.assurance.get("log", []), mb.assurance.get("log", [])
    out["log"] = log_b[len(log_a):] if log_b[:len(log_a)] == log_a else None   # None: not descended from a
    return out


def summary(res: dict) -> list[str]:
    """The comparison in plain lines, as the pane lists it."""
    out = []
    if res["modules_added"]:
        out.append("Modules added: " + ", ".join(res["modules_added"]) + ".")
    if res["modules_removed"]:
        out.append("Modules removed: " + ", ".join(res["modules_removed"]) + ".")
    for owner, b in res["rows"].items():
        bits = []
        if b["added"]:
            bits.append(f"{b['added']} row{'s' if b['added'] > 1 else ''} added")
        if b["removed"]:
            bits.append(f"{b['removed']} row{'s' if b['removed'] > 1 else ''} removed")
        if b["rewired"]:
            bits.append(f"{len(b['rewired'])} rewired ({', '.join(dict.fromkeys(b['rewired']))})")
        out.append(f"{owner}: " + ", ".join(bits) + ".")
    for i in res.get("inputs", []):
        out.append(f"Input changed: {i['input']} from {i['a']:,.4g} to {i['b']:,.4g}.")
    for bnd in res["bindings"]:
        out.append(f"{bnd['input']}: {'now drawn from the group assumptions' if bnd['now'] == 'group' else 'now a local value'}.")
    if res["assumptions"]:
        out.append(f"Group assumptions: version {res['assumptions']['a']} to {res['assumptions']['b']} (" +
                   ", ".join(f"{i['item']} {i['a']} to {i['b']}" for i in res["assumptions"]["items"]) + ").")
    for r in res["records"]:
        out.append(f"Record edited: {r['input']}.")
    for k in res.get("key_outputs", []):
        if k["a"] is not None and abs(k["a"] - k["b"]) > 0.005:
            out.append(f"{k['output']}: {k['a']:,.2f} to {k['b']:,.2f} ({k['b'] - k['a']:+,.2f}).")
    if res["log"]:
        out.append(f"Change log: {len(res['log'])} command{'s' if len(res['log']) > 1 else ''} in between (" +
                   "; ".join(e["command"] for e in res["log"]) + ").")
    return out or ["No differences."]
