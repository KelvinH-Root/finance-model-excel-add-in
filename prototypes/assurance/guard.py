"""Key outputs, the change log and the before-and-after check on every structural command.

Every structural command runs the same way in the add-in:

1. reconcile: read the register back into the metadata (register.read_back);
2. plan: work out the change plan and the key outputs the change can reach;
3. read the key outputs (the KO_ names on the contents), apply the plan, recalculate, read them again;
4. compare: a key output that moved where the command could not reach it, or moved at all for a
   command that should leave the numbers alone (rename, insert a summary, extend the timeline),
   is unexplained, and the result card says so before anything else;
5. log: one line in the change log in the model's metadata (who, when, the command, the preview,
   the key outputs before and after, what moved and what was unexplained).

What a change can reach is worked out from the layout alone: every row's formula names the rows
and named cells it reads, so the engine follows those references forward from the rows the plan
writes. Nothing here knows about revenue or debt; it works for any model the engine builds.

Here the live writer is LibreOffice through UNO (prototypes/assembly/live.py); in the add-in it is
Office.js, and step 3 is one Excel.run.
"""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "assembly"))
sys.path.insert(0, str(HERE))
from assemble import (FIRST_PERIOD_COL, Layout, Library, Model, assemble, col_letter,  # noqa: E402
                      open_model, plan_change, write_metadata)
from register import read_back  # noqa: E402

TOLERANCE = 0.005   # half a cent: anything smaller is rounding
_REF = re.compile(r"«([RPAC]\d*)\|([^»]+)»(?::«[RPAC]\d*\|([^»]+)»)?")
_WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def dependents(layout: Layout) -> dict[str, set[str]]:
    """Row id -> the rows that read it (by marker, range or defined name)."""
    order: dict[str, list[str]] = {s: [r.id for r in rows] for s, rows in layout.sheets}
    where = {rid: (s, k) for s, ids in order.items() for k, rid in enumerate(ids)}
    out: dict[str, set[str]] = {}
    for s, rows in layout.sheets:
        for r in rows:
            texts = [t for t in (r.first, r.formula, r.link) if t]
            texts += [v for v in r.cells.values() if isinstance(v, str) and v.startswith("=")]
            for t in texts:
                for m in _REF.finditer(t):
                    a, b = m.group(2), m.group(3)
                    if b:  # a range between two rows on one sheet: every row in it
                        sa, ka = where[a]
                        _, kb = where[b]
                        srcs = order[sa][min(ka, kb):max(ka, kb) + 1]
                    else:
                        srcs = [a]
                    for src in srcs:
                        out.setdefault(src, set()).add(r.id)
                for w in _WORD.findall(_REF.sub("", t)):
                    if w in layout.names:
                        out.setdefault(layout.names[w], set()).add(r.id)
    return out


def written_rows(plan, layout: Layout) -> set[str]:
    """The rows a plan writes (new or rewired), as row ids in the new layout."""
    at = {(s, r): rid for rid, (s, r) in layout.positions().items() if not rid.startswith("@")}
    return {at[(op["sheet"], op["row"])] for op in plan.ops
            if op["op"] == "write" and op.get("why") in ("new", "rewire") and (op["sheet"], op["row"]) in at}


def reachable(layout: Layout, start: set[str]) -> set[str]:
    deps = dependents(layout)
    seen, todo = set(start), list(start)
    while todo:
        for nxt in deps.get(todo.pop(), ()):
            if nxt not in seen:
                seen.add(nxt)
                todo.append(nxt)
    return seen


def expected_outputs(layout: Layout, plan, effect: str) -> list[str]:
    """Labels of the key outputs this change may move. effect 'none' means it must move none."""
    if effect == "none":
        return []
    reach = reachable(layout, written_rows(plan, layout))
    return [h["label"] for h in layout.headlines if f"contents/ko/{h['name']}" in reach]


def read_key_outputs(path: Path, layout: Layout) -> dict[str, float]:
    """Recalculate and read the KO_ cells, as the add-in reads them before and after a command."""
    from live import _url
    from hfgmodels.verify import _prop, libreoffice
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(path), "_blank", 0, (_prop("Hidden", True), _prop("ReadOnly", True)))
        try:
            doc.calculateAll()
            out = {}
            for h in layout.headlines:
                cell = doc.NamedRanges.getByName(h["name"]).getReferredCells().getCellByPosition(0, 0)
                out[h["label"]] = cell.getValue()
            return out
        finally:
            doc.close(True)


def run(src: Path, dst: Path, lib: Library, command: str, change: Callable[[Model], None], effect: str = "reach",
        who: str = "Kelvin", sabotage: Callable[[Layout], list[dict]] | None = None, when: str | None = None) -> dict:
    """Run one structural command on the file at src, write the result to dst, and log it.

    sabotage adds operations to the plan, standing in for a fault in the engine or a formula
    someone edited by hand, to show the check catches what it should."""
    from live import apply_plan

    model, _ = open_model(src, lib)
    old = assemble(model)
    picked_up = read_back(src, model, old)
    old = assemble(model)
    before = read_key_outputs(src, old)
    change(model)
    new = assemble(model)
    plan = plan_change(old, new, dialect="uno")
    if sabotage:
        plan.ops.extend(sabotage(new))
    expected = expected_outputs(new, plan, effect)
    apply_plan(src, plan, dst, model, new)
    after = read_key_outputs(dst, new)
    moved = [k for k in after if abs(after[k] - before.get(k, 0.0)) > TOLERANCE]
    unexplained = [k for k in moved if k not in expected]
    entry = {"when": when or datetime.now().strftime("%d %B %Y %H:%M").lstrip("0"), "who": who, "command": command,
             "effect": effect, "preview": plan.preview, "picked_up": sorted(picked_up),
             "outputs": {k: [round(before.get(k, 0.0), 2), round(after[k], 2)] for k in after},
             "expected": expected, "moved": moved, "unexplained": unexplained}
    model.assurance.setdefault("log", []).append(entry)
    write_metadata(dst, model, new)
    return entry


def result_card(entry: dict) -> list[str]:
    """The lines the result card shows, unexplained movements first."""
    out = []
    if entry["unexplained"]:
        names = entry["unexplained"]
        listed = names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]
        out.append(f"Check this change: {listed} moved, " +
                   ("and this command should not move any key output." if entry["effect"] == "none"
                    else f"but nothing this command changed feeds {'it' if len(names) == 1 else 'them'}."))
    for k in entry["moved"]:
        a, b = entry["outputs"][k]
        out.append(f"{k}: {a:,.2f} to {b:,.2f} ({b - a:+,.2f})")
    if not entry["moved"]:
        out.append("No key output moved.")
    return out


def facility_repayments_stop(layout: Layout, uid: str = "demo.facility#1") -> list[dict]:
    """A fault for the tests: the facility's repayments written as nil from month 2 (uno dialect)."""
    sheet, row = layout.positions()[f"{uid}/repay"]
    return [{"op": "write", "sheet": sheet, "row": row, "why": "fault", "kind": "series", "style": "", "unit": "$",
             "cells": {c: "=0" for c in range(FIRST_PERIOD_COL + 1, FIRST_PERIOD_COL + layout.periods)}}]


__all__ = ["dependents", "expected_outputs", "read_key_outputs", "run", "result_card", "facility_repayments_stop",
           "col_letter"]
