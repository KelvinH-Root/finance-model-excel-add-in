"""The Impacts command's live mode: change one input in the open model, recalculate, read the
statements, and put the input back, leaving the model exactly as it was.

In the add-in this runs through Office.js in one Excel.run (write the input, calculate, read the
statement rows, restore, calculate); here it runs through LibreOffice's UNO bridge on a model
built by the assembly engine, so the tests can compare every movement with the assembly
reference and check that nothing in the workbook changed.

`chain` reads the link records the engine keeps in the model's metadata and lists the path a
change takes from the module that owns the input to the statements, block by block, so the
pane can say why each line moved (Revenue line 1 to Debtors to receipts and closing debtors).
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "assembly"))
from assemble import FIRST_PERIOD_COL  # noqa: E402

STATEMENTS = "demo.statements#1"
SECTIONS = ("Income statement", "Cash flow", "Balance sheet")


@dataclass
class Line:
    id: str
    label: str
    section: str
    total: bool


def statement_lines(layout) -> list[Line]:
    """The statement rows of the model, by section, in sheet order (checks left out)."""
    out, section = [], None
    for sheet, rows in layout.sheets:
        for r in rows:
            if not r.id.startswith(STATEMENTS + "/"):
                continue
            if r.kind == "section":
                section = r.label
            elif r.kind == "series" and r.style != "check" and section in SECTIONS and not r.id.endswith("bs_check"):
                out.append(Line(r.id, r.label, section, r.style == "total"))
    return out


def block_of(layout, row_id: str) -> str:
    """The block (module instance, or a mirror block inside one) a row belongs to."""
    best = ""
    for b in layout.blocks:
        if (row_id == b or row_id.startswith(b + "/")) and len(b) > len(best):
            best = b
    return best


def chain(layout, uid: str) -> list[tuple[str, str, str]]:
    """(from block title, link, to block title) for every link a change to instance uid travels along."""
    reached = {b for b in layout.blocks if b == uid or b.startswith(uid + "/")}
    edges, frontier = [], set(reached)
    while frontier:
        nxt = set()
        for rec in layout.records:
            if rec["link"].startswith("check."):
                continue
            src, dst = block_of(layout, rec["from"]), block_of(layout, rec["to"])
            if src in frontier:
                e = (layout.blocks[src], rec["link"], layout.blocks[dst])
                if e not in edges:
                    edges.append(e)
                if dst not in reached:
                    reached.add(dst)
                    nxt.add(dst)
        frontier = nxt
    return edges


def setting_of(layout, name: str) -> tuple[str, str]:
    """(instance uid, setting key) behind a named input such as Rev1_Base."""
    rid = layout.names[name]
    uid, key = rid.rsplit("/", 1)
    return uid, key.removeprefix("set.")


def _fingerprint(doc) -> dict:
    out = {}
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        rng = sh.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
        out[sh.Name] = (rng.getFormulaArray(), rng.getDataArray())
    return out


def _read(doc, layout, lines: list[Line], periods: int) -> dict[str, list[float]]:
    pos = layout.positions()
    out = {}
    for ln in lines:
        sheet, row = pos[ln.id]
        rng = doc.Sheets.getByName(sheet).getCellRangeByPosition(FIRST_PERIOD_COL - 1, row - 1,
                                                                   FIRST_PERIOD_COL - 2 + periods, row - 1)
        out[ln.id] = [float(v) for v in rng.getDataArray()[0]]
    return out


def run(doc, layout, name: str, value: float) -> dict:
    """Change the named input to value, read every statement line, then restore it.

    Returns {"before", "after", "delta"} by line id and period, the lines, the link chain, and
    whether the workbook came back exactly as it was ("restored")."""
    periods = layout.periods
    lines = statement_lines(layout)
    cell = doc.NamedRanges.getByName(name).getReferredCells().getCellByPosition(0, 0)
    kept = cell.getFormula()
    print_before = _fingerprint(doc)
    before = _read(doc, layout, lines, periods)
    try:
        cell.setValue(value)
        doc.calculateAll()
        after = _read(doc, layout, lines, periods)
    finally:
        cell.setFormula(kept)
        doc.calculateAll()
    restored = _fingerprint(doc) == print_before
    delta = {k: [a - b for a, b in zip(after[k], before[k])] for k in before}
    uid, _ = setting_of(layout, name)
    return {"name": name, "value": value, "was": float(kept) if kept else 0.0, "lines": lines, "before": before,
            "after": after, "delta": delta, "chain": chain(layout, uid), "restored": restored}
