"""Release profiles: a clean copy of a finalised model for someone outside the finance team.

Release copy becomes a menu of profiles. Each profile names the results the recipient needs and
what they should see; the engine works out the rest from the layout:

- only the sheets needed to build those results (the precedents of every row on the result sheets,
  followed through formulas and names), plus the Input register and Checks when the profile asks;
- a contents for the sheets kept, with the model, the profile, the set and the key outputs;
- internal columns removed from the register (owner, evidence links, reasons), and the register
  kept as values, since it documents rather than calculates;
- formulas kept (auditor) or every value fixed (lender, board);
- no navigation links, no subtitles, no add-in metadata.

In the add-in this runs on a copy saved by Office.js (or rebuilt by the package writer); here
LibreOffice stands in, as in the other proofs.
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "assembly"))
sys.path.insert(0, str(HERE))
from assemble import (CONTENTS, GROUP_SHEET, RC, REGISTER_SHEET, Layout, Library, assemble,  # noqa: E402
                      open_model)
from guard import dependents  # noqa: E402

PROFILES = {
    "auditor": {"title": "Auditor copy", "results": ["Statements"], "values": False, "register": True, "checks": False,
                "drop": ["owner", "evidence", "reason"]},
    "lender": {"title": "Lender copy", "results": ["Statements", "Funding"], "values": True, "register": False,
               "checks": True, "drop": []},
    "board": {"title": "Board copy", "results": ["Statements", "Dashboard"], "values": True, "register": False,
              "checks": False, "drop": []},
}


def precedents(layout: Layout) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    for src, dsts in dependents(layout).items():
        for d in dsts:
            out.setdefault(d, set()).add(src)
    return out


def needed_sheets(layout: Layout, sheets: list[str]) -> list[str]:
    """The sheets the rows on the given sheets need, in workbook order."""
    rows = layout.sheet_rows()
    where = {r.id: s for s, rs in layout.sheets for r in rs}
    prec = precedents(layout)
    todo = [r.id for s in sheets if s in rows for r in rows[s]]
    seen = set(todo)
    while todo:
        for p in prec.get(todo.pop(), ()):
            if p not in seen:
                seen.add(p)
                todo.append(p)
    keep = {where[r] for r in seen if r in where}
    return [s for s, _ in layout.sheets if s in keep]


def plan_release(layout: Layout, profile: str) -> list[str]:
    """Sheets in the release copy: the contents, then the result sheets and (when formulas are kept)
    every sheet they need, plus Checks and the register when the profile asks. The register goes as
    values, so its links to every input do not pull their sheets in."""
    p = PROFILES[profile]
    present = layout.sheet_rows()
    roots = [s for s in list(p["results"]) + (["Checks"] if p["checks"] else []) if s in present]
    keep = set(roots) if p["values"] else set(needed_sheets(layout, roots))   # values travel without their precedents
    if p["register"] and REGISTER_SHEET in present:
        keep.add(REGISTER_SHEET)
    return [CONTENTS] + [s for s, _ in layout.sheets if s in keep and s != CONTENTS]


def release(src: Path, dst: Path, lib: Library, profile: str, who: str = "Kelvin", when: str = "") -> dict:
    from hfgmodels.verify import _prop, libreoffice
    from live import _url

    model, _ = open_model(src, lib)
    layout = assemble(model)
    p = PROFILES[profile]
    keep = plan_release(layout, profile)
    gset = model.assurance.get("set") or {}
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(src), "_blank", 0, (_prop("Hidden", True),))
        try:
            doc.calculateAll()
            sheets = doc.Sheets
            # Key outputs as values, read before the contents is rewritten.
            kos = [(h["label"], doc.NamedRanges.getByName(h["name"]).getReferredCells().getCellByPosition(0, 0).getValue())
                   for h in layout.headlines if doc.NamedRanges.hasByName(h["name"])]
            for i in range(sheets.Count):
                sh = sheets.getByIndex(i)
                if sh.Name in keep and (p["values"] or sh.Name == REGISTER_SHEET):
                    cur = sh.createCursor()
                    cur.gotoEndOfUsedArea(False)
                    rng = sh.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
                    rng.setDataArray(rng.getDataArray())
            for name in [sheets.getByIndex(i).Name for i in range(sheets.Count)]:
                if name not in keep:
                    sheets.removeByName(name)
            for n in list(doc.NamedRanges.ElementNames):
                nr = doc.NamedRanges.getByName(n)
                gone = "#REF" in nr.Content.upper()
                if gone or n.startswith(("HL_", "KO_")) or (n.startswith("Chk_") and "Checks" not in keep):
                    doc.NamedRanges.removeByName(n)
            for name in keep:
                sh = sheets.getByName(name)
                for a1 in ("A1", "A2", "B2"):   # navigation links and subtitles
                    sh.getCellRangeByName(a1).setString("")
            if REGISTER_SHEET in keep:
                reg = sheets.getByName(REGISTER_SHEET)
                for col in sorted((RC[k] for k in p["drop"]), reverse=True):
                    reg.Columns.removeByIndex(col - 1, 1)
            toc = sheets.getByName(CONTENTS)
            toc.getCellRangeByPosition(0, 0, 40, 200).clearContents(1 | 2 | 4 | 16 | 32 | 64)   # contents, formats, styles
            lines = [("Contents", None), ("", None), (f"{p['title']}", None),
                     (f"Released {when} by {who}" if when else f"Released by {who}", None),
                     (f"{gset.get('title', '')} version {gset.get('version')}" if gset else "", None), ("", None)]
            for s in keep[1:]:
                lines.append((s, f'=HYPERLINK("#\'{s}\'!A1";"{s}")'))   # Excel's form: the copy is for Excel
            lines.append(("", None))
            lines.append(("Key outputs", None))
            for label, v in kos:
                lines.append((label, v))
            body = toc.getCellRangeByPosition(0, 0, 8, len(lines))
            body.CharFontName, body.CharHeight, body.CharColor = "Segoe UI", 9, 0x404040
            money = doc.NumberFormats.queryKey('#,##0.00;(#,##0.00);"-"', doc.CharLocale, False)
            if money == -1:
                money = doc.NumberFormats.addNew('#,##0.00;(#,##0.00);"-"', doc.CharLocale)
            for r, (text, extra) in enumerate(lines):
                cell = toc.getCellByPosition(1, r)
                if isinstance(extra, str):
                    cell.setFormula(extra)
                    cell.CharColor, cell.CharUnderline = 0x0563C1, 1
                else:
                    cell.setString(text)
                    if text in ("Contents", "Key outputs"):
                        cell.CharWeight, cell.CharHeight = 150, 10
                    if isinstance(extra, float):
                        v = toc.getCellByPosition(8, r)
                        v.setValue(extra)
                        v.NumberFormat = money
            doc.calculateAll()
            doc.storeToURL(_url(dst), (_prop("FilterName", "Calc MS Excel 2007 XML"),))
        finally:
            doc.close(True)
    _strip_metadata(dst)
    return {"profile": profile, "sheets": keep, "key_outputs": dict(kos)}


def _strip_metadata(path: Path) -> None:
    """A release copy carries no add-in metadata (LibreOffice drops it anyway; Excel would not)."""
    import re
    import shutil
    import tempfile
    import zipfile
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    if not any(n.startswith("customXml/") for n in parts):
        return
    parts = {n: d for n, d in parts.items() if not n.startswith("customXml/")}
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = re.sub(r'<Relationship [^>]*customXml[^>]*/>', "", rels).encode()
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    import os
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    shutil.move(tmp, path)


__all__ = ["PROFILES", "needed_sheets", "plan_release", "release", "GROUP_SHEET"]
