"""Live writer stand-in: apply a change plan to an existing workbook.

The add-in will apply plans through Office.js. This proof applies the same
operations through LibreOffice, because LibreOffice, like Excel, shifts every
reference, name and range when rows are inserted or deleted. Each operation maps
to one Office.js call:

    insert_rows   sheet.getRange("7:9").insert(Excel.InsertShiftDirection.down)
    delete_rows   sheet.getRange("7:9").delete(Excel.DeleteShiftDirection.up)
    add_sheet     workbook.worksheets.add(name); worksheet.position = index
    delete_sheet  worksheet.delete()
    write         range.formulas = [[...]] / range.values = [[...]]
    add_name      workbook.names.add(name, range)
    delete_name   workbook.names.getItem(name).delete()
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from hfgmodels.verify import _prop, libreoffice  # noqa: E402

from assemble import (TOTAL_COL, Layout, Model, Plan, col_letter, frame_cells,  # noqa: E402
                      write_metadata)


def _url(p: Path) -> str:
    return "file://" + str(Path(p).resolve())


def apply_plan(src: Path, plan: Plan, dst: Path, model: Model, layout: Layout) -> Path:
    """Open src, apply the plan in order, save as dst, then update the metadata part."""
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(src), "_blank", 0, (_prop("Hidden", True),))
        try:
            sheets = doc.Sheets
            for op in plan.ops:
                kind = op["op"]
                if kind == "delete_name":
                    if doc.NamedRanges.hasByName(op["name"]):
                        doc.NamedRanges.removeByName(op["name"])
                elif kind == "delete_sheet":
                    sheets.removeByName(op["sheet"])
                elif kind == "add_sheet":
                    sheets.insertNewByName(op["sheet"], op["index"])
                    sh = sheets.getByName(op["sheet"])
                    for (r, c), v in frame_cells(op["sheet"], op["periods"]).items():
                        _set(sh.getCellByPosition(c - 1, r - 1), v)
                elif kind == "insert_rows":
                    sheets.getByName(op["sheet"]).Rows.insertByIndex(op["row"] - 1, op["count"])
                elif kind == "delete_rows":
                    sheets.getByName(op["sheet"]).Rows.removeByIndex(op["row"] - 1, op["count"])
                elif kind == "write":
                    sh = sheets.getByName(op["sheet"])
                    for c, v in op["cells"].items():
                        _set(sh.getCellByPosition(int(c) - 1, op["row"] - 1), v)
                elif kind == "add_name":
                    sheet = op["sheet"]
                    quoted = sheet if sheet.replace("_", "").isalnum() else f"'{sheet}'"
                    content = f"${quoted}.${col_letter(TOTAL_COL)}${op['row']}"
                    if doc.NamedRanges.hasByName(op["name"]):
                        doc.NamedRanges.removeByName(op["name"])
                    doc.NamedRanges.addNewByName(op["name"], content,
                                                 sheets.getByIndex(0).getCellByPosition(0, 0).CellAddress, 0)
                else:
                    raise ValueError(f"unknown operation {kind}")
            doc.calculateAll()
            doc.storeToURL(_url(dst), (_prop("FilterName", "Calc MS Excel 2007 XML"),))
        finally:
            doc.close(True)
    write_metadata(dst, model, layout)
    return Path(dst)


def _set(cell, v) -> None:
    if v is None:
        cell.setString("")
    elif isinstance(v, str) and v.startswith("="):
        cell.setFormula(v)
    elif isinstance(v, (int, float)):
        cell.setValue(v)
    else:
        cell.setString(str(v))


def snapshot(path: Path) -> dict:
    """Formulas, values and names of every sheet after a full recalculation, read the same way for any file."""
    out = {"sheets": {}, "names": {}, "errors": {}}
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(path), "_blank", 0, (_prop("Hidden", True), _prop("ReadOnly", True)))
        try:
            doc.calculateAll()
            for i in range(doc.Sheets.Count):
                sh = doc.Sheets.getByIndex(i)
                cur = sh.createCursor()
                cur.gotoEndOfUsedArea(False)
                rng = sh.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
                formulas = [list(r) for r in rng.getFormulaArray()]
                values = [list(r) for r in rng.getDataArray()]
                out["sheets"][sh.Name] = {"formulas": formulas, "values": values}
                err = sh.queryFormulaCells(4)
                out["errors"][sh.Name] = [a for a in err.getRangeAddressesAsString().split(";") if a]
            for n in doc.NamedRanges.ElementNames:
                out["names"][n] = doc.NamedRanges.getByName(n).Content
        finally:
            doc.close(True)
    return out
