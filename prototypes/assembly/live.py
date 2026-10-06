"""Live writer stand-in: apply a change plan to an existing workbook.

The add-in will apply plans through Office.js. This proof applies the same
operations through LibreOffice, because LibreOffice, like Excel, shifts every
reference, name and range when rows are inserted or deleted. Each operation maps
to one Office.js call:

    insert_rows   sheet.getRange("7:9").insert(Excel.InsertShiftDirection.down)
    delete_rows   sheet.getRange("7:9").delete(Excel.DeleteShiftDirection.up)
    add_sheet     workbook.worksheets.add(name); worksheet.position = index; header cells and links written
                  (section covers and the contents are ordinary sheets the plan writes)
    delete_sheet  worksheet.delete()
    write         range.formulas = [[...]] / range.values = [[...]]
    add_name      workbook.names.add(name, range)
    delete_name   workbook.names.getItem(name).delete()
    add_chart     sheet.charts.add(type, range); chart.series.add(name).setValues(range)
    set_chart     chart.series: delete, then add(name).setValues(range) and setXAxisValues(range)
    delete_chart  chart.delete()

Charts are found by title, which a module keeps unique on its sheet.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from hfgmodels.verify import _prop, libreoffice, uno_context  # noqa: E402

from assemble import Layout, Model, Plan, col_letter, write_metadata  # noqa: E402


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
                    for r, c, v in op["frame"]:
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
                    content = f"${quoted}.${col_letter(op['col'])}${op['row']}"
                    if doc.NamedRanges.hasByName(op["name"]):
                        doc.NamedRanges.removeByName(op["name"])
                    doc.NamedRanges.addNewByName(op["name"], content,
                                                 sheets.getByIndex(0).getCellByPosition(0, 0).CellAddress, 0)
                elif kind == "delete_chart":
                    sh = sheets.getByName(op["sheet"])
                    sh.Charts.removeByName(_chart_by_title(sh, op["title"]))
                elif kind == "add_chart":
                    _add_chart(sheets.getByName(op["sheet"]), op)
                elif kind == "set_chart":
                    sh = sheets.getByName(op["sheet"])
                    _set_chart(sh.Charts.getByName(_chart_by_title(sh, op["title"])).EmbeddedObject, op)
                else:
                    raise ValueError(f"unknown operation {kind}")
            doc.calculateAll()
            doc.storeToURL(_url(dst), (_prop("FilterName", "Calc MS Excel 2007 XML"),))
        finally:
            doc.close(True)
    write_metadata(dst, model, layout)
    return Path(dst)


CHART_SIZE_MM100 = (16000, 7500)


def _chart_title(emb) -> str:
    t = emb.getTitleObject()
    return "".join(x.getString() for x in t.getText()) if t else ""


def _chart_by_title(sheet, title: str) -> str:
    for n in sheet.Charts.ElementNames:
        if _chart_title(sheet.Charts.getByName(n).EmbeddedObject) == title:
            return n
    raise ValueError(f"{sheet.Name}: no chart titled {title!r}")


def _add_chart(sheet, op: dict) -> None:
    from com.sun.star.awt import Rectangle
    cell = sheet.getCellByPosition(op["anchor"]["col"] - 1, op["anchor"]["row"] - 1)
    name = f"hfg_chart_{len(sheet.Charts.ElementNames) + 1}"
    rect = Rectangle(cell.Position.X, cell.Position.Y, *CHART_SIZE_MM100)
    sheet.Charts.addNewByName(name, rect, (cell.getRangeAddress(),), False, False)
    emb = sheet.Charts.getByName(name).EmbeddedObject
    _set_chart(emb, op)
    emb.HasMainTitle = True
    emb.Title.String = op["title"]
    emb.HasLegend = True


def _set_chart(emb, op: dict) -> None:
    """Rebuild every series from the plan: stacked columns, then lines; the x axis from the categories."""
    import uno
    ctx = uno_context()
    smgr = ctx.ServiceManager
    provider = emb.getDataProvider()

    def seq(rng, role):
        s = provider.createDataSequenceByRangeRepresentation(rng)
        s.Role = role
        return s

    def labelled(values, label=None, role="values-y"):
        lds = smgr.createInstanceWithContext("com.sun.star.chart2.data.LabeledDataSequence", ctx)
        lds.setValues(seq(values, role))
        if label:
            lds.setLabel(seq(label, "label"))
        return lds

    def series(spec, stacked):
        s = smgr.createInstanceWithContext("com.sun.star.chart2.DataSeries", ctx)
        s.setData((labelled(spec["values"], spec["label"]),))
        if stacked:
            s.StackingDirection = uno.Enum("com.sun.star.chart2.StackingDirection", "Y_STACKING")
        return s

    cs = emb.getFirstDiagram().getCoordinateSystems()[0]
    types = []
    cols = [series(x, True) for x in op["series"] if x["kind"] == "column"]
    lines = [series(x, False) for x in op["series"] if x["kind"] == "line"]
    if cols:
        t = smgr.createInstanceWithContext("com.sun.star.chart2.ColumnChartType", ctx)
        t.setDataSeries(tuple(cols))
        types.append(t)
    if lines:
        t = smgr.createInstanceWithContext("com.sun.star.chart2.LineChartType", ctx)
        t.setDataSeries(tuple(lines))
        types.append(t)
    cs.setChartTypes(tuple(types))
    axis = cs.getAxisByDimension(0, 0)
    scale = axis.getScaleData()
    scale.Categories = labelled(op["categories"], role="categories")
    axis.setScaleData(scale)


def read_charts(doc) -> dict:
    """Every chart by sheet and title: its series ranges by type, the categories and the anchor cell."""
    out = {}
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        anchors = {}
        dp = sh.DrawPage
        for k in range(dp.Count):
            shape = dp.getByIndex(k)
            a = getattr(shape, "Anchor", None)
            if a is not None and hasattr(a, "CellAddress"):
                anchors[getattr(shape, "PersistName", shape.Name)] = (a.CellAddress.Row + 1, a.CellAddress.Column + 1)
        for n in sh.Charts.ElementNames:
            emb = sh.Charts.getByName(n).EmbeddedObject
            entry = {"columns": [], "lines": [], "categories": None, "anchor": anchors.get(n)}
            for cs in emb.getFirstDiagram().getCoordinateSystems():
                sd = cs.getAxisByDimension(0, 0).getScaleData()
                if sd.Categories is not None:
                    entry["categories"] = sd.Categories.getValues().SourceRangeRepresentation
                for ct in cs.getChartTypes():
                    kind = "lines" if ct.getChartType().endswith("LineChartType") else "columns"
                    for ser in ct.getDataSeries():
                        for lds in ser.getDataSequences():
                            lab = lds.getLabel().SourceRangeRepresentation if lds.getLabel() else None
                            stacked = kind == "columns" and ser.StackingDirection.value == "Y_STACKING"
                            entry[kind].append((lab, lds.getValues().SourceRangeRepresentation) + ((stacked,) if kind == "columns" else ()))
            out.setdefault(sh.Name, {})[_chart_title(emb)] = entry
    return out


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
    out = {"sheets": {}, "names": {}, "errors": {}, "charts": {}}
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
            out["charts"] = read_charts(doc)
        finally:
            doc.close(True)
    return out
