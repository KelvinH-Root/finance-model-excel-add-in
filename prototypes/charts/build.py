"""Phase 0 proof: native, live report charts and controls written by the package writer.

Builds a workbook with fictional numbers that shows what the Build tab's Charts and
Controls commands will produce, as ordinary Excel objects driven by formulas:

- a Z chart (monthly actual and forecast bars, budget bars behind, cumulative actual then
  forecast, cumulative budget, moving annual total), with the line chosen by a classic
  form control drop-down (combo box) linked to a named cell;
- two waterfalls (P&L walk, budget to actual) built as stacked columns, so totals,
  increases and decreases are coloured by formula and stay right when numbers change;
- IBCS column and variance charts (prior year grey, plan outlined, actual solid,
  forecast hatched; good variances green, bad red), with the comparison chosen by an
  in-cell drop-down;
- a contents sheet, a Lookups sheet with the lists, and a Checks sheet.

    python prototypes/charts/build.py [out.xlsx]

Nothing here needs the add-in to calculate; the charts update when the numbers or the
selections change. Form controls cannot be created through Office.js, so the package
writer writes them as file parts (a control properties part, a VML shape and the sheet's
controls element), the same parts Excel writes.
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import LegendEntry
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.fill import ColorChoice, PatternFillProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

INK, GREY, LIGHT, GOOD, BAD = "262626", "A6A6A6", "D9D9D9", "8CB400", "FF0000"
TEXT = "404040"
F_BODY = Font(name="Segoe UI", size=9, color=TEXT)
F_BOLD = Font(name="Segoe UI", size=9, color=TEXT, bold=True)
F_HEAD = Font(name="Segoe UI", size=10, color=TEXT, bold=True)
F_CHECK = Font(name="Segoe UI", size=9, color="9C0006")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
TOP = Border(top=Side(style="thin", color=TEXT))
NUM = '#,##0;(#,##0);"-"'

MONTHS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
LAST_ACTUAL = 6  # September
LINES = {
    # line: (prior year monthly, actual or forecast monthly, budget monthly)
    "Revenue": ([5, 6, 7, 7, 7, 7, 7, 4, 6, 7, 8, 7],
                [8, 8, 13, 7, 6, 2, 6, 9, 7, 7, 4, 5],
                [6, 6, 5, 10, 8, 3, 8, 7, 6, 7, 4, 6]),
    "Cost of sales": ([3, 4, 4, 4, 4, 4, 4, 3, 3, 4, 4, 4],
                      [5, 5, 7, 4, 4, 2, 4, 5, 4, 4, 3, 3],
                      [4, 4, 3, 6, 5, 2, 5, 4, 4, 4, 3, 4]),
}


def style_sheet(ws, title: str, purpose: str):
    ws.sheet_view.showGridLines = False
    ws["B1"], ws["B2"] = title, purpose
    ws["B1"].font, ws["B2"].font = F_HEAD, F_BODY
    ws.column_dimensions["A"].width = 2.5


def put(ws, ref, value, font=None, fmt=None, fill=None, border=None):
    c = ws[ref]
    c.value = value
    c.font = font or F_BODY
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


def name(wb, nm, ws, ref):
    col = "".join(ch for ch in ref if ch.isalpha())
    row = "".join(ch for ch in ref if ch.isdigit())
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{ws.title}'!${col}${row}")


def no_line():
    return LineProperties(noFill=True)


def build(path: Path) -> Path:
    wb = Workbook()
    contents = wb.active
    contents.title = "Contents"
    style_sheet(contents, "Contents", "Charts and controls the Build tab will insert (fictional numbers)")

    # Lookups --------------------------------------------------------------------
    lk = wb.create_sheet("Lookups")
    style_sheet(lk, "Lookups", "Lists behind the drop-downs; a list made from a category set grows with it")
    put(lk, "B4", "Chart lines", F_BOLD)
    names = list(LINES) + ["Gross profit"]
    for i, n in enumerate(names):
        put(lk, f"B{5 + i}", n)
    wb.defined_names["LU_Chart_Lines"] = DefinedName("LU_Chart_Lines", attr_text=f"Lookups!$B$5:$B${4 + len(names)}")
    put(lk, "D4", "Compare with", F_BOLD)
    put(lk, "D5", "Plan")
    put(lk, "D6", "Prior year")
    wb.defined_names["LU_Compare"] = DefinedName("LU_Compare", attr_text="Lookups!$D$5:$D$6")
    lk.column_dimensions["B"].width = 18
    lk.column_dimensions["D"].width = 14

    # Data ------------------------------------------------------------------------
    data = wb.create_sheet("Data")
    style_sheet(data, "Data", "Monthly prior year, actual then forecast, and budget for each line")
    put(data, "B4", "Line shown in the Z chart", F_BOLD)
    put(data, "H4", 1, fmt="0", fill=FILL_IN)
    name(wb, "DD_Chart_Line", data, "H4")
    put(data, "I4", "=INDEX(LU_Chart_Lines,DD_Chart_Line)", F_BOLD)
    put(data, "B5", "Last actual month", F_BOLD)
    put(data, "H5", LAST_ACTUAL, fmt="0", fill=FILL_IN)
    name(wb, "Ts_Last_Actual", data, "H5")
    put(data, "I5", "=INDEX($J$8:$U$8,Ts_Last_Actual)", F_BODY)
    put(data, "B8", "Month", F_BOLD)
    for m, mon in enumerate(MONTHS):
        put(data, f"{L(10 + m)}8", mon, F_BOLD)
    put(data, "H8", "Total", F_BOLD)
    r = 10
    blocks = {}
    for line in names:
        put(data, f"B{r}", line, F_HEAD)
        rows = {}
        for k, label in (("py", "Prior year"), ("acfc", "Actual then forecast"), ("bud", "Budget")):
            rr = r + 1 + len(rows)
            rows[k] = rr
            put(data, f"C{rr}", label)
            for m in range(12):
                cell = f"{L(10 + m)}{rr}"
                if line == "Gross profit":
                    a, b = blocks["Revenue"][k], blocks["Cost of sales"][k]
                    put(data, cell, f"={L(10 + m)}{a}-{L(10 + m)}{b}", fmt=NUM)
                else:
                    put(data, cell, LINES[line][("py", "acfc", "bud").index(k)][m], fmt=NUM, fill=FILL_IN)
            put(data, f"H{rr}", f"=SUM(J{rr}:U{rr})", F_BOLD, NUM)
        blocks[line] = rows
        r += 5
    data.column_dimensions["B"].width = 2.5
    data.column_dimensions["C"].width = 22
    for m in range(12):
        data.column_dimensions[L(10 + m)].width = 7

    # Classic form control: a combo box linked to DD_Chart_Line, list LU_Chart_Lines.
    # Written into the file after openpyxl saves (see add_form_control).

    # Z chart sheet ----------------------------------------------------------------
    zs = wb.create_sheet("Z chart")
    style_sheet(zs, "Z chart", "Monthly bars, year to date and moving annual total, actual then forecast against budget")
    put(zs, "B4", "Line", F_BOLD)
    put(zs, "D4", "=INDEX(LU_Chart_Lines,DD_Chart_Line)", F_HEAD)
    heads = ["Month", "AC", "FC", "Budget", "YTD AC", "YTD FC", "YTD budget", "MAT", "Scale"]
    for i, h in enumerate(heads):
        put(zs, f"{L(2 + i)}6", h, F_BOLD)
    # Row offsets of the selected line's rows on Data: block start 10 + 5 * (n - 1)
    sel = "(DD_Chart_Line-1)*5"
    for m in range(12):
        rr = 7 + m
        col = L(10 + m)
        acfc = f"INDEX(Data!${col}$1:${col}$40,12+{sel})"
        bud = f"INDEX(Data!${col}$1:${col}$40,13+{sel})"
        put(zs, f"B{rr}", MONTHS[m])
        put(zs, f"C{rr}", f"=IF({m + 1}<=Ts_Last_Actual,{acfc},NA())", fmt=NUM)
        put(zs, f"D{rr}", f"=IF({m + 1}>Ts_Last_Actual,{acfc},NA())", fmt=NUM)
        put(zs, f"E{rr}", f"={bud}", fmt=NUM)
        put(zs, f"F{rr}", f"=IF({m + 1}<=Ts_Last_Actual,SUM(INDEX(Data!$J$1:$U$40,12+{sel},1):INDEX(Data!$J$1:$U$40,12+{sel},{m + 1})),NA())", fmt=NUM)
        put(zs, f"G{rr}", f"=IF({m + 1}>=Ts_Last_Actual,SUM(INDEX(Data!$J$1:$U$40,12+{sel},1):INDEX(Data!$J$1:$U$40,12+{sel},{m + 1})),NA())", fmt=NUM)
        put(zs, f"H{rr}", f"=SUM(INDEX(Data!$J$1:$U$40,13+{sel},1):INDEX(Data!$J$1:$U$40,13+{sel},{m + 1}))", fmt=NUM)
        # MAT: this year's months to date plus the prior year's remaining months
        py_rest = (f"+SUM(INDEX(Data!$J$1:$U$40,11+{sel},{m + 2}):INDEX(Data!$J$1:$U$40,11+{sel},12))" if m < 11 else "")
        put(zs, f"I{rr}", f"=SUM(INDEX(Data!$J$1:$U$40,12+{sel},1):INDEX(Data!$J$1:$U$40,12+{sel},{m + 1})){py_rest}", fmt="0.0")
        put(zs, f"J{rr}", f"=MAX(IFERROR(F{rr},0),IFERROR(G{rr},0),H{rr},I{rr})", fmt=NUM)
    put(zs, "B20", "YTD variance to budget", F_BOLD)
    put(zs, "F20", "=INDEX(F7:F18,Ts_Last_Actual)-INDEX(H7:H18,Ts_Last_Actual)", F_BOLD, '+#,##0;-#,##0;"-"')
    put(zs, "B21", "Full year variance to budget", F_BOLD)
    put(zs, "F21", "=G18-H18", F_BOLD, '+#,##0;-#,##0;"-"')
    for c, w in zip("BCDEFGHIJ", [8, 7, 7, 8, 8, 8, 10, 7, 7]):
        zs.column_dimensions[c].width = w

    bars_bud = BarChart()
    bars_bud.type = "col"
    bars_bud.grouping = "clustered"
    bars_bud.gapWidth = 40
    bars_bud.add_data(Reference(zs, min_col=5, min_row=6, max_row=18), titles_from_data=True)
    bars_bud.set_categories(Reference(zs, min_col=2, min_row=7, max_row=18))
    s = bars_bud.series[0]
    s.graphicalProperties.solidFill = LIGHT
    s.graphicalProperties.line = no_line()

    bars_acfc = BarChart()
    bars_acfc.type = "col"
    bars_acfc.grouping = "clustered"
    bars_acfc.overlap = 100
    bars_acfc.gapWidth = 160
    bars_acfc.add_data(Reference(zs, min_col=3, max_col=4, min_row=6, max_row=18), titles_from_data=True)
    ac, fc = bars_acfc.series
    ac.graphicalProperties.solidFill = INK
    ac.graphicalProperties.line = no_line()
    fc.graphicalProperties.pattFill = PatternFillProperties(prst="wdUpDiag", fgClr=ColorChoice(srgbClr=INK), bgClr=ColorChoice(srgbClr="FFFFFF"))
    fc.graphicalProperties.line = LineProperties(solidFill=INK, w=9525)
    for ser in (ac, fc):
        ser.dLbls = DataLabelList()
        ser.dLbls.showVal = True
        ser.dLbls.showSerName = ser.dLbls.showCatName = ser.dLbls.showLegendKey = False
        ser.dLbls.position = "outEnd"
    bars_acfc.y_axis.axId = 200
    bars_acfc.y_axis.crosses = "max"
    bars_acfc.y_axis.delete = True

    lines = LineChart()
    lines.add_data(Reference(zs, min_col=6, max_col=9, min_row=6, max_row=18), titles_from_data=True)
    ytd_ac, ytd_fc, ytd_bud, mat = lines.series
    for ser, colour, dash, width in ((ytd_ac, INK, None, 22225), (ytd_fc, INK, "dash", 22225),
                                     (ytd_bud, GREY, None, 22225), (mat, INK, "sysDot", 15875)):
        ser.graphicalProperties.line = LineProperties(solidFill=colour, w=width, prstDash=dash)
        ser.marker.symbol = "circle"
        ser.marker.size = 4
        ser.marker.graphicalProperties.solidFill = colour
        ser.marker.graphicalProperties.line.solidFill = colour
        ser.smooth = False
    # Hidden series on the secondary axis carries the same maximum, so both axes scale alike.
    scale = LineChart()
    scale.add_data(Reference(zs, min_col=10, min_row=6, max_row=18), titles_from_data=True)
    scale.series[0].graphicalProperties.line = no_line()
    scale.series[0].marker.symbol = "none"
    scale.y_axis.axId = 200
    scale.y_axis.delete = True

    z = bars_bud
    z += bars_acfc
    z += lines
    z += scale
    z.title = None
    z.y_axis.delete = True
    z.y_axis.majorGridlines = None
    z.legend.position = "b"
    z.legend.legendEntry = [LegendEntry(idx=7, delete=True)]   # the hidden scale series
    z.height, z.width = 10, 22
    zs.add_chart(z, "L4")

    # Waterfalls --------------------------------------------------------------------
    wf = wb.create_sheet("Waterfall")
    style_sheet(wf, "Waterfall", "Bridges built as stacked columns: totals grey, increases green, decreases red")
    walk = [("Revenue", 1000, "total"), ("Materials", -350, "step"), ("Labour", -200, "step"), ("Other direct", -50, "step"),
            ("Gross profit", None, "total"), ("Staff", -120, "step"), ("Marketing", -45, "step"), ("Admin", -35, "step"),
            ("Other opex", -20, "step"), ("EBITDA", None, "total"), ("D&A", -30, "step"), ("Interest", -15, "step"),
            ("Tax", -40, "step"), ("Net profit", None, "total")]
    bva = [("Budget", 500, "total"), ("Revenue Δ", 85, "step"), ("COGS Δ", -30, "step"), ("Opex Δ", -15, "step"),
           ("Other Δ", 10, "step"), ("Actual", None, "total")]

    def waterfall(top_row, title, items, anchor):
        heads = ["Item", "Value", "Running", "Base", "Total", "Increase", "Decrease", "Label"]
        put(wf, f"B{top_row}", title, F_HEAD)
        for i, h in enumerate(heads):
            put(wf, f"{L(2 + i)}{top_row + 1}", h, F_BOLD)
        first = top_row + 2
        for k, (label, value, kind) in enumerate(items):
            rr = first + k
            put(wf, f"B{rr}", label)
            if kind == "total" and value is None:
                put(wf, f"C{rr}", f"=D{rr - 1}", fmt=NUM)          # subtotal = running total so far
                put(wf, f"D{rr}", f"=C{rr}", fmt=NUM)
            elif kind == "total":
                put(wf, f"C{rr}", value, fmt=NUM, fill=FILL_IN)
                put(wf, f"D{rr}", f"=C{rr}", fmt=NUM)
            else:
                put(wf, f"C{rr}", value, fmt=NUM, fill=FILL_IN)
                put(wf, f"D{rr}", f"=D{rr - 1}+C{rr}", fmt=NUM)
            is_total = kind == "total"
            put(wf, f"E{rr}", "=0" if is_total else f"=MIN(D{rr - 1},D{rr})", fmt=NUM)
            put(wf, f"F{rr}", f"=C{rr}" if is_total else "=0", fmt=NUM)
            put(wf, f"G{rr}", "=0" if is_total else f"=MAX(C{rr},0)", fmt=NUM)
            put(wf, f"H{rr}", "=0" if is_total else f"=MAX(-C{rr},0)", fmt=NUM)
            put(wf, f"I{rr}", f"=C{rr}", fmt='+#,##0;-#,##0;"-"' if not is_total else NUM)
        last = first + len(items) - 1
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "stacked"
        ch.overlap = 100
        ch.gapWidth = 30
        ch.add_data(Reference(wf, min_col=5, max_col=8, min_row=top_row + 1, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(wf, min_col=2, min_row=first, max_row=last))
        base, tot, inc, dec = ch.series
        base.graphicalProperties.noFill = True
        base.graphicalProperties.line = no_line()
        for ser, colour, fmt in ((tot, "808080", '#,##0;;'), (inc, GOOD, '"+"#,##0;;'), (dec, BAD, '"-"#,##0;;')):
            ser.graphicalProperties.solidFill = colour
            ser.graphicalProperties.line = LineProperties(solidFill="404040", w=6350)
            ser.dLbls = DataLabelList()
            ser.dLbls.showVal = True
            ser.dLbls.showSerName = ser.dLbls.showCatName = ser.dLbls.showLegendKey = False
            ser.dLbls.numFmt = fmt
            ser.dLbls.position = "inEnd" if ser is tot else "ctr"
        ch.legend = None
        ch.y_axis.delete = True
        ch.y_axis.majorGridlines = None
        ch.title = title
        ch.height, ch.width = 8.5, 20
        wf.add_chart(ch, anchor)
        return first, last

    w1 = waterfall(4, "P&L walk", walk, "L4")
    w2 = waterfall(22, "Budget to actual", bva, "L22")
    for c, w in zip("BCDEFGHI", [14, 8, 9, 8, 8, 9, 9, 8]):
        wf.column_dimensions[c].width = w

    # IBCS ---------------------------------------------------------------------------
    ib = wb.create_sheet("IBCS")
    style_sheet(ib, "IBCS", "Prior year grey, plan outlined, actual solid, forecast hatched; variances good green, bad red")
    put(ib, "B4", "Compare with", F_BOLD)
    put(ib, "E4", "Plan", fill=FILL_IN)
    dv = DataValidation(type="list", formula1="LU_Compare", allow_blank=False, showDropDown=False)
    dv.error, dv.errorTitle = "Choose Plan or Prior year", "Compare with"
    ib.add_data_validation(dv)
    dv.add("E4")
    name(wb, "DD_Compare", ib, "E4")
    heads = ["Line", "Cost line", "PY", "PL", "AC", "FC", "Compare", "Δ", "Good", "Bad"]
    for i, h in enumerate(heads):
        put(ib, f"{L(2 + i)}6", h, F_BOLD)
    lines_ib = [("Revenue", 0, 920, 1000, 1085, None), ("Materials", 1, 330, 350, 380, None), ("Labour", 1, 190, 200, 215, None),
                ("Overheads", 1, 140, 150, 140, None), ("EBITDA", 0, 260, 300, None, 350)]
    for k, (label, cost, py, pl, acv, fcv) in enumerate(lines_ib):
        rr = 7 + k
        put(ib, f"B{rr}", label)
        put(ib, f"C{rr}", cost, fmt="0", fill=FILL_IN)
        put(ib, f"D{rr}", py, fmt=NUM, fill=FILL_IN)
        put(ib, f"E{rr}", pl, fmt=NUM, fill=FILL_IN)
        put(ib, f"F{rr}", acv if acv is not None else "=NA()", fmt=NUM, fill=FILL_IN if acv is not None else None)
        put(ib, f"G{rr}", fcv if fcv is not None else "=NA()", fmt=NUM, fill=FILL_IN if fcv is not None else None)
        put(ib, f"H{rr}", f'=IF(DD_Compare="Prior year",D{rr},E{rr})', fmt=NUM)
        put(ib, f"I{rr}", f"=IFERROR(F{rr},G{rr})-H{rr}", fmt='+#,##0;-#,##0;"-"')
        put(ib, f"J{rr}", f"=IF((1-2*C{rr})*I{rr}>=0,I{rr},NA())", fmt=NUM)
        put(ib, f"K{rr}", f"=IF((1-2*C{rr})*I{rr}<0,I{rr},NA())", fmt=NUM)
    last_ib = 6 + len(lines_ib)
    col = BarChart()
    col.type = "col"
    col.grouping = "clustered"
    col.overlap = 30
    col.gapWidth = 70
    col.add_data(Reference(ib, min_col=4, max_col=7, min_row=6, max_row=last_ib), titles_from_data=True)
    col.set_categories(Reference(ib, min_col=2, min_row=7, max_row=last_ib))
    py_s, pl_s, ac_s, fc_s = col.series
    py_s.graphicalProperties.solidFill = GREY
    py_s.graphicalProperties.line = no_line()
    pl_s.graphicalProperties.noFill = True
    pl_s.graphicalProperties.line = LineProperties(solidFill=INK, w=12700)
    ac_s.graphicalProperties.solidFill = INK
    ac_s.graphicalProperties.line = no_line()
    fc_s.graphicalProperties.pattFill = PatternFillProperties(prst="wdUpDiag", fgClr=ColorChoice(srgbClr=INK), bgClr=ColorChoice(srgbClr="FFFFFF"))
    fc_s.graphicalProperties.line = LineProperties(solidFill=INK, w=9525)
    col.y_axis.delete = True
    col.y_axis.majorGridlines = None
    col.legend.position = "b"
    col.title = "PY, PL, AC and FC"
    col.height, col.width = 8, 16
    ib.add_chart(col, "M4")
    var = BarChart()
    var.type = "bar"
    var.grouping = "clustered"
    var.overlap = 100
    var.gapWidth = 60
    var.add_data(Reference(ib, min_col=10, max_col=11, min_row=6, max_row=last_ib), titles_from_data=True)
    var.set_categories(Reference(ib, min_col=2, min_row=7, max_row=last_ib))
    good, bad = var.series
    for ser, colour in ((good, GOOD), (bad, BAD)):
        ser.graphicalProperties.solidFill = colour
        ser.graphicalProperties.line = no_line()
        ser.dLbls = DataLabelList()
        ser.dLbls.showVal = True
        ser.dLbls.showSerName = ser.dLbls.showCatName = ser.dLbls.showLegendKey = False
        ser.dLbls.numFmt = '+#,##0;-#,##0;"0"'
    var.x_axis.scaling.orientation = "maxMin"
    var.x_axis.tickLblPos = "low"
    var.y_axis.delete = True
    var.y_axis.majorGridlines = None
    var.legend = None
    var.title = "Variance (green good, red bad)"
    var.height, var.width = 8, 12
    ib.add_chart(var, "M21")
    for c, w in zip("BCDEFGHIJK", [12, 8, 8, 8, 8, 8, 9, 8, 8, 8]):
        ib.column_dimensions[c].width = w

    # Checks -----------------------------------------------------------------------
    ck = wb.create_sheet("Checks")
    style_sheet(ck, "Checks", "Error checks must be nil")
    checks = [
        ("Line selection outside the list", "=IF(OR(DD_Chart_Line<1,DD_Chart_Line>ROWS(LU_Chart_Lines)),1,0)"),
        ("Last actual month outside the year", "=IF(OR(Ts_Last_Actual<1,Ts_Last_Actual>12),1,0)"),
        ("Z chart full year differs from the data", f"=IF(ABS('Z chart'!G18-INDEX(Data!$H$1:$H$40,12+(DD_Chart_Line-1)*5))>0.000001,1,0)"),
        ("P&L walk does not reach net profit", f"=IF(ABS(Waterfall!C{w1[1]}-SUM(Waterfall!C{w1[0]}:C{w1[0] + 3})-SUM(Waterfall!C{w1[0] + 5}:C{w1[0] + 8})-SUM(Waterfall!C{w1[0] + 10}:C{w1[0] + 12}))>0.000001,1,0)"),
        ("Budget to actual bridge does not close", f"=IF(ABS(Waterfall!C{w2[1]}-SUM(Waterfall!C{w2[0]}:C{w2[1] - 1}))>0.000001,1,0)"),
        ("Compare with is not Plan or Prior year", '=IF(COUNTIF(LU_Compare,DD_Compare)=0,1,0)'),
    ]
    for i, (label, f) in enumerate(checks):
        put(ck, f"C{4 + i}", label)
        put(ck, f"H{4 + i}", f, F_CHECK, "0")
    tot = 4 + len(checks)
    put(ck, f"C{tot}", "Error checks failing", F_BOLD)
    put(ck, f"H{tot}", f"=SUM(H4:H{tot - 1})", F_BOLD, "0", border=TOP)
    name(wb, "Chk_Errors", ck, f"H{tot}")
    ck.column_dimensions["C"].width = 40

    # Contents ---------------------------------------------------------------------
    toc = [("Data", "Monthly numbers, the line selector (classic drop-down) and the last actual month"),
           ("Z chart", "Z chart for the selected line"), ("Waterfall", "P&L walk and budget to actual bridges"),
           ("IBCS", "IBCS columns and variances, with the comparison chosen in-cell"),
           ("Lookups", "Lists behind the drop-downs"), ("Checks", "Error checks")]
    for i, (s, d) in enumerate(toc):
        c = put(contents, f"C{4 + i}", s)
        c.hyperlink = Hyperlink(ref=f"C{4 + i}", location=f"'{s}'!A1")
        put(contents, f"F{4 + i}", d)
    put(contents, "C11", "Error checks failing", F_BOLD)
    put(contents, "F11", "=Chk_Errors", F_CHECK, "0")
    contents.column_dimensions["C"].width = 14

    for ws in wb.worksheets:
        for rr in range(1, ws.max_row + 1):
            ws.row_dimensions[rr].height = 15
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    add_form_control(path, "Data", link="DD_Chart_Line", lst="LU_Chart_Lines", col=(3, 7), row=3, lines=len(names))
    return path


def add_form_control(path: Path, sheet: str, link: str, lst: str, col: tuple[int, int], row: int, lines: int) -> None:
    """Write a classic combo box (form control) into a saved workbook, as Excel stores one."""
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    wbx = parts["xl/workbook.xml"].decode()
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    rid = re.search(rf'<sheet [^>]*name="{re.escape(sheet)}"[^>]*r:id="(rId\d+)"', wbx).group(1)
    target = re.search(rf'<Relationship [^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels) or \
        re.search(rf'<Relationship [^>]*Target="([^"]+)"[^>]*Id="{rid}"', rels)
    sheet_part = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
    sheet_rels = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    c0, c1 = col
    shape_id = 1025
    parts["xl/ctrlProps/ctrlProp1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Drop" '
        f'dropLines="{lines}" dropStyle="combo" dx="26" fmlaLink="{link}" fmlaRange="{lst}" noThreeD="1" sel="1" val="0"/>').encode()
    parts["xl/drawings/vmlDrawing1.vml"] = (
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:x="urn:schemas-microsoft-com:office:excel">\n'
        ' <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
        '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">'
        '<v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>'
        '<o:lock v:ext="edit" shapetype="t"/></v:shapetype>'
        f'<v:shape id="_x0000_s{shape_id}" type="#_x0000_t201" style="position:absolute;margin-left:0;margin-top:0;width:150pt;height:15pt;z-index:1;mso-wrap-style:tight" '
        'stroked="f" strokecolor="windowText [64]" o:insetmode="auto"><o:lock v:ext="edit" rotation="t" text="t"/>'
        '<x:ClientData ObjectType="Drop"><x:SizeWithCells/>'
        f'<x:Anchor>{c0}, 0, {row}, 0, {c1}, 0, {row + 1}, 0</x:Anchor><x:AutoLine>False</x:AutoLine>'
        f'<x:FmlaLink>{link}</x:FmlaLink><x:Val>0</x:Val><x:Min>0</x:Min><x:Max>{lines}</x:Max><x:Inc>1</x:Inc><x:Page>{lines}</x:Page>'
        f'<x:Dx>26</x:Dx><x:FmlaRange>{lst}</x:FmlaRange><x:Sel>1</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType>'
        f'<x:LCT>Normal</x:LCT><x:DropStyle>Combo</x:DropStyle><x:DropLines>{lines}</x:DropLines></x:ClientData></v:shape></xml>').encode()
    anchor = (f'<xdr:from><xdr:col>{c0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
              f'<xdr:to><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row + 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>')
    parts["xl/drawings/drawingControls1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
        f'<xdr:twoCellAnchor editAs="oneCell">{anchor}<xdr:sp macro="" textlink=""><xdr:nvSpPr>'
        f'<xdr:cNvPr id="{shape_id}" name="Drop Down 1" hidden="1"><a:extLst><a:ext uri="{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}">'
        f'<a14:compatExt spid="_x0000_s{shape_id}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>'
        '<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>'
        '</xdr:wsDr>').encode()
    srels = parts.get(sheet_rels, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>').decode()
    if "/drawing\"" in srels or 'relationships/drawing"' in srels:
        raise RuntimeError(f"{sheet} already has a drawing; put the control on a sheet without charts")
    srels = srels.replace("</Relationships>",
        '<Relationship Id="rIdCtlD" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawingControls1.xml"/>'
        '<Relationship Id="rIdCtlV" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
        '<Relationship Id="rIdCtlP" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp1.xml"/>'
        "</Relationships>")
    parts[sheet_rels] = srels.encode()
    sx = parts[sheet_part].decode()
    if 'xmlns:r="' not in sx[:600]:
        sx = sx.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    if 'xmlns:xdr="' not in sx[:800]:
        sx = sx.replace("<worksheet ", '<worksheet xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ', 1)
    controls = ('<drawing r:id="rIdCtlD"/><legacyDrawing r:id="rIdCtlV"/>'
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls>'
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14">'
                f'<control shapeId="{shape_id}" r:id="rIdCtlP" name="Drop Down 1"><controlPr defaultSize="0" autoLine="0" autoPict="0">'
                f'<anchor moveWithCells="1">{anchor}</anchor></controlPr></control></mc:Choice></mc:AlternateContent>'
                '</controls></mc:Choice></mc:AlternateContent>')
    tail = re.search(r"(<tableParts|<extLst|</worksheet>)", sx)
    sx = sx[:tail.start()] + controls + sx[tail.start():]
    parts[sheet_part] = sx.encode()
    ct = parts["[Content_Types].xml"].decode()
    if 'Extension="vml"' not in ct:
        ct = ct.replace("</Types>", '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')
    ct = ct.replace("</Types>", '<Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>'
                    '<Override PartName="/xl/drawings/drawingControls1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    parts["[Content_Types].xml"] = ct.encode()
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    import os
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    shutil.move(tmp, path)


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[2] / "build" / "charts" / "charts_demo.xlsx"
    print(f"Built {build(out)}")
