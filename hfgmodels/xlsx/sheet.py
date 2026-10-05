"""Worksheet XML writer.

A Sheet collects cells (values, formulas, style indices), row properties,
column widths, freeze panes, hyperlinks and data validation, then renders a
SpreadsheetML worksheet part. Strings are written inline so the template's
shared string table is never touched.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

from .cells import col_letter

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EXCEL_EPOCH = dt.date(1899, 12, 30)


def excel_date(d: dt.date) -> int:
    return (d - EXCEL_EPOCH).days


@dataclass
class Cell:
    value: object = None
    formula: str | None = None
    style: int | None = None


@dataclass
class RowProps:
    height: float | None = None
    outline: int = 0
    hidden: bool = False


@dataclass
class Sheet:
    name: str
    default_row_height: float = 15.0
    cells: dict = field(default_factory=dict)
    rows: dict = field(default_factory=dict)
    cols: list = field(default_factory=list)  # (min, max, width, style)
    freeze: tuple | None = None  # (first unfrozen row, first unfrozen col)
    hyperlinks: list = field(default_factory=list)  # (ref, location, display)
    validations: list = field(default_factory=list)  # (sqref, list_formula)
    tab_rgb: str | None = None
    show_grid: bool = False
    zoom: int = 85
    outline_summary_below: bool = False

    def set(self, row: int, col: int, value=None, formula: str | None = None, style: int | None = None):
        if formula is not None and formula.startswith("="):
            formula = formula[1:]
        c = self.cells.get((row, col))
        if c is None:
            c = Cell()
            self.cells[(row, col)] = c
        if value is not None:
            c.value = value
        if formula is not None:
            c.formula = formula
        if style is not None:
            c.style = style
        return c

    def style_range(self, row: int, c1: int, c2: int, style: int):
        for c in range(c1, c2 + 1):
            self.set(row, c, style=style)

    def row_props(self, row: int, height: float | None = None, outline: int | None = None, hidden: bool | None = None):
        p = self.rows.setdefault(row, RowProps())
        if height is not None:
            p.height = height
        if outline is not None:
            p.outline = outline
        if hidden is not None:
            p.hidden = hidden

    @property
    def max_row(self) -> int:
        return max([r for r, _ in self.cells] + [1])

    @property
    def max_col(self) -> int:
        return max([c for _, c in self.cells] + [1])

    # ------------------------------------------------------------------
    def to_xml(self, values: dict | None = None) -> bytes:
        """Render the worksheet. `values` optionally maps (row, col) to a cached
        result for formula cells, so readers that do not calculate still see
        numbers."""
        values = values or {}
        out = []
        w = out.append
        w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        w(f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">')
        pr = '<sheetPr>'
        if self.tab_rgb:
            pr += f'<tabColor rgb="{self.tab_rgb}"/>'
        pr += f'<outlinePr summaryBelow="{1 if self.outline_summary_below else 0}"/></sheetPr>'
        w(pr)
        w(f'<dimension ref="A1:{col_letter(self.max_col)}{self.max_row}"/>')
        view = f'<sheetView showGridLines="{1 if self.show_grid else 0}" zoomScale="{self.zoom}" zoomScaleNormal="{self.zoom}" workbookViewId="0">'
        if self.freeze:
            fr, fc = self.freeze
            xs, ys = fc - 1, fr - 1
            tl = f"{col_letter(fc)}{fr}"
            attrs = []
            if xs:
                attrs.append(f'xSplit="{xs}"')
            if ys:
                attrs.append(f'ySplit="{ys}"')
            pane = "bottomRight" if xs and ys else ("topRight" if xs else "bottomLeft")
            view += f'<pane {" ".join(attrs)} topLeftCell="{tl}" activePane="{pane}" state="frozen"/>'
            view += f'<selection pane="{pane}" activeCell="{tl}" sqref="{tl}"/>'
        else:
            view += '<selection activeCell="B4" sqref="B4"/>'
        view += "</sheetView>"
        w(f"<sheetViews>{view}</sheetViews>")
        max_outline = max([p.outline for p in self.rows.values()] + [0])
        fmt = f'<sheetFormatPr defaultRowHeight="{self.default_row_height}" customHeight="1"'
        if max_outline:
            fmt += f' outlineLevelRow="{max_outline}"'
        w(fmt + "/>")
        if self.cols:
            w("<cols>")
            for mn, mx, width, style in sorted(self.cols):
                s = f' style="{style}"' if style is not None else ""
                w(f'<col min="{mn}" max="{mx}" width="{width}" customWidth="1"{s}/>')
            w("</cols>")
        w("<sheetData>")
        by_row: dict[int, list] = {}
        for (r, c), cell in self.cells.items():
            by_row.setdefault(r, []).append((c, cell))
        all_rows = sorted(set(by_row) | set(self.rows))
        for r in all_rows:
            p = self.rows.get(r, RowProps())
            ht = p.height if p.height is not None else self.default_row_height
            attrs = f'r="{r}" ht="{ht}" customHeight="1"'
            if p.outline:
                attrs += f' outlineLevel="{p.outline}"'
            if p.hidden:
                attrs += ' hidden="1"'
            cells = sorted(by_row.get(r, []), key=lambda x: x[0])
            if not cells:
                w(f"<row {attrs}/>")
                continue
            w(f"<row {attrs}>")
            for c, cell in cells:
                w(self._cell_xml(r, c, cell, values.get((r, c))))
            w("</row>")
        w("</sheetData>")
        # element order follows CT_Worksheet: dataValidations before hyperlinks
        if self.validations:
            w(f'<dataValidations count="{len(self.validations)}">')
            for sqref, lst in self.validations:
                w(f'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="{sqref}">'
                  f"<formula1>{escape(lst)}</formula1></dataValidation>")
            w("</dataValidations>")
        if self.hyperlinks:
            w("<hyperlinks>")
            for ref, loc, disp in self.hyperlinks:
                w(f'<hyperlink ref="{ref}" location="{escape(loc)}" display="{escape(disp)}"/>')
            w("</hyperlinks>")
        w('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
        w('<pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/>')
        w("</worksheet>")
        return "".join(out).encode("utf-8")

    @staticmethod
    def _cell_xml(r: int, c: int, cell: Cell, cached) -> str:
        a = f'r="{col_letter(c)}{r}"'
        if cell.style is not None:
            a += f' s="{cell.style}"'
        if cell.formula is not None:
            f = f"<f>{escape(cell.formula)}</f>"
            if cached is None:
                return f"<c {a}>{f}</c>"
            if isinstance(cached, bool):
                return f'<c {a} t="b">{f}<v>{int(cached)}</v></c>'
            if isinstance(cached, (int, float)):
                return f"<c {a}>{f}<v>{repr(float(cached)) if isinstance(cached, float) else cached}</v></c>"
            if isinstance(cached, str) and cached.startswith("#"):
                return f'<c {a} t="e">{f}<v>{escape(cached)}</v></c>'
            return f'<c {a} t="str">{f}<v>{escape(str(cached))}</v></c>'
        v = cell.value
        if v is None:
            return f"<c {a}/>"
        if isinstance(v, bool):
            return f'<c {a} t="b"><v>{int(v)}</v></c>'
        if isinstance(v, dt.datetime):
            v = v.date()
        if isinstance(v, dt.date):
            return f"<c {a}><v>{excel_date(v)}</v></c>"
        if isinstance(v, (int, float)):
            return f"<c {a}><v>{v}</v></c>"
        s = str(v)
        space = ' xml:space="preserve"' if s != s.strip() else ""
        return f'<c {a} t="inlineStr"><is><t{space}>{escape(s)}</t></is></c>'
