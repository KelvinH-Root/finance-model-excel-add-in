"""Row layout engine.

Modules describe their rows (sections, inputs, time series calculations,
checks) with keys. Layout happens in two passes: first every sheet's rows are
numbered, then formulas are rendered, so any row can refer to any other row by
key on any sheet without worrying about build order.

A formula function receives a context `x` with helpers:
  x.c            column letter of the current period
  x.t            current period number (1 based)
  x.cnt          counter cell of the current column (row 9), e.g. "J$9"
  x.v(key)       the key's value in the current column (same sheet or qualified)
  x.prev(key)    the key's value in the previous column
  x.rng(key)     the key's full timeline range, absolute
  x.tot(key)     the key's total cell (column I)
  x.cell(key, col) a specific column of a scalar or table row
  x.s(key)       a scalar: its defined name if it has one, else its H cell
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable

from . import frame
from .xlsx.cells import col_index, col_letter, quote_sheet
from .xlsx.sheet import Sheet

INDENT_COLS = [2, 3, 4, 5, 6, 7]  # B to G


@dataclass
class Row:
    kind: str  # section, sub, head, blank, ts, scalar, table
    key: str | None = None
    label: object = ""  # text or ("formula", "...")
    indent: int = 0
    style: str = "num"
    unit: str = ""
    total: str | None = "sum"  # ts rows: sum, max, min, last, none
    total_style: str | None = None
    fn: Callable | None = None
    value: object = None
    name: str | None = None
    input: bool = False
    validation: str | None = None
    cells: dict = field(default_factory=dict)  # table rows: col letter -> (value|fn|("f", str), style, validation)
    outline: int = 0
    row: int = 0


class SheetLayout:
    def __init__(self, model: "Model", name: str, title: str, timeline: bool, after: str | None = None,
                 tab_rgb: str | None = None, last_col: str | None = None):
        self.model = model
        self.name = name
        self.title = title
        self.timeline = timeline
        self.after = after
        self.tab_rgb = tab_rgb
        self.rows: list[Row] = []
        self.r = frame.CONTENT_START_ROW if timeline else 5
        self.extra_cols: list = []  # (min, max, width) for non-timeline sheets
        self.last_col = last_col
        model.sheets.append(self)

    # ---------------------------------------------------------------- add
    def _add(self, row: Row, height: int = 1) -> Row:
        row.row = self.r
        self.rows.append(row)
        if row.key:
            if row.key in self.model.index:
                raise ValueError(f"duplicate key {row.key}")
            self.model.index[row.key] = (self.name, row.row)
        self.r += height
        return row

    def section(self, label, key: str | None = None):
        if self.r > (frame.CONTENT_START_ROW if self.timeline else 5):
            self.r += 1
        self._add(Row("section", key=key, label=label))
        self.r += 1

    def sub(self, label):
        if self.rows and self.rows[-1].kind not in ("section", "sub") and self.rows[-1].row == self.r - 1:
            self.r += 1
        self._add(Row("sub", label=label))
        self.r += 1

    def head(self, label, indent: int = 0, cells: dict | None = None):
        self._add(Row("head", label=label, indent=indent, cells=cells or {}))

    def blank(self, n: int = 1):
        self.r += n

    def ts(self, key, label, fn, style="num", unit="$", total="sum", indent=1, total_style=None, outline=0):
        return self._add(Row("ts", key=key, label=label, fn=fn, style=style, unit=unit, total=total, indent=indent,
                             total_style=total_style, outline=outline))

    def scalar(self, key, label, value=None, fn=None, style="num", unit="", indent=1, name=None, input=False,
               validation=None):
        return self._add(Row("scalar", key=key, label=label, value=value, fn=fn, style=style, unit=unit,
                             indent=indent, name=name, input=input, validation=validation))

    def table(self, key, label, cells: dict, indent=1):
        return self._add(Row("table", key=key, label=label, cells=cells, indent=indent))


class Ctx:
    def __init__(self, model: "Model", sheet: SheetLayout, t: int | None):
        self.model = model
        self.sheet = sheet
        self.t = t
        self.c = frame.ts_letter(t) if t else None
        self.cnt = f"{self.c}$9" if t else None
        self.first = t == 1
        self.last_letter = frame.ts_letter(model.n)

    def _where(self, key):
        try:
            return self.model.index[key]
        except KeyError:
            raise KeyError(f"unknown row key '{key}'") from None

    def _q(self, sheet_name: str) -> str:
        return "" if sheet_name == self.sheet.name else quote_sheet(sheet_name) + "!"

    def v(self, key, lag: int = 0) -> str:
        sh, r = self._where(key)
        t = self.t - lag
        if t < 1:
            raise ValueError(f"{key}: period {t} before the timeline")
        return f"{self._q(sh)}{frame.ts_letter(t)}{r}"

    def prev(self, key) -> str:
        return self.v(key, 1)

    def rng(self, key) -> str:
        sh, r = self._where(key)
        return f"{self._q(sh)}$J${r}:${self.last_letter}${r}"

    def tot(self, key) -> str:
        sh, r = self._where(key)
        return f"{self._q(sh)}$I${r}"

    def last(self, key) -> str:
        sh, r = self._where(key)
        return f"{self._q(sh)}${self.last_letter}${r}"

    def cell(self, key, col: str) -> str:
        sh, r = self._where(key)
        return f"{self._q(sh)}${col}${r}"

    def s(self, key) -> str:
        if key in self.model.key_names:
            return self.model.key_names[key]
        return self.cell(key, "H")


class Model:
    def __init__(self, n_periods: int):
        self.n = n_periods
        self.sheets: list[SheetLayout] = []
        self.index: dict[str, tuple[str, int]] = {}
        self.key_names: dict[str, str] = {}
        self.names: dict[str, str] = {}  # defined name -> address

    def sheet(self, *a, **k) -> SheetLayout:
        return SheetLayout(self, *a, **k)

    def register_names(self):
        for sh in self.sheets:
            for row in sh.rows:
                if row.kind == "scalar" and row.name:
                    self.key_names[row.key] = row.name
                    self.names[row.name] = f"{quote_sheet(sh.name)}!$H${row.row}"

    # ------------------------------------------------------------- render
    def render(self, st) -> dict[str, Sheet]:
        self.register_names()
        out = {}
        for sh in self.sheets:
            out[sh.name] = self._render_sheet(sh, st)
        return out

    def _render_sheet(self, sl: SheetLayout, st) -> Sheet:
        n = self.n if sl.timeline else None
        ws = Sheet(sl.name, tab_rgb=sl.tab_rgb)
        frame.write_header(ws, st, sl.title, n)
        if not sl.timeline:
            ws.freeze = (5, 1)
            for mn, mx, width in sl.extra_cols:
                ws.cols.append((mn, mx, width, None))
        last_c = frame.ts_col(self.n) if sl.timeline else col_index(sl.last_col or "H")
        for row in sl.rows:
            r = row.row
            if row.kind in ("section", "sub"):
                key = "section" if row.kind == "section" else "sub"
                self._label(ws, r, frame.LABEL_COL, row.label, st[key])
                for c in range(frame.LABEL_COL + 1, last_c + 1):
                    ws.set(r, c, style=st[key])
                continue
            col = INDENT_COLS[min(row.indent, len(INDENT_COLS) - 1)]
            if row.kind == "head":
                self._label(ws, r, col, row.label, st["h3"])
                for letter, spec in row.cells.items():
                    val, style = spec[0], spec[1]
                    self._put(ws, r, col_index(letter), val, st[style], Ctx(self, sl, None))
                continue
            self._label(ws, r, col, row.label, st["label"])
            if row.kind == "ts":
                if row.unit:
                    ws.set(r, frame.UNIT_COL, row.unit, style=st["unit"])
                for t in range(1, self.n + 1):
                    x = Ctx(self, sl, t)
                    f = row.fn(x)
                    c = frame.ts_col(t)
                    if isinstance(f, (int, float)):
                        ws.set(r, c, f, style=st[row.style])
                    else:
                        ws.set(r, c, formula=f, style=st[row.style])
                tot_style = st[row.total_style or row.style]
                rng = f"J{r}:{frame.ts_letter(self.n)}{r}"
                if row.total == "sum":
                    ws.set(r, frame.TOTAL_COL, formula=f"SUM({rng})", style=tot_style)
                elif row.total == "max":
                    ws.set(r, frame.TOTAL_COL, formula=f"MAX({rng})", style=tot_style)
                elif row.total == "min":
                    ws.set(r, frame.TOTAL_COL, formula=f"MIN({rng})", style=tot_style)
                elif row.total == "last":
                    ws.set(r, frame.TOTAL_COL, formula=f"{frame.ts_letter(self.n)}{r}", style=tot_style)
                if row.outline:
                    ws.row_props(r, outline=row.outline)
            elif row.kind == "scalar":
                x = Ctx(self, sl, None)
                if row.unit:
                    ws.set(r, frame.TOTAL_COL, row.unit, style=st["text"])
                val = row.fn(x) if row.fn else row.value
                self._put(ws, r, frame.UNIT_COL, val, st[row.style], x)
                if row.validation:
                    ws.validations.append((f"H{r}", row.validation))
            elif row.kind == "table":
                x = Ctx(self, sl, None)
                for letter, spec in row.cells.items():
                    val, style = spec[0], spec[1]
                    v = val(x) if callable(val) else val
                    self._put(ws, r, col_index(letter), v, st[style], x)
                    if len(spec) > 2 and spec[2]:
                        ws.validations.append((f"{letter}{r}", spec[2]))
        return ws

    @staticmethod
    def _label(ws: Sheet, r: int, c: int, label, style):
        if isinstance(label, tuple) and label[0] == "f":
            ws.set(r, c, formula=label[1], style=style)
        else:
            ws.set(r, c, label, style=style)

    @staticmethod
    def _put(ws: Sheet, r: int, c: int, val, style, x):
        if isinstance(val, tuple) and val and val[0] == "f":
            ws.set(r, c, formula=val[1], style=style)
        elif isinstance(val, str) and val.startswith("="):
            ws.set(r, c, formula=val[1:], style=style)
        else:
            ws.set(r, c, val, style=style)
