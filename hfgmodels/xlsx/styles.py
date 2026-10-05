"""Style book: maps frame style keys to cell format (xf) indices in the template.

The template's own cell formats are reused wherever one exists, so new sheets
look exactly like the template's sheets and follow its theme colours. Where the
template has no suitable format (for example an input cell formatted as a
date), a new format is derived from an existing one and appended to
styles.xml. Existing format indices are never changed, so every template sheet
keeps its look.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field

from lxml import etree

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = {"m": NS}


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


@dataclass
class Derived:
    """A format derived from a base format."""

    base: str | int
    num_fmt: str | None = None  # format code, added to numFmts if new
    bold: bool | None = None
    font_rgb: str | None = None  # e.g. "FFC00000"
    border_top: str | None = None  # "thin", "dashed", "double"
    border_bottom: str | None = None
    h_align: str | None = None  # "left", "center", "right"
    indent: int | None = None
    fill_rgb: str | None = None
    extra: dict = field(default_factory=dict)


class StyleBook:
    def __init__(self, styles_xml: bytes, base_map: dict[str, int], derived: dict[str, Derived]):
        self.root = etree.fromstring(styles_xml)
        self.cell_xfs = self.root.find("m:cellXfs", N)
        self.fonts = self.root.find("m:fonts", N)
        self.borders = self.root.find("m:borders", N)
        self.fills = self.root.find("m:fills", N)
        self.num_fmts = self.root.find("m:numFmts", N)
        if self.num_fmts is None:
            self.num_fmts = etree.Element(_q("numFmts"), count="0")
            self.root.insert(0, self.num_fmts)
        self.keys: dict[str, int] = {}
        n = len(self.cell_xfs)
        for k, idx in base_map.items():
            if not 0 <= idx < n:
                raise ValueError(f"style {k}: index {idx} outside the template's {n} formats")
            self.keys[k] = idx
        self._derived_specs = derived
        for k in derived:
            self._resolve(k)
        self.dirty = bool(derived)

    # ------------------------------------------------------------------
    def __getitem__(self, key: str) -> int:
        if key not in self.keys:
            raise KeyError(f"unknown style key '{key}'")
        return self.keys[key]

    def describe(self, idx: int) -> str:
        xf = self.cell_xfs[idx]
        return etree.tostring(xf).decode()

    def _resolve(self, key: str) -> int:
        if key in self.keys:
            return self.keys[key]
        spec = self._derived_specs[key]
        base = spec.base if isinstance(spec.base, int) else self._resolve(spec.base)
        xf = copy.deepcopy(self.cell_xfs[base])
        if spec.num_fmt is not None:
            xf.set("numFmtId", str(self._num_fmt_id(spec.num_fmt)))
            xf.set("applyNumberFormat", "1")
        if spec.bold is not None or spec.font_rgb is not None:
            xf.set("fontId", str(self._font(int(xf.get("fontId", "0")), spec.bold, spec.font_rgb)))
            xf.set("applyFont", "1")
        if spec.border_top is not None or spec.border_bottom is not None:
            xf.set("borderId", str(self._border(int(xf.get("borderId", "0")), spec.border_top, spec.border_bottom)))
            xf.set("applyBorder", "1")
        if spec.fill_rgb is not None:
            xf.set("fillId", str(self._fill(spec.fill_rgb)))
            xf.set("applyFill", "1")
        if spec.h_align is not None or spec.indent is not None:
            al = xf.find("m:alignment", N)
            if al is None:
                al = etree.SubElement(xf, _q("alignment"))
                # alignment must precede protection
                prot = xf.find("m:protection", N)
                if prot is not None:
                    xf.remove(al)
                    prot.addprevious(al)
            if spec.h_align is not None:
                al.set("horizontal", spec.h_align)
            if spec.indent is not None:
                al.set("indent", str(spec.indent))
            xf.set("applyAlignment", "1")
        self.cell_xfs.append(xf)
        self.cell_xfs.set("count", str(len(self.cell_xfs)))
        idx = len(self.cell_xfs) - 1
        self.keys[key] = idx
        return idx

    def _num_fmt_id(self, code: str) -> int:
        builtin = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10}
        if code in builtin:
            return builtin[code]
        used = []
        for nf in self.num_fmts.findall("m:numFmt", N):
            used.append(int(nf.get("numFmtId")))
            if nf.get("formatCode") == code:
                return int(nf.get("numFmtId"))
        new_id = max([163] + used) + 1
        etree.SubElement(self.num_fmts, _q("numFmt"), numFmtId=str(new_id), formatCode=code)
        self.num_fmts.set("count", str(len(self.num_fmts)))
        return new_id

    def _font(self, base_id: int, bold: bool | None, rgb: str | None) -> int:
        f = copy.deepcopy(self.fonts[base_id])
        if bold is not None:
            b = f.find("m:b", N)
            if bold and b is None:
                f.insert(0, etree.Element(_q("b")))
            if not bold and b is not None:
                f.remove(b)
        if rgb is not None:
            c = f.find("m:color", N)
            if c is None:
                c = etree.Element(_q("color"))
                sz = f.find("m:sz", N)
                (sz.addnext(c) if sz is not None else f.append(c))
            c.attrib.clear()
            c.set("rgb", rgb)
        self.fonts.append(f)
        self.fonts.set("count", str(len(self.fonts)))
        return len(self.fonts) - 1

    def _border(self, base_id: int, top: str | None, bottom: str | None) -> int:
        b = copy.deepcopy(self.borders[base_id])
        for side, style in (("top", top), ("bottom", bottom)):
            if style is None:
                continue
            e = b.find(f"m:{side}", N)
            if e is None:
                e = etree.SubElement(b, _q(side))
            e.attrib.clear()
            for ch in list(e):
                e.remove(ch)
            if style != "none":
                e.set("style", style)
                etree.SubElement(e, _q("color"), auto="1")
        # keep child order: left, right, top, bottom, diagonal
        order = {"left": 0, "start": 0, "right": 1, "end": 1, "top": 2, "bottom": 3, "diagonal": 4,
                 "vertical": 5, "horizontal": 6}
        kids = sorted(list(b), key=lambda e: order.get(etree.QName(e).localname, 9))
        for k in kids:
            b.remove(k)
        for k in kids:
            b.append(k)
        self.borders.append(b)
        self.borders.set("count", str(len(self.borders)))
        return len(self.borders) - 1

    def _fill(self, rgb: str) -> int:
        f = etree.SubElement(self.fills, _q("fill"))
        p = etree.SubElement(f, _q("patternFill"), patternType="solid")
        etree.SubElement(p, _q("fgColor"), rgb=rgb)
        etree.SubElement(p, _q("bgColor"), indexed="64")
        self.fills.set("count", str(len(self.fills)))
        return len(self.fills) - 1

    def to_bytes(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
