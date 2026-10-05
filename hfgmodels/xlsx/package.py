"""Package-level editing of an existing .xlsx template.

openpyxl drops form controls, charts, Python in Excel results, web extensions
and other parts it does not understand, so it cannot be used to edit the
template. This module works on the zip package directly: every part of the
template is kept byte for byte unless the build changes it on purpose, and new
worksheet parts are added alongside.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

from .cells import parse_ref

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
WS_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
WS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
N = {"m": NS, "r": RNS}


def _q(tag: str, ns: str = NS) -> str:
    return f"{{{ns}}}{tag}"


class TemplatePackage:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.order = z.namelist()
            self.parts: dict[str, bytes] = {n: z.read(n) for n in self.order}
        self.wb = etree.fromstring(self.parts["xl/workbook.xml"])
        self.rels = etree.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        self.ct = etree.fromstring(self.parts["[Content_Types].xml"])
        self._sheet_cache: dict[str, etree._Element] = {}

    # -------------------------------------------------------------- sheets
    def _sheets_el(self):
        return self.wb.find("m:sheets", N)

    def sheet_names(self) -> list[str]:
        return [s.get("name") for s in self._sheets_el()]

    def sheet_part(self, name: str) -> str:
        for s in self._sheets_el():
            if s.get("name") == name:
                rid = s.get(_q("id", RNS))
                for rel in self.rels:
                    if rel.get("Id") == rid:
                        t = rel.get("Target")
                        return t.lstrip("/") if t.startswith("/") else "xl/" + t
        raise KeyError(f"sheet not found: {name}")

    def add_sheet(self, name: str, xml: bytes, after: str | None = None) -> str:
        if name in self.sheet_names():
            raise ValueError(f"sheet already exists: {name}")
        if len(name) > 31:
            raise ValueError(f"sheet name longer than 31 characters: {name}")
        nums = [int(m.group(1)) for p in self.parts if (m := re.match(r"xl/worksheets/sheet(\d+)\.xml$", p))]
        part = f"xl/worksheets/sheet{max(nums + [0]) + 1}.xml"
        self.parts[part] = xml
        rid_nums = [int(m.group(1)) for rel in self.rels if (m := re.match(r"rId(\d+)$", rel.get("Id", "")))]
        rid = f"rId{max(rid_nums + [0]) + 1}"
        etree.SubElement(self.rels, _q("Relationship", PRNS), Id=rid, Type=WS_TYPE,
                         Target=part[len("xl/"):])
        etree.SubElement(self.ct, _q("Override", CTNS), PartName="/" + part, ContentType=WS_CT)
        sheets = self._sheets_el()
        sheet_id = max(int(s.get("sheetId")) for s in sheets) + 1
        el = etree.Element(_q("sheet"))
        el.set("name", name)
        el.set("sheetId", str(sheet_id))
        el.set(_q("id", RNS), rid)
        names = self.sheet_names()
        pos = (names.index(after) + 1) if after else len(names)
        sheets.insert(pos, el)
        self._shift_local_ids(pos)
        return part

    def replace_sheet_xml(self, name: str, xml: bytes):
        self.parts[self.sheet_part(name)] = xml
        self._sheet_cache.pop(name, None)

    def _shift_local_ids(self, pos: int):
        dn = self.wb.find("m:definedNames", N)
        if dn is not None:
            for d in dn:
                lid = d.get("localSheetId")
                if lid is not None and int(lid) >= pos:
                    d.set("localSheetId", str(int(lid) + 1))
        for bv in self.wb.iterfind("m:bookViews/m:workbookView", N):
            for attr in ("activeTab", "firstSheet"):
                v = bv.get(attr)
                if v is not None and int(v) >= pos:
                    bv.set(attr, str(int(v) + 1))

    def set_active_sheet(self, name: str):
        idx = self.sheet_names().index(name)
        for bv in self.wb.iterfind("m:bookViews/m:workbookView", N):
            bv.set("activeTab", str(idx))
            bv.set("firstSheet", "0")
            bv.attrib.pop("minimized", None)

    # ------------------------------------------------------- defined names
    def _dn_el(self):
        dn = self.wb.find("m:definedNames", N)
        if dn is None:
            dn = etree.Element(_q("definedNames"))
            self._sheets_el().addnext(dn)
        return dn

    def defined_names(self) -> dict[str, str]:
        dn = self.wb.find("m:definedNames", N)
        if dn is None:
            return {}
        return {d.get("name"): d.text for d in dn if d.get("localSheetId") is None}

    def set_defined_name(self, name: str, ref: str, comment: str | None = None):
        dn = self._dn_el()
        for d in dn:
            if d.get("name") == name and d.get("localSheetId") is None:
                d.text = ref
                return
        el = etree.SubElement(dn, _q("definedName"), name=name)
        if comment:
            el.set("comment", comment)
        el.text = ref

    # ---------------------------------------------------- existing cells
    def _sheet_root(self, name: str):
        if name not in self._sheet_cache:
            self._sheet_cache[name] = etree.fromstring(self.parts[self.sheet_part(name)])
        return self._sheet_cache[name]

    def set_cell_value(self, sheet: str, a1: str, value):
        """Set a constant in an existing template cell, keeping its style."""
        root = self._sheet_root(sheet)
        r, _ = parse_ref(a1)
        sd = root.find("m:sheetData", N)
        row = next((x for x in sd if x.get("r") == str(r)), None)
        if row is None:
            raise KeyError(f"{sheet}!{a1}: row not present in template")
        c = next((x for x in row if x.get("r") == a1), None)
        if c is None:
            raise KeyError(f"{sheet}!{a1}: cell not present in template")
        for ch in list(c):
            c.remove(ch)
        c.attrib.pop("t", None)
        if isinstance(value, str):
            c.set("t", "inlineStr")
            is_ = etree.SubElement(c, _q("is"))
            etree.SubElement(is_, _q("t")).text = value
        else:
            etree.SubElement(c, _q("v")).text = str(value)

    def append_cells(self, sheet: str, cells: list[tuple[str, object, int | None]], merges: list[str] = (),
                     links: list[tuple[str, str, str]] = (), row_height: float | None = None):
        """Add new cells below an existing sheet's content (new rows only).

        cells: (A1, value, style index); merges: ranges; links: (A1, location, display).
        """
        from .cells import col_letter  # noqa: F401
        root = self._sheet_root(sheet)
        sd = root.find("m:sheetData", N)
        existing = {int(r.get("r")) for r in sd}
        by_row: dict[int, list] = {}
        for a1, value, style in cells:
            r, c = parse_ref(a1)
            if r in existing:
                raise ValueError(f"{sheet}!{a1}: row {r} already exists in the template")
            by_row.setdefault(r, []).append((c, a1, value, style))
        for r in sorted(by_row):
            row = etree.SubElement(sd, _q("row"), r=str(r))
            if row_height:
                row.set("ht", str(row_height))
                row.set("customHeight", "1")
            for c, a1, value, style in sorted(by_row[r], key=lambda x: x[0]):
                el = etree.SubElement(row, _q("c"), r=a1)
                if style is not None:
                    el.set("s", str(style))
                if value is None:
                    continue
                if isinstance(value, str):
                    el.set("t", "inlineStr")
                    is_ = etree.SubElement(el, _q("is"))
                    etree.SubElement(is_, _q("t")).text = value
                else:
                    etree.SubElement(el, _q("v")).text = str(value)
        if merges:
            mc = root.find("m:mergeCells", N)
            if mc is None:
                raise ValueError(f"{sheet}: add a mergeCells element first")
            for ref in merges:
                etree.SubElement(mc, _q("mergeCell"), ref=ref)
            mc.set("count", str(len(mc)))
        if links:
            hl = root.find("m:hyperlinks", N)
            if hl is None:
                raise ValueError(f"{sheet}: template sheet has no hyperlinks element")
            for a1, loc, disp in links:
                etree.SubElement(hl, _q("hyperlink"), ref=a1, location=loc, display=disp)
        dim = root.find("m:dimension", N)
        if dim is not None:
            last = max(int(r.get("r")) for r in sd)
            ref = dim.get("ref")
            if ":" in ref:
                a, b = ref.split(":")
                col = "".join(ch for ch in b if ch.isalpha())
                dim.set("ref", f"{a}:{col}{last}")

    def get_cell_xml(self, sheet: str, a1: str) -> str | None:
        root = self._sheet_root(sheet)
        for c in root.iter(_q("c")):
            if c.get("r") == a1:
                return etree.tostring(c).decode()
        return None

    # ------------------------------------------------------------ theme
    def set_theme_colours(self, colours: dict[str, str]):
        """Replace theme colour slots (dk2, lt2, accent1..6, hlink, folHlink)
        with RGB hex values. Every cell that uses theme colours follows."""
        part = "xl/theme/theme1.xml"
        a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        root = etree.fromstring(self.parts[part])
        scheme = root.find(f".//{{{a}}}clrScheme")
        for slot, hexv in colours.items():
            el = scheme.find(f"{{{a}}}{slot}")
            if el is None:
                raise KeyError(f"theme slot not found: {slot}")
            for ch in list(el):
                el.remove(ch)
            etree.SubElement(el, f"{{{a}}}srgbClr", val=hexv.lstrip("#").upper())
        self.parts[part] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ------------------------------------------------------------- calc
    def remove_calc_chain(self):
        target = "xl/calcChain.xml"
        if target in self.parts:
            del self.parts[target]
        for rel in list(self.rels):
            if rel.get("Target", "").endswith("calcChain.xml"):
                self.rels.remove(rel)
        for o in list(self.ct):
            if o.get("PartName") == "/xl/calcChain.xml":
                self.ct.remove(o)

    def set_full_calc_on_load(self):
        cp = self.wb.find("m:calcPr", N)
        if cp is None:
            cp = etree.SubElement(self.wb, _q("calcPr"))
        cp.set("fullCalcOnLoad", "1")
        cp.set("calcMode", "auto")

    # ------------------------------------------------------------- save
    def save(self, out: str | Path):
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.parts["xl/workbook.xml"] = etree.tostring(self.wb, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.parts["xl/_rels/workbook.xml.rels"] = etree.tostring(self.rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.parts["[Content_Types].xml"] = etree.tostring(self.ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        for name, root in self._sheet_cache.items():
            self.parts[self.sheet_part(name)] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        order = [n for n in self.order if n in self.parts]
        order += [n for n in self.parts if n not in order]
        if "[Content_Types].xml" in order:
            order.remove("[Content_Types].xml")
            order.insert(0, "[Content_Types].xml")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for n in order:
                z.writestr(n, self.parts[n])
        return out
