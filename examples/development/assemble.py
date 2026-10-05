"""Assemble the development model into a copy of the template."""

from __future__ import annotations

import os
from pathlib import Path

from hfgmodels import frame
from hfgmodels.xlsx.package import TemplatePackage
from hfgmodels.xlsx.styles import StyleBook
from .build import build_model
from .recipe import Recipe, load


def template_path(recipe: Recipe, override: str | None = None) -> Path:
    p = override or os.environ.get(recipe.template_env)
    if not p:
        raise SystemExit(f"Set {recipe.template_env} to the path of Budget_Template.xlsx (it is not kept in git).")
    path = Path(p).expanduser()
    if not path.exists():
        raise SystemExit(f"Template not found: {path}")
    return path


COVER_STYLES = {"number": 219, "section": 894, "letter": 22, "item": 22, "pad": 23}


def add_cover_contents(pkg: TemplatePackage, model, start_row: int = 55, number: int = 6):
    """Add a 'Development Model' section to the Cover sheet's table of contents,
    linking to every new sheet, in the template's own TOC styles."""
    s = COVER_STYLES
    r = start_row
    cells = [(f"B{r}", number, s["number"]), (f"C{r}", "Development Model", s["section"])]
    cells += [(f"{c}{r}", None, s["section"]) for c in "DEFGHIJKL"]
    links = [(f"C{r}", f"{model.sheets[0].name}!A1", "Development Model")]
    for i, sl in enumerate(model.sheets):
        rr = r + 1 + i
        letter = chr(ord("a") + i) + "."
        cells += [(f"C{rr}", letter, s["letter"]), (f"D{rr}", sl.title, s["item"]),
                  (f"E{rr}", None, s["pad"]), (f"F{rr}", None, s["pad"]), (f"G{rr}", None, s["pad"])]
        links += [(f"C{rr}", f"{sl.name}!A1", letter), (f"D{rr}", f"{sl.name}!A1", sl.title)]
    pkg.append_cells("Cover", cells, merges=[f"C{r}:L{r}"], links=links, row_height=12)


def assemble(recipe_path: str | Path, out: str | Path, template: str | None = None, values: dict | None = None):
    """Build the workbook. `values` (sheet -> {(row, col): value}) adds cached
    results to formula cells, normally taken from a LibreOffice recalculation."""
    r = load(recipe_path)
    pkg = TemplatePackage(template_path(r, template))
    st = StyleBook(pkg.parts["xl/styles.xml"], frame.TEMPLATE_STYLES, frame.DERIVED_STYLES)
    model = build_model(r)
    sheets = model.render(st)
    pkg.parts["xl/styles.xml"] = st.to_bytes()
    if r.theme:
        pkg.set_theme_colours(r.theme)
    # timeline settings on the template's Time sheet
    pkg.set_cell_value("Time", "H13", r.periods)
    pkg.set_cell_value("Time", "H26", r.last_actual)
    for sl in model.sheets:
        xml = sheets[sl.name].to_xml((values or {}).get(sl.name))
        pkg.add_sheet(sl.name, xml, after=sl.after)
    add_cover_contents(pkg, model)
    existing = pkg.defined_names()
    for name, addr in model.names.items():
        if name in existing:
            raise SystemExit(f"defined name clash with the template: {name}")
        pkg.set_defined_name(name, addr)
    pkg.remove_calc_chain()
    pkg.set_full_calc_on_load()
    pkg.set_active_sheet("Dev_Sites")
    pkg.save(out)
    return r, model, sheets


def build(recipe_path: str | Path, out: str | Path, template: str | None = None, cache_values: bool = True):
    """Build the workbook; when cache_values is set, recalculate once in
    LibreOffice and write the results into the new sheets' formula cells so
    previews and non-calculating readers show numbers. Excel still
    recalculates everything on open."""
    r, model, sheets = assemble(recipe_path, out, template)
    if not cache_values:
        return r, model, sheets, None
    from hfgmodels.verify import recalculate

    names = [s.name for s in model.sheets]
    cw = recalculate(out, names)
    values = {}
    for name, ws in sheets.items():
        grid = cw.values[name]
        v = {}
        for (row, col), cell in ws.cells.items():
            if cell.formula is None:
                continue
            if row - 1 < len(grid) and col - 1 < len(grid[row - 1]):
                v[(row, col)] = grid[row - 1][col - 1]
        values[name] = v
    r, model, sheets = assemble(recipe_path, out, template, values=values)
    return r, model, sheets, cw
