"""Screenshots of an example model: chosen sheet areas as PNGs, and one captioned PDF.

    python prototypes/models/screenshots.py budget|development [out_dir]

Each view is recalculated in LibreOffice, frozen to values and exported with row and column
headers, the way a spreadsheet shows it, then trimmed. LibreOffice stands in for Excel here,
so fonts and chart details differ slightly from what Excel draws. Needs LibreOffice,
pdftoppm and ImageMagick; the development example also needs HFG_TEMPLATE.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))

VIEWS = {
    "budget": ("Budget and actuals example", [
        ("Contents", "A1:G82", "Contents: numbered sections, lettered sheets and the headings on each sheet, every entry a link, "
                               "with the error and alert totals at the end. Rebuilt after every structural change."),
        ("Dashboards", "A1:G20", "Section cover: separates the Dashboards section, with links to the contents and to the sheets "
                                 "either side. A cover is created with its section's first sheet."),
        ("Income summary", "A1:Y39", "Income summary (Phase 1): the year shown drop-down and the module's five charts."),
        ("Budget summary", "A1:Y39", "Budget summary (Phase 1): actual (solid) and forecast (hatched) against the line chosen in "
                                     "Compared with: the approved budget, last month's reforecast, the latest reforecast, the budget "
                                     "being built, or any saved version."),
        ("Version comparison", "A1:Y58", "Version comparison (Phase 1): the month, year to date and full year against two saved versions "
                                         "(here the approved budget and last month's reforecast), then the full-year outturn by reforecast, "
                                         "both comparisons by month, what each version expected for the month, and the walk to the outturn."),
        ("Versions", "A1:Q32", "Versions: the register of saved budgets and monthly reforecasts, written by Save version. Each is a "
                               "full copy of the statements as values in the version store, with a checksum to show it has not changed."),
        ("Seasonality", "A1:U20", "Seasonality (Phase 1): each month's share of prior years' revenue, a year can be left out, "
                                  "a typed override, and the annual budget phased by the profile."),
        ("Inputs", "A1:V44", "Inputs: monthly actuals then the base forecast, by category."),
        ("Statements", "A1:V64", "Statements: income statement, balance sheet and cash flow for the active scenario, then the "
                                 "budget and each scenario."),
        ("Income report", "A1:Y54", "Income report (Phase 2): the first nine of its 28 charts."),
        ("Checks", "A1:H50", "Checks: every error check and alert with totals. A2 on every sheet links here and shows a tick "
                             "when the error checks are clear."),
    ]),
    "development": ("Development example", [
        ("Contents", "A1:G60", "Contents: the development sheets with each site's block listed under its sheet."),
        ("Model", "A1:G20", "Section cover for the Development Model section."),
        ("Development summary", "A1:X56", "Development summary: key figures for the portfolio and four native charts, every number "
                                          "a formula on the development sheets."),
        ("Dev_Sites", "A1:L62", "Sites: model-wide assumptions, then one input block per site (general, opening position, cost "
                                "budget by line, sale and exit, funding)."),
        ("Dev_Costs", "A1:V60", "Costs: phasing share by line and forecast cost, portfolio first, then each site."),
        ("Dev_Funding", "A1:V68", "Funding: debt ceiling, finance costs, draws and equity, covenant ratios and cash flows for IRR."),
        ("Dev_FS", "A1:V70", "Financial statements: WIP, income statement, balance sheet and cash flow, with their checks."),
        ("Dev_Returns", "A1:L41", "Returns by site and portfolio."),
        ("Dev_Checks", "A1:L31", "Checks: every check in plain words, by site, with error and alert totals."),
    ]),
}


def render(workbook: Path, views: list[tuple[str, str, str]], out: Path) -> list[Path]:
    from hfgmodels.verify import _apply_data_tables, _prop, data_tables, libreoffice
    out.mkdir(parents=True, exist_ok=True)
    url = lambda p: "file://" + str(Path(p).resolve())        # noqa: E731
    pdf = Path(tempfile.mkdtemp()) / "views.pdf"
    tables = data_tables(workbook)
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(url(workbook), "_blank", 0, (_prop("Hidden", True),))
        try:
            if tables:   # LibreOffice reads only part of an Excel data table; rebuild it as Excel works it out
                _apply_data_tables(doc, tables)
            doc.calculateAll()
            for n in list(doc.Sheets.ElementNames):
                sh = doc.Sheets.getByName(n)
                cur = sh.createCursor()
                cur.gotoEndOfUsedArea(False)
                rng = sh.getCellRangeByPosition(0, 0, cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow)
                rng.setDataArray(rng.getDataArray())
            keep = [v[0] for v in views]
            for n in list(doc.Sheets.ElementNames):
                if n not in keep:
                    doc.Sheets.removeByName(n)
            for i, (name, area, _) in enumerate(views):
                doc.Sheets.moveByName(name, i)
                sh = doc.Sheets.getByName(name)
                sh.setPrintAreas((sh.getCellRangeByName(area).getRangeAddress(),))
                style = doc.StyleFamilies.getByName("PageStyles").getByName(sh.PageStyle)
                style.PrintHeaders = True
                style.HeaderIsOn = style.FooterIsOn = False
                style.TopMargin = style.BottomMargin = style.LeftMargin = style.RightMargin = 300
                style.IsLandscape = True
                style.Width, style.Height = 42000, 29700
                style.ScaleToPagesX = style.ScaleToPagesY = 1
            doc.storeToURL(url(pdf), (_prop("FilterName", "calc_pdf_Export"),))
        finally:
            doc.close(True)
    pngs = []
    for i, (name, _, _) in enumerate(views, start=1):
        base = out / f"{i:02d}_{name.replace(' ', '_')}"
        subprocess.run(["pdftoppm", "-r", "110", "-png", "-singlefile", "-f", str(i), "-l", str(i), str(pdf), str(base)],
                       check=True, stderr=subprocess.DEVNULL)
        png = base.with_suffix(".png")
        subprocess.run(["convert", str(png), "-trim", "-bordercolor", "white", "-border", "16", str(png)], check=True)
        pngs.append(png)
    return pngs


def gallery(title: str, views, pngs: list[Path], pdf: Path) -> Path:
    """One page per view: the caption above the screenshot."""
    from PIL import Image, ImageDraw, ImageFont
    bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 26)
    body = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
    pages = []
    for k, ((name, _, caption), png) in enumerate(zip(views, pngs), start=1):
        shot = Image.open(png).convert("RGB")
        width = max(shot.width, 1400)
        words, lines, line = caption.split(), [], ""
        for w in words:
            trial = (line + " " + w).strip()
            if body.getlength(trial) > width - 80:
                lines.append(line)
                line = w
            else:
                line = trial
        lines.append(line)
        head = 40 + 34 + 28 * len(lines) + 24
        page = Image.new("RGB", (width, head + shot.height), "white")
        d = ImageDraw.Draw(page)
        d.text((40, 30), f"{title}  |  {k}. {name}", font=bold, fill=(64, 64, 64))
        for j, text in enumerate(lines):
            d.text((40, 74 + 28 * j), text, font=body, fill=(64, 64, 64))
        d.line((40, head - 12, width - 40, head - 12), fill=(103, 157, 181), width=3)
        page.paste(shot, (0, head))
        pages.append(page)
    pages[0].save(pdf, save_all=True, append_images=pages[1:], resolution=110)
    return pdf


def main(which: str, out: Path) -> None:
    title, views = VIEWS[which]
    work = Path(tempfile.mkdtemp())
    if which == "budget":
        sys.path.insert(0, str(ROOT / "prototypes" / "reports"))
        import build
        book = build.build(work / "budget_example.xlsx")
    else:
        import development
        book = development.build(work / "development_example.xlsx")
    pngs = render(book, views, out / which)
    print(gallery(title, views, pngs, out / f"{which}_screenshots.pdf"))


if __name__ == "__main__":
    sys.path.insert(0, str(HERE))
    main(sys.argv[1], Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "build" / "models" / "screenshots")
