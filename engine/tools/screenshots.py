"""Screenshots of a workbook the engine built, for checking the look by eye.

    python engine/tools/screenshots.py engine/build/demo_HF.xlsx [out_dir]

Uses the model screenshot helper (LibreOffice, pdftoppm, ImageMagick). LibreOffice stands in for
Excel and substitutes a font for Segoe UI, so spacing and glyphs differ slightly from Excel.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "prototypes" / "models"))

from screenshots import gallery, render  # noqa: E402

VIEWS = [
    ("Contents", "A1:K40", "Contents: the entity, the model name line with its status phrases and the Prepared by line, "
                           "the logo, notes, then the table of contents grouped by section and sheet. Every entry is a cell "
                           "hyperlink to a Go_ name with a screen tip."),
    ("Model", "A1:C14", "Section cover: section number, links to the contents and the sheets either side, section notes."),
    ("Revenue", "A1:V27", "A calculation sheet: header rows, the timeline block (rows 7 to 15 grouped and collapsed), "
                          "a section bar per module, inputs shaded in the entity's accent with a border all round."),
    ("Statements", "A1:V62", "Statements: section sub-headings, a dashed rule under each list's last item, bold totals "
                             "with a rule above, spacer rows of 6 and 9 points, checks that turn bold red when not clear."),
    ("Dashboard", "A1:AG32", "A summary module with its chart, coloured by theme slot."),
    ("Settings", "A1:V62", "Settings: the model, timeline, actuals, budget and display choices New model collects, with "
                           "drop-downs and check boxes over their linked cells, and the timeline block every sheet reads."),
    ("Lookups", "A1:F60", "Lookups: the lists behind the drop-downs, each under a List_ name."),
    ("Checks", "A1:V32", "Checks: error checks and alerts with their totals."),
]

# The full financial model (library/hfg, recipe full_model.json): one view per sheet, wide enough to
# show the last months of actuals (to period 18, column AA) and the first forecast months.
FULL_VIEWS = [
    ("Contents", "A1:K70", "Contents: the sections and sheets of the full model, every entry a link."),
    ("Income summary", "A1:V56", "Income summary: the year shown drop-down and the module's five charts, every number a formula on the statements."),
    ("Balance summary", "A1:V56", "Balance summary: the month shown and six balance sheet charts (movement, mix, make-up, bridge)."),
    ("Budget summary", "A1:V62", "Budget summary: actual (solid) and forecast (hatched) against the comparison chosen in Compared with, here the approved Budget FY2027."),
    ("Scenario summary", "A1:V56", "Scenario summary: each scenario side by side, by year and by month of the year shown, read from the data table on the Scenarios sheet."),
    ("Cash summary", "A1:V56", "Cash summary: the cash bridge for the year shown, operating cash flows, working capital, investing and financing, and cash."),
    ("Scenarios", "A1:L50", "Scenarios: the active scenario drop-down, the three scenario names, and the adjustments "
                            "each scenario makes to the rows that take one (columns J to L)."),
    ("Historical IS", "A1:AD62", "Historical income statement: a typed line for every module row that declares one, "
                                 "in groups with totals. Actual months of the calculation sheets read these lines."),
    ("Historical BS", "A1:AD68", "Historical balance sheet: the typed opening balance in column I, then the actual months."),
    ("Revenue and expenses", "A1:AD90", "Revenue and expenses: revenue categories with GST treatment drop-downs, cost of "
                                        "sales, staff and operating expenses. Actual months read the history; forecast "
                                        "months use the drivers, with the active scenario's adjustment."),
    ("Working capital", "A1:AD100", "Working capital: debtor and creditor days, inventory and payroll payables."),
    ("Assets", "A1:AD104", "Assets: fixed and intangible assets with capital expenditure, depreciation and amortisation."),
    ("Capital", "A1:AD62", "Capital: the debt facility and equity, with interest and dividends."),
    ("Tax", "A1:AD34", "Tax: GST (15%, two-monthly returns) and income tax with New Zealand provisional and terminal dates."),
    ("Other items", "A1:AD32", "Other items: interest on cash and the other balance sheet lines."),
    ("Financials", "A1:AD105", "Financials: the income statement, balance sheet and cash flow, actual then forecast, "
                               "with the balance and cash checks."),
    ("Income report", "A1:V95", "Income report (first 12 of its 28 charts): revenue, cost of sales and margin against the years either side, make-up, ranking and cumulative."),
    ("Balance report", "A1:V75", "Balance report: the seven balance sheet charts for the year and month shown."),
    ("Cash report", "A1:V95", "Cash report (all 12 charts): bridge, flows, working capital, closing cash and where cash came from and went."),
    ("Budget report", "A1:V62", "Budget report: six income statement lines against the comparison chosen."),
    ("Scenario report", "A1:V95", "Scenario report (first 12 of its 20 charts): every scenario side by side, by year, then by month of the year shown."),
    ("Versions", "A1:Z36", "Saved versions: the register of budgets and monthly reforecasts, with the checksum taken when each was saved and the checks."),
    ("Version store", "A1:V60", "Version store: each saved version's lines as values, keyed by version and line."),
    ("Settings", "A1:V40", "Settings: the timeline, actuals and budget drop-downs over their linked cells."),
    ("Lookups", "A1:F70", "Lookups: the frame's lists and the modules' lists (GST treatment, scenarios)."),
    ("Checks", "A1:L45", "Checks: every error check and alert with totals."),
]

if __name__ == "__main__":
    import zipfile
    book = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else book.parent / "shots" / book.stem
    with zipfile.ZipFile(book) as z:
        full = 'name="Historical IS"' in z.read("xl/workbook.xml").decode()
    views = FULL_VIEWS if full else VIEWS
    pngs = render(book, views, out)
    print(gallery(f"Engine build: {book.name}", views, pngs, out.with_suffix(".pdf")))
