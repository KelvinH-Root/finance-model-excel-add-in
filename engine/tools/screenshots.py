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

if __name__ == "__main__":
    book = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else book.parent / "shots" / book.stem
    pngs = render(book, VIEWS, out)
    print(gallery(f"Engine build: {book.name}", VIEWS, pngs, out.with_suffix(".pdf")))
