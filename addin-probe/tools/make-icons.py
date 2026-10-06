"""Draw the probe's command icons (16, 32 and 80 px) from the glyphs in src/commands.js.

    python tools/make-icons.py

Placeholder icons for the probe only: one glyph per command in the HF palette (steel blue on the
main tab, navy on the Build tab, slate on the right-click menu). The real icon set is a Phase 1 design job.
"""
import json
import subprocess
from pathlib import Path

from fontTools.ttLib import TTFont
from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent.parent
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
OUT = HERE / "assets" / "cmd"
# HF palette: steel blue (#679DB5) on the HFG Model tab, navy (#09122C) on the HFG Build tab, and
# slate (#3C4E60) on the right-click menu, whose 16 px icons need the extra contrast.
COLOURS = {"main": (0x67, 0x9D, 0xB5, 255), "build": (0x09, 0x12, 0x2C, 255), "context": (0x3C, 0x4E, 0x60, 255)}

cmds = json.loads(subprocess.check_output(
    ["node", "--input-type=module", "-e",
     "await import('./src/commands.js');console.log(JSON.stringify(globalThis.HfgCommands.all().filter(x=>!x.parent)))"],
    cwd=HERE))
cmap = TTFont(FONT).getBestCmap()
OUT.mkdir(parents=True, exist_ok=True)
for c in cmds:
    g = c["glyph"]
    assert all(ord(ch) in cmap for ch in g), f"{c['key']}: glyph {g!r} is not in DejaVu Sans"
    for size in (16, 32, 80):
        scale = 8
        big = size * scale
        img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        font = ImageFont.truetype(BOLD if g.isalnum() or g == "?" else FONT, int(big * 0.9))
        box = d.textbbox((0, 0), g, font=font)
        w, h = box[2] - box[0], box[3] - box[1]
        fit = 0.86 * big / max(w, h)
        font = ImageFont.truetype(BOLD if g.isalnum() or g == "?" else FONT, int(big * 0.9 * fit))
        box = d.textbbox((0, 0), g, font=font)
        w, h = box[2] - box[0], box[3] - box[1]
        d.text(((big - w) / 2 - box[0], (big - h) / 2 - box[1]), g, font=font, fill=COLOURS[c["tab"]])
        img.resize((size, size), Image.LANCZOS).save(OUT / f"{c['key']}-{size}.png", optimize=True)
print(f"{len(cmds)} icons x 3 sizes written to {OUT}")

# The add-in's own icon (manifest and ribbon tab): a 3 by 3 grid in white on HF navy, with the
# last cell in steel blue.
NAVY, STEEL, WHITE = (0x09, 0x12, 0x2C, 255), (0x67, 0x9D, 0xB5, 255), (255, 255, 255, 255)
for size in (16, 32, 64, 80):
    scale = 8
    big = size * scale
    img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((0, 0, big - 1, big - 1), radius=int(big * 0.08), fill=NAVY)
    lo, hi = big * 0.18, big * 0.82
    d.rectangle((lo, lo, hi, hi), fill=WHITE)
    line = (hi - lo) * 0.08
    cell = (hi - lo - 4 * line) / 3
    for r in range(3):
        for c in range(3):
            x0 = lo + line + c * (cell + line)
            y0 = lo + line + r * (cell + line)
            d.rectangle((x0, y0, x0 + cell, y0 + cell), fill=STEEL if (r, c) == (2, 2) else NAVY)
    img.resize((size, size), Image.LANCZOS).save(HERE / "assets" / f"icon-{size}.png", optimize=True)
print("add-in icon written in 4 sizes")
