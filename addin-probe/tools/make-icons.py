"""Draw the probe's command icons (16, 32 and 80 px) from the glyphs in src/commands.js.

    python tools/make-icons.py

Placeholder icons for the probe only: one glyph per command, dark grey on the main tab and
HFG blue on the Build tab. The real icon set is a Phase 1 design job.
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
COLOURS = {"main": (64, 64, 64, 255), "build": (21, 96, 130, 255), "context": (64, 64, 64, 255)}

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
