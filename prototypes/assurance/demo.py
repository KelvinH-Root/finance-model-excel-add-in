"""Take the assembly demo through the assurance features and print what the add-in would show.

    python prototypes/assurance/demo.py [out_dir]

Writes to out_dir (default build/assurance; workbooks stay out of git):
  a0.xlsx           the demo with Model assurance: Group assumptions, Input register, key outputs
  a1.xlsx ... a5    after each command, every one checked before and after and logged
  fault.xlsx        a command with a fault planted in its plan, which the check catches
  auditor.xlsx, lender.xlsx, board.xlsx   release copies
  gst.xlsx          GST return periods, due dates and cash timing (two-monthly, odd months)
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
for p in (ROOT, HERE.parent / "assembly", HERE):
    sys.path.insert(0, str(p))

import compare as CMP  # noqa: E402
import gst  # noqa: E402
import guard  # noqa: E402
import register as REG  # noqa: E402
import release as REL  # noqa: E402
from assemble import Library, assemble, write_workbook  # noqa: E402


def main(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    lib = Library.load()
    m = REG.assured_model(lib)
    write_workbook(assemble(m), out / "a0.xlsx", m)
    print(f"Built {out / 'a0.xlsx'}")
    steps = [
        ("a0", "a1", "Insert Income summary", lambda x: x.insert("demo.dashboard", first=1), "none"),
        ("a1", "a2", "Insert Revenue line 3", lambda x: x.insert("demo.revenue_line", base=150, growth=0.02), "reach"),
        ("a2", "a3", "Open (version 4 of the group assumptions published)", lambda x: REG.note_latest(x, 4), "none"),
        ("a3", "a4", "Update to latest group assumptions", lambda x: REG.use_set(x, 4), "reach"),
        ("a4", "a5", "Insert Debt facility 2", lambda x: x.insert("demo.facility", amount=500, rate=0.08, instalment=25), "reach"),
    ]
    for src, dst, label, change, effect in steps:
        e = guard.run(out / f"{src}.xlsx", out / f"{dst}.xlsx", lib, label, change, effect=effect, when="7 October 2026")
        print(f"\n{label}")
        for line in guard.result_card(e):
            print("  " + line)
    e = guard.run(out / "a2.xlsx", out / "fault.xlsx", lib, "Rename Revenue line 3 (with a planted fault)", lambda x: None,
                  effect="none", sabotage=guard.facility_repayments_stop, when="7 October 2026")
    print("\nRename Revenue line 3, with a fault planted in the plan")
    for line in guard.result_card(e):
        print("  " + line)

    print("\nModel compare: a1 against a5")
    for line in CMP.summary(CMP.compare(out / "a1.xlsx", out / "a5.xlsx", lib)):
        print("  " + line)

    for profile in REL.PROFILES:
        res = REL.release(out / "a5.xlsx", out / f"{profile}.xlsx", lib, profile, when="7 October 2026")
        print(f"\n{REL.PROFILES[profile]['title']}: " + ", ".join(res["sheets"]))

    start, n = date(2026, 4, 30), 24
    net = [((-1) ** t) * 100.0 + (t % 5) * 37.5 - (400.0 if t in (3, 4, 11) else 0.0) for t in range(n)]
    gst.write_workbook(out / "gst.xlsx", start, n, 2, 1, 1, net)   # two-monthly, periods ending in odd months
    print(f"\nWrote {out / 'gst.xlsx'}")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "build" / "assurance")
