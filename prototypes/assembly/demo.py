"""Build a small model, then insert modules into the built workbook and show what links itself in.

    python prototypes/assembly/demo.py [out_dir]

Writes base.xlsx, live.xlsx (base with the plan applied, as the add-in would)
and fresh.xlsx (built from scratch with the same modules) to out_dir, default
build/assembly. Workbooks stay out of git.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from assemble import Library, Model, assemble, open_model, plan_change, write_workbook  # noqa: E402


def base_model(lib: Library) -> Model:
    m = Model(lib, periods=12)
    m.insert("demo.statements")
    m.insert("demo.checks")
    m.insert("demo.revenue_line", base=100, growth=0.01)
    m.insert("demo.revenue_line", base=60, growth=0.03)
    m.insert("demo.cost_line", amount=90, inflation=0.002)
    m.insert("demo.debtors")
    m.insert("demo.facility", amount=1000, rate=0.06, instalment=50)
    return m


def insert_two(m: Model) -> None:
    m.insert("demo.revenue_line", base=150, growth=0.02)
    m.insert("demo.facility", amount=500, rate=0.08, instalment=25)


def main(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    lib = Library.load()
    base = base_model(lib)
    write_workbook(assemble(base), out / "base.xlsx", base)
    print(f"Built {out / 'base.xlsx'}")

    model, meta = open_model(out / "base.xlsx", lib)  # what the add-in does on open
    old = assemble(model)
    assert {s: [r.id for r in rows] for s, rows in old.sheets} == meta["rows"], "workbook drifted from its metadata"
    insert_two(model)
    new = assemble(model)
    plan = plan_change(old, new, dialect="uno")
    print("\nPreview: insert Revenue line 3 and Debt facility 2")
    for line in plan.preview:
        print("  " + line)
    print(f"\n{len(plan.ops)} operations: " + ", ".join(
        f"{k} x{sum(1 for o in plan.ops if o['op'] == k)}" for k in dict.fromkeys(o["op"] for o in plan.ops)))

    try:
        from live import apply_plan
    except ImportError:
        print("LibreOffice bridge not available; skipping the live apply")
        return
    apply_plan(out / "base.xlsx", plan, out / "live.xlsx", model, new)
    write_workbook(new, out / "fresh.xlsx", model)
    print(f"\nWrote {out / 'live.xlsx'} and {out / 'fresh.xlsx'}")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[2] / "build" / "assembly")
