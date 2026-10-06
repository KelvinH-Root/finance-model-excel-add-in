"""Golden fixtures for the TypeScript engine, written by the Python proof it replaces.

    python engine/tools/golden.py

Each scenario is a model before and after a change (as model dicts, so the TypeScript side
needs no copy of the scenario code), the layout and every rendered cell after the change in
both dialects, and the change plan between the two. The TypeScript tests must reproduce them
exactly. Fictional demo data only.
"""
from __future__ import annotations

import json
import sys
from dataclasses import asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "prototypes" / "assembly"))
sys.path.insert(0, str(ROOT / "prototypes" / "assurance"))

from assemble import (FIRST_ROW, Library, Model, assemble, chart_refs, frame_cells,  # noqa: E402
                      plan_change, row_cells)
import register as REG  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "test" / "fixtures"


def base_model(lib):
    m = Model(lib, periods=12)
    m.insert("demo.statements")
    m.insert("demo.checks")
    m.insert("demo.revenue_line", base=100, growth=0.01)
    m.insert("demo.revenue_line", base=60, growth=0.03)
    m.insert("demo.cost_line", amount=90, inflation=0.002)
    m.insert("demo.debtors")
    m.insert("demo.facility", amount=1000, rate=0.06, instalment=50)
    return m


def layout_dict(lay):
    return {
        "periods": lay.periods,
        "sheets": [[s, [{k: v for k, v in asdict(r).items()} for r in rows]] for s, rows in lay.sheets],
        "names": [[k, v] for k, v in lay.names.items()],
        "records": lay.records, "warnings": lay.warnings,
        "blocks": [[k, v] for k, v in lay.blocks.items()],
        "charts": [asdict(c) for c in lay.charts],
        "kinds": lay.kinds, "titles": lay.titles, "name_cols": lay.name_cols, "headlines": lay.headlines,
    }


def cells_dict(lay, dialect):
    pos = lay.positions()
    out = {}
    for s, rows in lay.sheets:
        sheet = {}
        for (r, c), v in frame_cells(lay, s, dialect).items():
            sheet.setdefault(str(r), {})[str(c)] = v
        for k, row in enumerate(rows):
            rn = FIRST_ROW + k
            for c, v in row_cells(lay, s, row, rn, pos, dialect).items():
                sheet.setdefault(str(rn), {})[str(c)] = v
        out[s] = sheet
    out["@charts"] = [chart_refs(lay, c, dialect) for c in lay.charts]
    return out


def plan_dict(old, new, dialect):
    p = plan_change(old, new, dialect)
    ops = []
    for o in p.ops:
        o = dict(o)
        if "cells" in o:
            o["cells"] = {str(k): v for k, v in o["cells"].items()}
        ops.append(o)
    return {"ops": ops, "preview": p.preview}


def scenario(name, old, change):
    new = old.copy()
    change(new)
    lo, ln = assemble(old), assemble(new)
    d = {"name": name, "old": old.to_dict(), "new": new.to_dict(), "layout": layout_dict(ln),
         "cells": {"excel": cells_dict(ln, "excel"), "uno": cells_dict(ln, "uno")},
         "plan": {"excel": plan_dict(lo, ln, "excel"), "uno": plan_dict(lo, ln, "uno")}}
    (OUT / f"{name}.json").write_text(json.dumps(d, ensure_ascii=False, separators=(",", ":")))
    return new


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    lib = Library.load()
    base = base_model(lib)
    empty = Model(lib, periods=12)
    scenario("01_build_base", empty, lambda m: [m.instances.extend(base.instances), m.counters.update(base.counters)])
    m2 = scenario("02_insert_revenue_and_facility", base,
                  lambda m: (m.insert("demo.revenue_line", base=150, growth=0.02), m.insert("demo.facility", amount=500, rate=0.08, instalment=25)))
    m3 = scenario("03_insert_dashboard", m2, lambda m: m.insert("demo.dashboard", first=1))
    m4 = scenario("04_revenue_joins_chart", m3, lambda m: m.insert("demo.revenue_line", base=80, growth=0.04))
    scenario("05_remove_revenue_line", m4, lambda m: m.remove("demo.revenue_line#2"))
    scenario("06_change_setting", m4, lambda m: m.instance("demo.revenue_line#1").settings.update(base=120))
    a0 = REG.assured_model(lib)
    scenario("07_build_assured", Model(lib, periods=a0.periods), lambda m: [m.instances.extend(a0.instances), m.counters.update(a0.counters), m.assurance.update(a0.assurance)])
    a1 = scenario("08_note_latest_set", a0, lambda m: REG.note_latest(m, 4))
    a2 = scenario("09_use_latest_set", a1, lambda m: REG.use_set(m, 4))
    a3 = scenario("10_unbind_rate", a2, lambda m: m.unbind("demo.facility#1", "rate", 0.065))
    scenario("11_insert_into_assured", a3, lambda m: m.insert("demo.revenue_line", base=150, growth=0.02))
    print(f"wrote {len(list(OUT.glob('*.json')))} fixtures to {OUT}")


if __name__ == "__main__":
    main()
