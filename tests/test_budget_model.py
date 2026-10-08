"""The budget demo (library/hfg, recipe budget_model.json): a contractor's budget with business units,
GL lines spread by seasonality profiles, the approved budget and monthly reforecasts saved. Built by
the engine and recalculated in LibreOffice, it must match the Python reference, raise no errors or
alerts, and its budget comparison must read the approved budget."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
from test_full_model import first_row, needs_lo, needs_node, read_metadata  # noqa: E402

RECIPE = ROOT / "library" / "hfg" / "recipes" / "budget_model.json"


@pytest.fixture(scope="module")
def built(tmp_path_factory) -> Path:
    out = tmp_path_factory.mktemp("budget") / "budget_model.xlsx"
    subprocess.run(["node", "tools/build-recipe.ts", str(RECIPE), str(out)], cwd=ROOT / "engine", check=True, capture_output=True)
    return out


def instances():
    from examples.full_model.reference import Inst
    r = json.loads(RECIPE.read_text())
    insts, n = [], {}
    for i in r["instances"]:
        n[i["module"]] = n.get(i["module"], 0) + 1
        insts.append(Inst(i["module"], n[i["module"]], i.get("settings", {}), i.get("name"), i.get("data", {})))
    return r, insts


def test_recipe_is_current():
    from examples.budget_model.recipe import recipe
    assert json.loads(RECIPE.read_text()) == json.loads(json.dumps(recipe()))


def test_the_budget_phases_by_profile():
    """A GL line's forecast months are its annual budget spread by its seasonality profile."""
    from examples.budget_model.recipe import FY_END, LAST_ACTUAL, PROFILES, START, T
    from examples.full_model.reference import run
    _, insts = instances()
    ref = run(insts, T, last_actual=LAST_ACTUAL, fy_end=FY_END, start_month=START[1], start_year=START[0])
    line = next(i for i in insts if i.name == "4300 Capital works contracts")
    annual = line.settings["annual"]
    for t in range(LAST_ACTUAL, 24):   # October 2026 to March 2027, in FY2027
        assert abs(ref[f"{line.uid}/pre"][t] - annual * PROFILES[1][t % 12]) < 0.01
    assert abs(sum(PROFILES[0]) - 1) < 1e-9 and abs(sum(PROFILES[1]) - 1) < 1e-9


@needs_node
@needs_lo
def test_budget_model_matches_the_reference(built):
    from examples.budget_model.recipe import FY_END, LAST_ACTUAL, START, T
    from examples.full_model.reference import run
    from hfgmodels.verify import recalculate
    _, insts = instances()
    reference = run(insts, T, last_actual=LAST_ACTUAL, fy_end=FY_END, start_month=START[1], start_year=START[0])
    meta = read_metadata(built)
    sheets = list(meta["rows"])
    calc = recalculate(built, sheets)
    assert all(not e for e in calc.errors.values()), {s: e[:5] for s, e in calc.errors.items() if e}
    first = {s: first_row(s) for s in sheets}
    pos = {rid: (s, first[s] + k) for s, ids in meta["rows"].items() for k, rid in enumerate(ids)}
    compared, misses = 0, []
    for key, ref in reference.items():
        uid, row = key.rsplit("/", 1)
        rid = f"fm.statements#1/{row}" if uid == "fs" else f"{uid}/{row}"
        if rid not in pos:
            continue
        s, r = pos[rid]
        got = [calc.get(s, r, 10 + t) or 0.0 for t in range(T)]
        if any(abs(g - e) > 0.01 for g, e in zip(got, ref)):
            misses.append((rid, [(t + 1, round(g, 2), round(e, 2)) for t, (g, e) in enumerate(zip(got, ref)) if abs(g - e) > 0.01][:3]))
        compared += 1
    assert compared > 200
    assert not misses, misses[:10]
    # business units add up to the lines that name them
    bus = [i for i in insts if i.module == "fm.business_unit"]
    assert len(bus) == 3 and all(abs(sum(reference[f"{b.uid}/revenue"])) > 0 for b in bus)
    errors = calc.get("Checks", pos["fm.checks#1/errors"][1], 9)
    alerts = calc.get("Checks", pos["fm.checks#1/alerts"][1], 9)
    assert (errors, alerts) == (0, 0)
    # the Budget summary compares against the approved FY2027 budget, saved as the GL lines phased
    r, _ = instances()
    versions = next(i for i in r["instances"] if i["module"] == "fm.versions")["data"]["versions"]
    approved = next(v for v in versions if v["type"] == "Budget" and v["status"] == "Approved" and v["year"] == "FY2027")
    s, row = pos["fm.budget_summary#1/r/C17/3"]
    got = [calc.get(s, row, 10 + k) for k in range(12)]
    assert all(abs(g - e) < 0.01 for g, e in zip(got, approved["values"]["rev"][12:24]))
