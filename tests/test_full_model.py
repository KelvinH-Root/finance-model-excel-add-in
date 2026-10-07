"""The full financial model demo (library/hfg, recipe full_model.json): built by the engine,
recalculated in LibreOffice, it must match the Python reference in every module row the reference
works out, raise no errors and no alerts, and balance."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
RECIPE = ROOT / "library" / "hfg" / "recipes" / "full_model.json"
needs_lo = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice is not installed")
needs_node = pytest.mark.skipif(not shutil.which("node"), reason="Node is not installed")


# Sheets without the timeline block: content starts on row 5 (17 on timeline sheets).
PLAIN_SHEETS = {"Contents", "Lookups", "Scenarios", "Model", "Dashboards", "Appendices", "Income summary", "Balance summary",
                "Cash summary", "Budget summary", "Version comparison", "Scenario summary", "Versions", "Budget", "Reports", "Income report", "Balance report",
                "Cash report", "Budget report", "Scenario report"}


def read_metadata(path: Path) -> dict:
    import html
    import re
    with zipfile.ZipFile(path) as z:
        xml = z.read("customXml/item1.xml").decode()
    return json.loads(html.unescape(re.search(r">(\{.*\})<", xml, re.S).group(1)))


@pytest.fixture(scope="module")
def built(tmp_path_factory) -> Path:
    out = tmp_path_factory.mktemp("full") / "full_model.xlsx"
    subprocess.run(["node", "tools/build-recipe.ts", str(RECIPE), str(out)], cwd=ROOT / "engine", check=True, capture_output=True)
    return out


@pytest.fixture(scope="module")
def reference():
    from examples.full_model.recipe import FY_END, LAST_ACTUAL, START, T
    from examples.full_model.reference import Inst, run
    r = json.loads(RECIPE.read_text())
    insts, n = [], {}
    for i in r["instances"]:
        n[i["module"]] = n.get(i["module"], 0) + 1
        insts.append(Inst(i["module"], n[i["module"]], i.get("settings", {}), i.get("name"), i.get("data", {})))
    return run(insts, T, last_actual=LAST_ACTUAL, fy_end=FY_END, start_month=START[1], start_year=START[0])


def test_recipe_is_current():
    """The recipe in the library is what examples/full_model/recipe.py writes."""
    from examples.full_model.recipe import recipe
    assert json.loads(RECIPE.read_text()) == json.loads(json.dumps(recipe()))


def test_history_generator_balances():
    from examples.full_model.recipe import T, generate
    _, gen = generate()
    for t in range(T):
        assert abs(gen["fs/net_assets"][t] - gen["fs/equity"][t]) < 1e-6
        prev = gen["fs/cash"][t - 1] if t else 900_000
        assert abs(gen["fs/cash"][t] - prev - gen["fs/net_cash"][t]) < 1e-6


@needs_node
@needs_lo
def test_full_model_matches_the_reference(built, reference):
    from hfgmodels.verify import recalculate
    meta = read_metadata(built)
    sheets = list(meta["rows"])
    calc = recalculate(built, sheets)
    assert all(not e for e in calc.errors.values()), {s: e[:5] for s, e in calc.errors.items() if e}
    first = {s: (5 if s in PLAIN_SHEETS else 17) for s in sheets}
    pos = {rid: (s, first[s] + k) for s, ids in meta["rows"].items() for k, rid in enumerate(ids)}
    T = meta["model"]["periods"]
    compared = 0
    misses = []
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
    assert compared > 150
    assert not misses, misses[:10]
    errors = calc.get("Checks", pos["fm.checks#1/errors"][1], 9)
    alerts = calc.get("Checks", pos["fm.checks#1/alerts"][1], 9)
    assert (errors, alerts) == (0, 0)


@pytest.fixture(scope="module")
def calc(built):
    from hfgmodels.verify import recalculate
    meta = read_metadata(built)
    return meta, recalculate(built, list(meta["rows"]))


def row_values(meta, calc, rid, n, col=10):
    """A layout row's values from column J (or another column), by its row id."""
    for sheet, ids in meta["rows"].items():
        if rid in ids:
            first = 5 if sheet in PLAIN_SHEETS else 17
            return [calc.get(sheet, first + ids.index(rid), col + k) for k in range(n)]
    raise KeyError(rid)


@needs_node
@needs_lo
def test_summaries_read_the_statements(calc, reference):
    """Report tables are the statement lines for the window the selections choose."""
    meta, c = calc
    near = lambda a, b: all(abs((x or 0) - y) < 0.01 for x, y in zip(a, b))
    npat = reference["fs/npat"]
    rows = [row_values(meta, c, f"fm.income_summary#1/r/C01/{k}", 12) for k in (1, 2, 3)]
    assert near(rows[0], npat[0:12]) and near(rows[1], npat[12:24]) and near(rows[2], npat[24:36])
    # revenue by line for FY2027, in model order
    revs = sorted(k for k in reference if k.startswith("fm.revenue#") and k.endswith("/revenue"))
    for k, key in enumerate(revs, start=1):
        assert near(row_values(meta, c, f"fm.income_summary#1/r/C02/{k}", 12), reference[key][12:24])
    # the cash bridge runs from the cash at March 2026 to the cash at March 2027
    bridge = row_values(meta, c, "fm.cash_summary#1/r/C12/1", 7)
    assert abs(bridge[0] - reference["fs/cash"][11]) < 0.01 and abs(bridge[6] - reference["fs/cash"][23]) < 0.01
    assert abs(sum(bridge[1:6]) - (reference["fs/cash"][23] - reference["fs/cash"][11])) < 0.01
    # net assets built up at March 2027 (the month shown on the Balance summary)
    na = row_values(meta, c, "fm.balance_summary#1/r/C10/1", 5)
    assert abs(na[4] - reference["fs/net_assets"][23]) < 0.01


@needs_node
@needs_lo
def test_budget_summary_reads_the_approved_budget(calc):
    meta, c = calc
    r = json.loads(RECIPE.read_text())
    versions = next(i for i in r["instances"] if i["module"] == "fm.versions")["data"]["versions"]
    approved = next(v for v in versions if v["type"] == "Budget" and v["status"] == "Approved" and v["year"] == "FY2027")
    got = row_values(meta, c, "fm.budget_summary#1/r/C17/3", 12)
    assert all(abs(g - e) < 0.01 for g, e in zip(got, approved["values"]["rev"][12:24]))
    # every saved version still matches its checksum
    reg = [row_values(meta, c, f"fm.versions#1/reg/{v['id']}", 2, col=21) for v in versions]
    assert all(abs(a - b) < 0.01 for a, b in reg)


@needs_node
@needs_lo
def test_scenario_results_match_each_scenario(calc):
    """The data table on the Scenarios sheet works the model out under each scenario, as the reference does."""
    from examples.full_model.recipe import FY_END, LAST_ACTUAL, START, T
    from examples.full_model.reference import Inst, run
    meta, c = calc
    r = json.loads(RECIPE.read_text())
    insts, n = [], {}
    for i in r["instances"]:
        n[i["module"]] = n.get(i["module"], 0) + 1
        insts.append(Inst(i["module"], n[i["module"]], i.get("settings", {}), i.get("name"), i.get("data", {})))
    for s in (1, 2, 3):
        ref = run(insts, T, last_actual=LAST_ACTUAL, fy_end=FY_END, start_month=START[1], start_year=START[0], scenario=s)
        for y in range(3):
            got = row_values(meta, c, f"scnres/fm.scenario_summary#1/S01/{y + 1}", 3)[s - 1]
            assert abs(got - sum(ref["fs/revenue"][12 * y:12 * y + 12])) < 0.01
            got = row_values(meta, c, f"scnres/fm.scenario_summary#1/S03/{y + 1}", 3)[s - 1]
            assert abs(got - ref["fs/cash"][12 * y + 11]) < 0.01
        for j in range(12):
            got = row_values(meta, c, f"scnres/fm.scenario_summary#1/S06/{j + 1}", 3)[s - 1]
            assert abs(got - ref["fs/npat"][12 + j]) < 0.01
