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
    first = {s: (17 if s not in ("Contents", "Lookups", "Scenarios", "Model", "Dashboards", "Appendices") else 5) for s in sheets}
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
