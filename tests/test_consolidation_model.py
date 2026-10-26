"""The group consolidation model type (library/hfg, recipe consolidation_model.json): built by the
engine from the fictional group in examples/consolidation, recalculated in LibreOffice, every
group's consolidated trial balance must match the reference (story.py re-records every event as
each group sees it and never reads the register) in every month, with NCI and the surplus
attributable to it worked out once at each node. Every check clear; the one alert is the on-charge
the story leaves accrued at a month end.

Equity is presented as the statements need it: the outside investors' capital and their share of
distributions sit in NCI, where the reference leaves them in capital and distributions. The test
holds the two to the same total equity.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
RECIPE = ROOT / "library" / "hfg" / "recipes" / "consolidation_model.json"
needs_lo = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice is not installed")
needs_node = pytest.mark.skipif(not shutil.which("node"), reason="Node is not installed")
B = "fm.consolidation#1"
PLAIN = {"Contents", "Settings", "Lookups", "Group", "Appendices", "Entities", "Intercompany"}
EQUITY = (3000, 3100, 3200, 3300, 3400)


def build(tmp: Path, group: int = 1) -> Path:
    r = json.loads(RECIPE.read_text())
    r["instances"][0]["settings"]["group"] = group
    src = tmp / "recipes" / f"cons_{group}.json"
    src.parent.mkdir(parents=True, exist_ok=True)
    src.write_text(json.dumps(r))
    # the recipe builder reads the library beside the recipe's folder
    lib = tmp / "library"
    if not lib.exists():
        shutil.copytree(ROOT / "library" / "hfg", lib, ignore=shutil.ignore_patterns("recipes"))
    (lib / "recipes").mkdir(exist_ok=True)
    shutil.copy(src, lib / "recipes" / src.name)
    out = tmp / f"cons_{group}.xlsx"
    subprocess.run(["node", "tools/build-recipe.ts", str(lib / "recipes" / src.name), str(out)], cwd=ROOT / "engine",
                   check=True, capture_output=True)
    return out


class Calc:
    def __init__(self, path: Path):
        from hfgmodels.verify import recalculate
        from test_full_model import read_metadata
        self.meta = read_metadata(path)
        sheets = list(self.meta["rows"])
        self.calc = recalculate(path, sheets)
        self.T = self.meta["model"]["periods"]
        self.pos = {rid: (s, (5 if s in PLAIN else 17) + k) for s, ids in self.meta["rows"].items() for k, rid in enumerate(ids)}

    def row(self, rid: str, col: int = 10, n: int | None = None) -> list[float]:
        s, r = self.pos[rid]
        return [self.calc.get(s, r, col + k) or 0.0 for k in range(self.T if n is None else n)]

    def value(self, rid: str, col: int = 9) -> float:
        s, r = self.pos[rid]
        return self.calc.get(s, r, col) or 0.0


@pytest.fixture(scope="module")
def ref():
    from examples.consolidation import story
    return story, story.reference(story.book())


@pytest.fixture(scope="module")
def shown(tmp_path_factory):
    return Calc(build(tmp_path_factory.mktemp("cons")))


def near(a, b, tol=1e-6):
    return all(abs(x - y) <= tol for x, y in zip(a, b, strict=True))


def test_recipe_is_current():
    """The recipe in the library is what examples/consolidation/recipe.py writes."""
    from examples.consolidation.recipe import recipe
    assert json.loads(RECIPE.read_text()) == json.loads(json.dumps(recipe()))


@needs_node
@needs_lo
def test_every_group_matches_the_reference_every_month(shown, ref):
    story, r = ref
    c = shown
    assert not any(c.calc.errors.values()), {s: e[:5] for s, e in c.calc.errors.items() if e}
    T = c.T
    for head, name in story.TIERS:
        g = r[head]
        zero = [0.0] * T
        for a in sorted(set(g["tb"]) | {1000}):
            if a in EQUITY:
                continue
            assert near(c.row(f"{B}/bg/{head}/cons/{a}"), g["tb"].get(a, zero)), (name, a)
        assert near(c.row(f"{B}/bg/{head}/cons/3400"), [-x for x in g["nci"]]), (name, "NCI")
        assert near(c.row(f"{B}/bg/{head}/cons/5900"), g["nci_pl"]), (name, "NCI share of the surplus")
        assert near(c.row(f"{B}/bg/{head}/surplus"), g["surplus"]), (name, "surplus")
        # the same total equity, however it is split between the owners and NCI (whose share of the surplus is in 5900)
        got = [sum(c.row(f"{B}/bg/{head}/cons/{a}")[t] for a in (*EQUITY, 5900)) for t in range(T)]
        exp = [sum(g["tb"].get(a, zero)[t] for a in EQUITY) for t in range(T)]
        assert near(got, exp), (name, "equity")


@needs_node
@needs_lo
def test_checks_are_clear_and_the_alert_is_the_accrued_on_charge(shown):
    c = shown
    for nm in ("Cn_Chk_TB", "Cn_Chk_Tree", "Cn_Chk_Join", "Cn_Chk_Pairs", "Cn_Chk_Inv", "Cn_Chk_Elim", "Cn_Chk_Cons", "Cn_Chk_NCI",
               "Cn_Chk_Top", "Cn_Chk_CF", "Cn_Alert_Planned"):
        assert c.value(f"{B}/bg/chk/{nm}") == 0, nm
    assert c.value(f"{B}/bg/chk/Cn_Alert_Netting") > 0
    assert (c.value("fm.checks#1/errors"), c.value("fm.checks#1/alerts")) == (0, 1)


def statement_checks(c: Calc, story, g: dict, head: int):
    T = c.T
    zero = [0.0] * T
    tb = lambda a: g["tb"].get(a, zero)            # noqa: E731
    move = lambda a, t: tb(a)[t] - (tb(a)[t - 1] if t else 0.0)   # noqa: E731
    # the income statement is each month's movement; the balance sheet the month end
    assert near(c.row(f"{B}/gs/is/revenue/4100"), [-move(4100, t) for t in range(T)])
    assert near(c.row(f"{B}/gs/is/cogs/5000"), [move(5000, t) for t in range(T)])
    assert near(c.row(f"{B}/gs/is/npat"), [g["surplus"][t] - (g["surplus"][t - 1] if t else 0.0) for t in range(T)])
    assert near(c.row(f"{B}/gs/bs/ca/1000"), tb(1000))
    assert near(c.row(f"{B}/gs/bs/nca/1250"), tb(1250))
    assert near(c.row(f"{B}/gs/bs/nci"), g["nci"])
    # the balance sheet balances and the cash flow ends at the cash on it
    assert near(c.row(f"{B}/gs/bs/check"), zero)
    assert near(c.row(f"{B}/gs/cf/check"), zero)
    assert near(c.row(f"{B}/gs/cf/end"), tb(1000))


@needs_node
@needs_lo
def test_statements_show_the_group_chosen(shown, ref, tmp_path_factory):
    story, r = ref
    statement_checks(shown, story, r[9000], 9000)
    holdings = Calc(build(tmp_path_factory.mktemp("cons2"), group=2))
    statement_checks(holdings, story, r[9004], 9004)
