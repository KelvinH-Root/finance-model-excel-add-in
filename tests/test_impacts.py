"""Phase 0 proof: the Impacts command, both modes.

Live mode (prototypes/impacts/impact_live.py): change one input in an open model, recalculate,
read the statements and put the input back. On a model built by the assembly engine, every
movement must match the assembly reference with that input changed, the surplus must carry to
the accumulated surplus and the net cash flow to cash, and the workbook must come back exactly
as it was. The link chain explains the movement.

Sheets mode (prototypes/impacts/impact_sheets.py): Add to model writes one formula sheet per
kind of transaction the model holds, in its own accounts and entities. In the consolidation
example every sheet must match the Python reference for its defaults, each switch and a changed
input, and a broken entry must raise the check.
"""

import importlib.util
import itertools
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
for sub in ("impacts", "assembly", "consolidation"):
    sys.path.insert(0, str(ROOT / "prototypes" / sub))
sys.path.insert(0, str(ROOT))
import group as G  # noqa: E402
import impact_live as LV  # noqa: E402
import impact_sheets as IS  # noqa: E402
from assemble import Library, assemble, write_workbook  # noqa: E402
from demo import base_model  # noqa: E402
from reference import reference  # noqa: E402


@pytest.fixture(scope="module")
def demo():
    lib = Library.load()
    m = base_model(lib)
    return m, assemble(m)


# ---------------------------------------------------------------- the items a model is offered

def test_items_follow_the_models_accounts_and_entities(demo):
    m, layout = demo
    hfg = IS.from_consolidation(G)
    assert [i.key for i in IS.offered(hfg)] == ["sale", "opex", "facility", "interest", "equity", "land", "construction",
                                                "capint", "homes", "claim", "oncharge"]
    assert [i.key for i in IS.offered(IS.from_assembly(layout, m))] == ["sale", "opex", "facility", "interest"]
    no_wip = IS.Context(hfg.model, {k: v for k, v in hfg.accounts.items() if k != "wip"}, hfg.entities)
    assert {"land", "construction", "capint", "homes", "claim", "oncharge"}.isdisjoint(i.key for i in IS.offered(no_wip))
    # Labels come from the model: the assembly demo's own line names
    ctx = IS.from_assembly(layout, m)
    assert ctx.accounts["revenue"][0] == "Revenue line 1" and ctx.accounts["loans"][0] == "Debt facility 1"


def test_every_item_balances_and_ties_for_every_switch(demo):
    ctx = IS.from_consolidation(G)
    for item in IS.offered(ctx):
        names = [k for k, _, _ in item.switches]
        for combo in itertools.product((0, 1), repeat=len(names)):
            p = IS.panels(item, ctx, dict(zip(names, combo)))
            for col, v in p.items():
                assert v["balance"] == pytest.approx(0, abs=1e-9), (item.key, col, combo)
                assert v["cash_tie"] == pytest.approx(0, abs=1e-9), (item.key, col, combo)
            if item.group:
                g = p["Group"]
                # Nothing is earned inside the group; only an on-charge the receiver expenses is a group cost
                expensed = 120 * (1 - dict(zip(names, combo)).get("capitalised", 1)) if item.key == "oncharge" else 0
                assert g["surplus"] == pytest.approx(-expensed, abs=1e-9), (item.key, combo)
    homes = IS.panels(IS.BY_KEY["homes"], ctx)
    assert homes["Group"]["inv_property"] == pytest.approx(3520)          # the homes stay at the group's cost
    assert homes["Group"]["Investing"] == pytest.approx(0) and homes["Group"]["Operating"] == pytest.approx(0)
    claim = IS.panels(IS.BY_KEY["claim"], ctx)
    assert claim["Group"]["wip"] == pytest.approx(850) and claim["Group"]["cash"] == pytest.approx(-850)
    unreg = IS.panels(IS.BY_KEY["construction"], ctx, {"registered": 0})
    assert unreg[ctx.entities["devlp"]]["wip"] == pytest.approx(1150)        # GST it cannot claim is part of the cost


# ---------------------------------------------------------------- LibreOffice

@pytest.fixture(scope="module")
def office():
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")
    from hfgmodels.verify import libreoffice
    with libreoffice() as desktop:
        yield desktop


def _open(desktop, path):
    from hfgmodels.verify import _prop
    d = desktop.loadComponentFromURL("file://" + str(Path(path).resolve()), "_blank", 0, (_prop("Hidden", True),))
    d.calculateAll()
    return d


@pytest.fixture(scope="module")
def consolidation(office, tmp_path_factory):
    spec = importlib.util.spec_from_file_location("consolidation_build_imp", ROOT / "prototypes" / "consolidation" / "build.py")
    build = importlib.util.module_from_spec(spec)
    sys.modules["consolidation_build_imp"] = build
    spec.loader.exec_module(build)
    path = build.build(tmp_path_factory.mktemp("impacts") / "consolidation.xlsx")
    d = _open(office, path)
    yield d
    d.close(True)


def read_effect(doc, item, ctx):
    """The Effect on the statements block read back: {column: {key: value}} and the input and switch cells."""
    sh = doc.Sheets.getByName(item.sheet)
    data = sh.getCellRangeByPosition(0, 0, 14, 90).getDataArray()
    start = next(i for i, row in enumerate(data) if row[1] == "Effect on the statements")
    hdr = start + 1
    cols = {data[hdr][c]: c for c in range(9, 14) if data[hdr][c]}
    labels = {IS.line_label(ctx, k): k for k in list(ctx.accounts) + list(IS.FIXED)}
    out = {c: {} for c in cols}
    for row in data[hdr + 1:]:
        key = labels.get(row[2])
        if key:
            for c, j in cols.items():
                out[c][key] = row[j]
    inputs = {}
    for i, row in enumerate(data[5:start]):
        label = row[2]
        for name, lab, _ in item.inputs + item.switches:
            if lab == label:
                inputs[name] = sh.getCellByPosition(7, 5 + i)
    return out, inputs


def assert_matches(got, want, where):
    for col, keys in want.items():
        for k, v in keys.items():
            if k in got[col]:
                assert got[col][k] == pytest.approx(v, abs=1e-6), (where, col, k)


def test_impacts_sheets_match_the_reference_and_follow_their_switches(consolidation):
    doc = consolidation
    ctx = IS.from_consolidation(G)
    for item in IS.offered(ctx):
        got, cells = read_effect(doc, item, ctx)
        assert set(got) == set(IS.columns(item, ctx)), item.key
        assert_matches(got, IS.panels(item, ctx), (item.key, "defaults"))
        for name, label, default in item.switches:
            cells[name].setString("No" if default else "Yes")
            doc.calculateAll()
            got, _ = read_effect(doc, item, ctx)
            assert_matches(got, IS.panels(item, ctx, {name: 1 - default}), (item.key, name))
            cells[name].setString("Yes" if default else "No")
        name, _, default = item.inputs[0]
        cells[name].setValue(default * 2)
        doc.calculateAll()
        got, _ = read_effect(doc, item, ctx)
        assert_matches(got, IS.panels(item, ctx, {name: default * 2}), (item.key, name, "doubled"))
        cells[name].setValue(default)
        doc.calculateAll()
    nm = lambda n: doc.NamedRanges.getByName(n).getReferredCells().getCellByPosition(0, 0).getValue()   # noqa: E731
    assert (nm("Chk_Errors"), nm("Chk_Alerts")) == (0, 2)


def test_a_broken_entry_raises_the_impacts_check(consolidation):
    doc = consolidation
    sh = doc.Sheets.getByName(IS.BY_KEY["claim"].sheet)
    data = sh.getCellRangeByPosition(0, 0, 9, 40).getDataArray()
    r = next(i for i, row in enumerate(data) if row[2] == "4100 Construction revenue")
    cell = sh.getCellByPosition(7, r)
    kept = cell.getFormula()
    nm = lambda n: doc.NamedRanges.getByName(n).getReferredCells().getCellByPosition(0, 0).getValue()   # noqa: E731
    try:
        cell.setValue(-900)                       # someone types over the revenue entry
        doc.calculateAll()
        assert nm("Chk_Errors") == 1
    finally:
        cell.setFormula(kept)
        doc.calculateAll()
    assert nm("Chk_Errors") == 0


@pytest.fixture(scope="module")
def built_demo(office, demo, tmp_path_factory):
    m, layout = demo
    path = write_workbook(layout, tmp_path_factory.mktemp("impacts_live") / "base.xlsx", m)
    d = _open(office, path)
    yield d
    d.close(True)


CASES = [("Rev1_Base", 110.0, "demo.revenue_line#1", "base"), ("Cost1_Amount", 99.0, "demo.cost_line#1", "amount"),
         ("Fac1_Amount", 1500.0, "demo.facility#1", "amount"), ("Rev2_Growth", 0.05, "demo.revenue_line#2", "growth"),
         ("Fac1_Rate", 0.075, "demo.facility#1", "rate")]


@pytest.mark.parametrize("name,value,uid,key", CASES)
def test_live_impact_matches_the_reference_and_leaves_the_model_as_it_was(built_demo, demo, name, value, uid, key):
    m, layout = demo
    res = LV.run(built_demo, layout, name, value)
    assert res["restored"], "the workbook changed"
    changed = m.copy()
    next(i for i in changed.instances if i.uid == uid).settings[key] = value
    r0, r1 = reference(m), reference(changed)
    for k in ("revenue", "opex", "interest", "receipts", "cash", "assets", "debt", "equity"):
        want = [b - a for a, b in zip(r0[k], r1[k])]
        assert res["delta"][f"demo.statements#1/{k}"] == pytest.approx(want, abs=1e-6), (name, k)
    d = res["delta"]
    run_surplus = list(itertools.accumulate(d["demo.statements#1/surplus"]))
    run_cash = list(itertools.accumulate(d["demo.statements#1/net_cash"]))
    assert d["demo.statements#1/equity"] == pytest.approx(run_surplus, abs=1e-6)       # surplus carried to equity
    assert d["demo.statements#1/cash"] == pytest.approx(run_cash, abs=1e-6)            # net cash flow carried to cash
    assert [a - b - c for a, b, c in zip(d["demo.statements#1/assets"], d["demo.statements#1/debt"],
                                         d["demo.statements#1/equity"])] == pytest.approx([0] * m.periods, abs=1e-6)


def test_the_chain_explains_where_a_change_goes(demo):
    _, layout = demo
    assert LV.chain(layout, "demo.revenue_line#1") == [
        ("Revenue line 1", "is.revenue", "Debtors: Revenue line 1"), ("Revenue line 1", "is.revenue", "Financial statements"),
        ("Debtors: Revenue line 1", "cf.receipts", "Financial statements"),
        ("Debtors: Revenue line 1", "bs.debtors", "Financial statements")]
    assert {e[1] for e in LV.chain(layout, "demo.facility#1")} == {"is.interest", "cf.financing", "bs.debt"}
    assert LV.setting_of(layout, "Fac1_Rate") == ("demo.facility#1", "rate")
