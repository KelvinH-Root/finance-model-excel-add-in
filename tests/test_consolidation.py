"""Phase 0 proof: group consolidation on a fictional group shaped like HFG's.

The reference (prototypes/consolidation/group.py) re-records every event as each group sees
it and never reads the intercompany register. The workbook eliminates from the register with
formulas. The tests check the two agree for every group, account and year, and that the
HFG-specific cases come out right: at-cost on-charges through the netting accounts (one into
WIP), margin on construction, fees and interest capitalised into WIP, two portfolio sales into
the Fund (one from a partly owned LP), homes sold outside the group, GST a receiver cannot claim,
and NCI worked out once at each node.
"""

import importlib.util
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "prototypes" / "consolidation"))
import group as G  # noqa: E402

FIRST = 10
Y = G.PERIODS


@pytest.fixture(scope="module")
def ref():
    b = G.book()
    return {"book": b, "tb": G.trial_balances(b), "groups": G.reference(b)}


# ---------------------------------------------------------------- reference

def test_entity_books_balance_and_the_netting_accounts_clear(ref):
    tb = ref["tb"]
    for e in G.CODES:
        for p in range(Y):
            assert sum(tb[(e, a)][p] for a, _, _ in G.ACCOUNTS) == pytest.approx(0, abs=1e-9), (e, p)
    # Only Foundation's netting account with Dev LP A carries a balance, at FY2026's cut-off, and Dev LP A accrued it
    assert tb[(9000, G.CLR)] == pytest.approx([20, 0, 0])
    assert tb[(9006, G.ACCR)] == pytest.approx([-20, 0, 0])
    # The tiers: the lowest group holding both sides
    assert G.TIER_MEMBERS[9005] == {9005, 9006, 9007} and G.TIER_MEMBERS[9008] == {9008, 9009}


def test_reference_follows_hfg_rules(ref):
    g = ref["groups"]
    top, hold, devco, fund = g[9000], g[9004], g[9005], g[9008]
    # Site A's group cost in the Foundation group: BuildCo's and DevManager's costs plus the shared cost at cost
    assert top["basis"]["SA"][1] == pytest.approx(1700 + 850 + 210 + 70 + 120)
    # In the Holdings group BuildCo and DevManager sit outside, so their charges are costs at price
    assert hold["basis"]["SA"][1] == pytest.approx(3520)
    # Interest on the Holdings loan is a group cost nowhere it is eliminated: Holdings and above
    assert hold["basis"]["SB"][2] == pytest.approx(1000 + 150 + 120 + 60 + 1200 + 500)
    assert devco["basis"]["SB"][2] == pytest.approx(1000 + 150 + 120 + 40 + 60 + 1200 + 43.2 + 500 + 46.656)
    # NCI: the Fund's outside investors (40%) and Partner LP's (30%), less 30% of Partner LP's margin still held
    assert top["nci"][0] == pytest.approx(0.4 * 2000 + 0.3 * 1000)
    assert top["nci"][1] == pytest.approx(0.4 * (2000 + fund["surplus"][1]) + 0.3 * (1000 + 300) - 0.3 * 300)
    # FY2028: five of site A's twenty homes sold outside the group for 1,300, at each group's own cost of the site
    assert top["surplus"][2] == pytest.approx(120 + 120 + (650 - 130 - 192) - 18 + (1300 - 0.25 * 2950) - 30)
    assert hold["surplus"][2] == pytest.approx((650 - 130 - 192 - 138) + (1300 - 0.25 * 3520) - 30)


# ---------------------------------------------------------------- LibreOffice

@pytest.fixture(scope="module")
def doc(tmp_path_factory):
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")
    from hfgmodels.verify import _prop, libreoffice
    spec = importlib.util.spec_from_file_location("consolidation_build", ROOT / "prototypes" / "consolidation" / "build.py")
    build = importlib.util.module_from_spec(spec)
    sys.modules["consolidation_build"] = build
    spec.loader.exec_module(build)
    path = build.build(tmp_path_factory.mktemp("consolidation") / "consolidation_demo.xlsx")
    with libreoffice() as desktop:
        d = desktop.loadComponentFromURL("file://" + str(Path(path).resolve()), "_blank", 0, (_prop("Hidden", True),))
        d.calculateAll()
        try:
            yield d
        finally:
            d.close(True)


class Book:
    def __init__(self, doc):
        self.doc = doc
        self.grp = doc.Sheets.getByName("By group")
        col = self.grp.getCellRangeByPosition(1, 0, 2, 400).getDataArray()
        self.blocks = {}
        head = None
        for i, (a, b) in enumerate(col):
            text = str(a)
            if text.endswith(")") and "(head " in text:
                head = int(text.split("(head ")[1].rstrip(")"))
                self.blocks[head] = {}
            elif head is not None and isinstance(a, float) and a.is_integer() and int(a) in G.ACCT:
                self.blocks[head][int(a)] = i

    def cons(self, head, acct):
        r = self.blocks[head][acct]
        return [self.grp.getCellByPosition(FIRST - 1 + 2 * Y + p, r).getValue() for p in range(Y)]

    def named(self, nm):
        return self.doc.NamedRanges.getByName(nm).getReferredCells().getCellByPosition(0, 0).getValue()

    def checks(self):
        return self.named("Chk_Errors"), self.named("Chk_Alerts")

    def raised(self):
        rows = self.doc.Sheets.getByName("Checks").getCellRangeByPosition(2, 0, 7, 80).getDataArray()
        return [r[0] for r in rows if isinstance(r[5], float) and r[5] != 0 and not str(r[0]).endswith("raised")]

    def errors(self):
        out = {}
        for i in range(self.doc.Sheets.Count):
            sh = self.doc.Sheets.getByIndex(i)
            for ra in sh.queryFormulaCells(4).getRangeAddresses():
                out.setdefault(sh.Name, 0)
                out[sh.Name] += (ra.EndRow - ra.StartRow + 1) * (ra.EndColumn - ra.StartColumn + 1)
        return out


def test_every_group_matches_the_reference(doc, ref):
    bk = Book(doc)
    assert bk.errors() == {}
    assert bk.checks() == (0, 2)        # alerts: Foundation's FY2026 on-charge accrued at the cut-off; GST the Fund cannot claim
    for head, name in G.TIERS:
        r = ref["groups"][head]
        for acct, label, cls in G.ACCOUNTS:
            if cls in ("asset", "liability", "revenue", "expense"):
                assert bk.cons(head, acct) == pytest.approx(r["tb"][acct], abs=1e-6), (name, acct, label)
        assert bk.cons(head, G.DIFF) == pytest.approx([0] * Y)
        assert [-x for x in bk.cons(head, G.NCI)] == pytest.approx(r["nci"], abs=1e-6), name
        assert bk.cons(head, G.NCIPL) == pytest.approx(r["nci_pl"], abs=1e-6), name


def test_hfg_cases_come_out_right(doc, ref):
    bk = Book(doc)
    top, hold, devco = 9000, 9004, 9005
    # At-cost on-charges: the P&L is untouched and the shared cost sits in the group once (Foundation 100 + PropManager 80)
    assert bk.cons(top, G.OPEX)[0] == pytest.approx(ref["groups"][top]["tb"][G.OPEX][0])
    assert bk.cons(top, G.CLR) == pytest.approx([0, 0, 0]) and bk.cons(top, G.ACCR) == pytest.approx([0, 0, 0])
    # Construction and fee margins come out of WIP in the Foundation group only; Holdings sees them as related party costs
    wip_entities = sum(ref["tb"][(e, G.WIP)][0] for e in G.CODES)
    assert bk.cons(top, G.WIP)[0] == pytest.approx(wip_entities - (300 + 150 + 90) - (90 + 45) - 40)
    assert bk.cons(hold, G.WIP)[0] == pytest.approx(wip_entities - 40)               # only the capitalised Holdings interest
    assert bk.cons(devco, G.WIP)[0] == pytest.approx(sum(ref["tb"][(e, G.WIP)][0] for e in G.TIER_MEMBERS[devco]))
    # The portfolio sale's margin is out of the Fund's property in Holdings and above, but not in the Devco group
    assert bk.cons(devco, G.SALES)[1] == pytest.approx(-4200)
    assert bk.cons(hold, G.SALES)[1] == pytest.approx(0)
    assert bk.cons(hold, G.PROP)[1] == pytest.approx(3520 + 1000)
    # Homes sold outside the group in FY2028: the cost of sale is a quarter of the group's cost of site A
    assert bk.cons(top, G.COS)[2] == pytest.approx(0.25 * 2950)
    assert bk.cons(hold, G.COS)[2] == pytest.approx(0.25 * 3520)
    # GST the Fund cannot claim on PropManager's fee stays as a group cost
    assert bk.cons(top, G.GSTPAY)[1] == pytest.approx(-15)
    # NCI: Partner LP's outside investors bear 30% of its margin on the sale to the Fund while the homes are held
    assert -bk.cons(hold, G.NCI)[1] == pytest.approx(0.4 * 2029 + 0.3 * 1300 - 0.3 * 300)


def test_statements_follow_the_group_and_year_shown(doc, ref):
    bk = Book(doc)
    st = doc.Sheets.getByName("Group statements")
    labels = [r[0] for r in st.getCellRangeByPosition(2, 0, 2, 120).getDataArray()]
    row = {lab: i for i, lab in enumerate(labels) if lab}

    def show(group, year):
        st.getCellRangeByName("H5").setString(group)
        st.getCellRangeByName("H6").setString(year)
        doc.calculateAll()
        return lambda lab: st.getCellByPosition(7, row[lab]).getValue()

    try:
        for head, name in G.TIERS:
            for p, year in enumerate(G.YEAR_LABELS):
                v = show(name, year)
                r = ref["groups"][head]
                assert v("Surplus for the year") == pytest.approx(r["surplus"][p], abs=1e-6), (name, year)
                assert v("Net assets") == pytest.approx(r["net_assets"][p], abs=1e-6), (name, year)
                assert v("Non-controlling interests") == pytest.approx(r["nci"][p], abs=1e-6), (name, year)
                assert v("Net assets") == pytest.approx(v("Equity"))
        # Related parties: in the Holdings group, BuildCo's construction claims are related party purchases
        v = show("Holdings group", "FY2026")
        cap = [i for i, lab in enumerate(labels) if lab == "Trading, capitalised"][0]
        assert st.getCellByPosition(8, cap).getValue() == pytest.approx(2000 + 1000 + 600 + 300 + 150)
    finally:
        show(G.TIERS[0][1], G.YEAR_LABELS[-1])
    assert bk.checks() == (0, 2)


def test_checks_catch_breaks_and_unraised_on_charges(doc):
    bk = Book(doc)
    ic = doc.Sheets.getByName("Intercompany")
    data = doc.Sheets.getByName("Entity data")
    # A buyer recording a different amount: a break, and the difference left in every group holding both sides
    cell = ic.getCellRangeByName("M8")
    kept = cell.getValue()
    try:
        cell.setValue(kept - 10)
        doc.calculateAll()
        raised = bk.raised()
        assert "The two sides of every intercompany pair agree (a break is a timing difference or an error to clear)" in raised
        assert bk.checks()[0] >= 1
    finally:
        cell.setValue(kept)
        doc.calculateAll()
    assert bk.checks() == (0, 2)
    # A netting account balance with no on-charge in the register
    vals = data.getCellRangeByPosition(1, 0, 2, 400).getDataArray()
    r = next(i for i, (e, a) in enumerate(vals) if e == 9003.0 and a == 2050.0)
    c = data.getCellByPosition(FIRST - 1 + 1, r)
    try:
        c.setValue(5)
        cash = data.getCellByPosition(FIRST - 1 + 1, next(i for i, (e, a) in enumerate(vals) if e == 9003.0 and a == 1000.0))
        cash.setValue(cash.getValue() - 5)
        doc.calculateAll()
        assert "Each netting account's balance is in the register as an on-charge still to raise" in bk.raised()
    finally:
        c.setValue(0)
        cash.setValue(cash.getValue() + 5)
        doc.calculateAll()
    assert bk.checks() == (0, 2)
