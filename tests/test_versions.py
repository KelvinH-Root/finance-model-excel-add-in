"""Phase 0 proof: saved budgets and monthly reforecasts, and the comparisons that read them.

The budget and actuals example carries two approved budgets and a reforecast for every month
from April 2025 to September 2026, written from the reference. The LibreOffice tests read the
store back, move the comparisons, and run Save version live (as the add-in will through
Office.js): roll forward and save a reforecast, save and re-approve a budget, check that saved
values do not move when the model does, and replay a past month to show the live writer
stores exactly what the package writer did.
"""

import importlib.util
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
_spec = importlib.util.spec_from_file_location("reports_build_versions", ROOT / "prototypes" / "reports" / "build.py")
B = importlib.util.module_from_spec(_spec)
sys.modules["reports_build_versions"] = B
_spec.loader.exec_module(B)
V = B.V

FY = lambda y: slice(12 * (y - 1), 12 * y)        # noqa: E731


@pytest.fixture(scope="module")
def ref():
    inp = B.demo_inputs()
    hist = V.history(B, inp)
    return {"inp": inp, "hist": hist, "base": B.reference(inp, 1), "bud": B.reference_budget(inp)}


# ---------------------------------------------------------------- reference

def test_history_is_two_budgets_and_a_reforecast_every_month(ref):
    hist, inp = ref["hist"], ref["inp"]
    assert [v["id"] for v in hist] == list(range(1, 21))
    assert [v["label"] for v in hist if v["kind"] == "Budget"] == ["Budget FY2026", "Budget FY2027"]
    rfs = [v for v in hist if v["kind"] == "Reforecast"]
    assert [v["asat"] for v in rfs] == list(range(13, 31))
    assert rfs[0]["label"] == "Reforecast Apr 2025 (1+11)" and rfs[-1]["label"] == "Reforecast Sep 2026 (6+6)"
    assert [v["locked"] for v in rfs].count("No") == 1 and rfs[-1]["locked"] == "No"     # only this month's is open
    # The trend the history starts from is the current forecast after the last actual month
    trend = B.demo_inputs(last_actual=0)
    for g in ("rev", "cogs", "sal", "opx"):
        for k in range(len(inp[g])):
            assert trend[g][k][B.LAST_ACTUAL:] == inp[g][k][B.LAST_ACTUAL:]
    # This month's reforecast is the model as it stands; last month's saw a different future
    latest, prior = V.find(hist, "Reforecast", 30), V.find(hist, "Reforecast", 29)
    for key in ("rev", "npat", "cash", "te"):
        assert latest["values"][key] == pytest.approx(ref["base"][key])
    assert prior["values"]["rev"][:29] == pytest.approx(ref["base"]["rev"][:29])
    assert prior["values"]["rev"][29] != pytest.approx(ref["base"]["rev"][29])
    # Budgets: FY2027's equals the budget being built; FY2026's was phased on FY2025 alone
    b27, b26 = V.find(hist, "Budget", year=3), V.find(hist, "Budget", year=2)
    assert b27["values"]["rev"] == pytest.approx(ref["bud"]["rev"])
    assert set(b27["values"]) == {k for k, _ in B.BUDGET_LINES}
    assert b26["values"]["rev"][FY(2)] != pytest.approx(ref["bud"]["rev"][FY(2)])
    assert sum(b26["values"]["rev"][FY(2)]) == pytest.approx(sum(ref["bud"]["rev"][FY(2)]))   # same annual budget


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
    path = B.build(tmp_path_factory.mktemp("versions") / "versions_demo.xlsx")
    with libreoffice() as desktop:
        d = desktop.loadComponentFromURL("file://" + str(Path(path).resolve()), "_blank", 0, (_prop("Hidden", True),))
        d.calculateAll()
        try:
            yield d
        finally:
            d.close(True)


def value(c):
    """A cell's value: a number, a string, or None for an error or an empty result."""
    if c.getError():
        return None
    kind = c.getType().value
    if kind == "EMPTY":
        return None
    if kind == "TEXT" or (kind == "FORMULA" and c.FormulaResultType2 == 2):
        return c.getString() or None
    return c.getValue()


class Book:
    def __init__(self, doc):
        self.doc = doc

    def cell(self, sheet, ref):
        return self.doc.Sheets.getByName(sheet).getCellRangeByName(ref)

    def set(self, sheet, ref, text):
        self.cell(sheet, ref).setString(text)
        self.doc.calculateAll()

    def named(self, nm):
        return value(self.doc.NamedRanges.getByName(nm).getReferredCells().getCellByPosition(0, 0))

    def named_row(self, nm):
        return list(self.doc.NamedRanges.getByName(nm).getReferredCells().getDataArray()[0])

    def checks(self):
        return self.named("Chk_Errors"), self.named("Chk_Alerts")

    def raised(self):
        rows = self.doc.Sheets.getByName("Checks").getCellRangeByPosition(2, 0, 7, 200).getDataArray()
        return [r[0] for r in rows if isinstance(r[5], float) and r[5] != 0 and not r[0].endswith("raised")]

    def table(self, sheet, cid):
        col = self.doc.Sheets.getByName(sheet).getCellRangeByPosition(1, 0, 1, 1500).getDataArray()
        return next(i + 1 for i, (v,) in enumerate(col) if str(v).startswith(cid + "  "))

    def row(self, sheet, cid, k, n=12):
        """Values of the k-th row below a chart's header; None for #N/A or a blank slot."""
        sh = self.doc.Sheets.getByName(sheet)
        r = self.table(sheet, cid) + 2 + k
        out = [value(sh.getCellByPosition(c, r)) for c in range(9, 9 + n)]
        return [None if isinstance(v, str) else v for v in out]

    def labels(self, sheet, cid, k, n=12):
        sh = self.doc.Sheets.getByName(sheet)
        r = self.table(sheet, cid) + 2 + k
        return [sh.getCellByPosition(c, r).getString() for c in range(9, 9 + n)]

    def variance(self, line):
        """The Version comparison table's row for a line: 15 numbers (month, year to date, full year)."""
        sh = self.doc.Sheets.getByName(V.CMP_SHEET)
        labels = [r[0] for r in sh.getCellRangeByPosition(6, 0, 6, 40).getDataArray()]
        r = labels.index(line)
        return [None if sh.getCellByPosition(c, r).getError() else sh.getCellByPosition(c, r).getValue() for c in range(9, 24)]


def test_store_holds_every_version_as_saved(doc, ref):
    bk = Book(doc)
    assert bk.checks() == (0, 0)
    assert bk.named("Ver_Count") == 20 and bk.named("Ver_Latest_Asat") == 30
    for v in ref["hist"]:
        got = V.stored_live(doc, v["id"])
        assert set(got) == set(v["values"]), v["label"]
        for key, vals in v["values"].items():
            assert got[key] == pytest.approx(vals, abs=1e-6), (v["label"], key)
    now = bk.doc.NamedRanges.getByName("Ver_Sum_Now").getReferredCells().getDataArray()
    assert [r[0] for r in now[:20]] == pytest.approx([v["checksum"] for v in ref["hist"]])


def test_compared_with_reads_the_chosen_version(doc, ref):
    bk, hist = Book(doc), ref["hist"]
    ids = {v["label"]: v["id"] for v in hist}
    want = {"Budget": ids["Budget FY2027"], "Last month's reforecast": ids["Reforecast Aug 2026 (5+7)"],
            "Latest reforecast": ids["Reforecast Sep 2026 (6+6)"], "Budget being built": -1,
            "Reforecast Jun 2026 (3+9)": ids["Reforecast Jun 2026 (3+9)"]}
    try:
        for choice, vid in want.items():
            bk.set("Budget summary", "H7", choice)
            assert bk.named("BudS_Cmp") == vid, choice
            line = bk.row("Budget summary", "C17", 2)
            values = ref["bud"]["rev"] if vid == -1 else next(v for v in hist if v["id"] == vid)["values"]["rev"]
            assert line == pytest.approx(values[FY(3)]), choice
            assert bk.checks() == (0, 0), choice
        title = bk.cell("Budget summary", f"G{bk.table('Budget summary', 'C17') + 1}").getString()
        assert title == "Revenue against Reforecast Jun 2026 (3+9), FY2027"
        # A year with no approved budget: blank, raised as an alert; a typed choice that is not in the list is an error
        bk.set("Budget summary", "H7", "Budget")
        bk.set("Budget summary", "H4", "FY2028")
        assert bk.row("Budget summary", "C17", 2) == [None] * 12
        assert bk.checks() == (0, 1) and "Budget summary: nothing is saved" in bk.raised()[0]
        bk.set("Budget summary", "H4", "FY2026")
        assert bk.row("Budget summary", "C17", 2) == pytest.approx(V.find(hist, "Budget", year=2)["values"]["rev"][FY(2)])
        bk.set("Budget summary", "H7", "Forecast 3")
        assert bk.checks()[0] == 1
    finally:
        bk.set("Budget summary", "H4", "FY2027")
        bk.set("Budget summary", "H7", "Budget")
    assert bk.checks() == (0, 0)


def test_version_comparison_month_year_to_date_and_full_year(doc, ref):
    bk, hist, base = Book(doc), ref["hist"], ref["base"]
    bud, aug = V.find(hist, "Budget", year=3)["values"], V.find(hist, "Reforecast", 29)["values"]
    m, ys = 29, 24                                                                       # September 2026, FY2027 from April
    rev = bk.variance("Revenue")
    assert rev[:5] == pytest.approx([base["rev"][m], bud["rev"][m], base["rev"][m] - bud["rev"][m], aug["rev"][m],
                                     base["rev"][m] - aug["rev"][m]])
    ytd = lambda d, k: sum(d[k][ys:m + 1])                                                # noqa: E731
    assert rev[5:10] == pytest.approx([ytd(base, "rev"), ytd(bud, "rev"), ytd(base, "rev") - ytd(bud, "rev"),
                                       ytd(aug, "rev"), ytd(base, "rev") - ytd(aug, "rev")])
    full = lambda d, k: sum(d[k][FY(3)])                                                  # noqa: E731
    assert rev[10:] == pytest.approx([full(base, "rev"), full(bud, "rev"), full(base, "rev") - full(bud, "rev"),
                                      full(aug, "rev"), full(base, "rev") - full(aug, "rev")])
    sal = bk.variance("Salaries and wages")                                              # costs: favourable is spending less
    assert sal[2] == pytest.approx(bud["sal"][m] - base["sal"][m])
    # Full-year outturn by version: the budget, the reforecasts saved in FY2027 to September, then the current forecast
    trend = [bk.row(V.CMP_SHEET, "C96", k, 13) for k in (2, 3, 4)]
    assert trend[0][0] == pytest.approx(full(bud, "rev")) and trend[0][1:] == [None] * 12
    rfs = [full(V.find(hist, "Reforecast", a)["values"], "rev") for a in range(25, 31)]
    assert trend[1][1:7] == pytest.approx(rfs) and trend[1][7:12] == [None] * 5
    assert trend[2][12] == pytest.approx(full(base, "rev"))
    # What each version expected for September 2026: the budget, then the reforecasts as at September 2025 to August 2026
    acc = [bk.row(V.CMP_SHEET, "C98", k, 13) for k in (2, 3, 4)]
    assert bk.labels(V.CMP_SHEET, "C98", -1, 13)[:3] == ["Budget", "Sep 25", "Oct 25"]
    assert acc[0][0] == pytest.approx(bud["rev"][m])
    assert acc[1][1:] == pytest.approx([V.find(hist, "Reforecast", a)["values"]["rev"][m] for a in range(18, 30)])
    assert acc[2] == pytest.approx([base["rev"][m]] * 13)
    # The walk from the budget's profit after tax to the outturn's
    walk = bk.row(V.CMP_SHEET, "C99", 0, 9)
    assert walk[0] == pytest.approx(full(bud, "npat")) and walk[-1] == pytest.approx(full(base, "npat"))
    assert walk[1] == pytest.approx(full(base, "rev") - full(bud, "rev"))
    assert walk[2] == pytest.approx(full(bud, "cogs") - full(base, "cogs"))
    # The line shown and the month shown move the charts; a forecast month shows what is saved so far
    try:
        bk.set(V.CMP_SHEET, "H4", "Net profit after tax")
        assert bk.row(V.CMP_SHEET, "C96", 3, 13)[1:7] == pytest.approx([full(V.find(hist, "Reforecast", a)["values"], "npat")
                                                                         for a in range(25, 31)])
        bk.set(V.CMP_SHEET, "H4", "Revenue")
        bk.set(V.CMP_SHEET, "H5", "Dec 26")
        assert bk.cell(V.CMP_SHEET, "J11").getString() == "Forecast"
        acc = bk.row(V.CMP_SHEET, "C98", 3, 13)                                       # as at December 2025 to November 2026
        assert acc[1:11] == pytest.approx([V.find(hist, "Reforecast", a)["values"]["rev"][32] for a in range(21, 31)])
        assert acc[11:] == [None, None]                                                  # October and November are not saved yet
        assert bk.checks() == (0, 0)
    finally:
        bk.set(V.CMP_SHEET, "H4", "Revenue")
        bk.set(V.CMP_SHEET, "H5", "Sep 26")


def test_save_version_live(doc, ref):
    """Roll forward and save a reforecast, save and re-approve a budget, then replay a past month."""
    bk, inp, hist = Book(doc), ref["inp"], ref["hist"]
    ts = doc.NamedRanges.getByName("Ts_Last_Actual").getReferredCells().getCellByPosition(0, 0)
    # Month end for October 2026: the actuals are in, so roll forward; the reforecast is not saved yet
    ts.setValue(31)
    doc.calculateAll()
    assert bk.checks() == (0, 1) and "This month's reforecast is not saved yet" in bk.raised()[0]
    new = V.save_live(doc, B, "Reforecast")
    assert new == 21 and bk.checks() == (0, 0)
    reg = doc.Sheets.getByName(V.REG_SHEET)
    assert reg.getCellRangeByName(f"D{V.REG_FIRST + 20}").getString() == "Reforecast Oct 2026 (7+5)"
    want = B.reference(inp, 1, last_actual=31)
    got = V.stored_live(doc, new)
    for key in V.store_keys(B):
        assert got[key] == pytest.approx(want[key], abs=1e-6), key
    bk.set("Budget summary", "H7", "Latest reforecast")
    assert bk.named("BudS_Cmp") == 21
    bk.set("Budget summary", "H7", "Last month's reforecast")
    assert bk.named("BudS_Cmp") == V.find(hist, "Reforecast", 30)["id"]
    bk.set("Budget summary", "H7", "Budget")
    # Saved values do not move when the model does; editing the store is caught by the checksum
    inputs = doc.Sheets.getByName("Inputs")
    labels = [r[0] for r in inputs.getCellRangeByPosition(2, 0, 2, 200).getDataArray()]
    cell = inputs.getCellByPosition(B.FIRST_COL - 1 + 39, labels.index("Build contracts"))
    before = cell.getValue()
    cell.setValue(before + 500)
    doc.calculateAll()
    assert bk.named_row("St_Rev")[39] == pytest.approx(want["rev"][39] + 500)
    assert V.stored_live(doc, new)["rev"][39] == pytest.approx(want["rev"][39])
    assert bk.checks() == (0, 0)
    cell.setValue(before)
    store = doc.Sheets.getByName(V.STORE_SHEET)
    poke = store.getCellByPosition(B.FIRST_COL - 1 + 5, V.STORE_FIRST - 1)
    kept = poke.getValue()
    poke.setValue(kept + 1)
    doc.calculateAll()
    assert bk.checks()[0] == 1 and "values have changed since it was saved" in bk.raised()[0]
    poke.setValue(kept)
    doc.calculateAll()
    assert bk.checks() == (0, 0)
    # Next year's budget: none approved yet (an alert on the Budget summary), then saved and approved, then revised
    bk.set("Budget summary", "H4", "FY2028")
    assert bk.checks() == (0, 1)
    b28 = V.save_live(doc, B, "Budget", year=4, source="budget", locked="Yes")
    assert bk.named("BudS_Cmp") == b28 and bk.checks() == (0, 0)
    assert bk.row("Budget summary", "C17", 2) == pytest.approx(ref["bud"]["rev"][FY(4)])
    seasonality = doc.Sheets.getByName("Seasonality")
    seasonality.getCellRangeByName("M18").setValue(seasonality.getCellRangeByName("M18").getValue() + 1200)   # FY2028 budget raised
    doc.calculateAll()
    assert bk.row("Budget summary", "C17", 2) == pytest.approx(ref["bud"]["rev"][FY(4)])           # the approved budget holds
    b28r = V.save_live(doc, B, "Budget", year=4, source="budget", label="Budget FY2028 (revised)", locked="Yes")
    assert reg.getCellRangeByName(f"G{V.REG_FIRST + b28 - 1}").getString() == "Superseded"
    assert bk.named("BudS_Cmp") == b28r and bk.checks() == (0, 0)
    assert sum(bk.row("Budget summary", "C17", 2)) == pytest.approx(sum(ref["bud"]["rev"][FY(4)]) + 1200)
    bk.set("Budget summary", "H4", "FY2027")
    # Replay August 2026: the inputs as they stood then and the last actual month, saved live, equal the package's copy
    trend = B.demo_inputs(last_actual=0)
    V.set_inputs_live(doc, B, V.inputs_at(B, inp, trend, 29))
    ts.setValue(29)
    doc.calculateAll()
    replay = V.save_live(doc, B, "Other", label="Replay of Reforecast Aug 2026")
    got, saved = V.stored_live(doc, replay), V.find(hist, "Reforecast", 29)["values"]
    for key in V.store_keys(B):
        assert got[key] == pytest.approx(saved[key], abs=1e-6), key
    V.set_inputs_live(doc, B, inp)
    ts.setValue(31)
    doc.calculateAll()
    assert bk.checks() == (0, 0)
