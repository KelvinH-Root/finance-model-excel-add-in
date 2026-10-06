"""Phase 0 proof: every chart the summary and report modules bring, as native charts on a live model.

The structural tests read the register and the workbook's parts. The LibreOffice tests
recalculate the workbook, compare the statements and the chart tables with the Python
reference, and move the selections to show the charts follow them.
"""

import importlib.util
import re
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
# Loaded under its own name: the charts proof also has a build.py.
_spec = importlib.util.spec_from_file_location("reports_build", ROOT / "prototypes" / "reports" / "build.py")
B = importlib.util.module_from_spec(_spec)
sys.modules["reports_build"] = B
_spec.loader.exec_module(B)

REGISTER = B.load_register()
MODULES = {m["key"]: m for m in REGISTER["modules"]}
CHARTS = {c["id"]: c for c in REGISTER["charts"]}
SHEET = {c["id"]: MODULES[c["module"]]["title"] for c in REGISTER["charts"]}


@pytest.fixture(scope="module")
def demo(tmp_path_factory):
    return B.build(tmp_path_factory.mktemp("reports") / "reports_demo.xlsx")


def expected_series(ch: dict) -> int:
    r = ch["recipe"]
    if r == "compare":
        return len(ch.get("periods", ["prior", "shown", "next"]))
    if r in ("mix", "depth"):
        g = ch["group"]
        compare = 2 if (r == "mix" and ch.get("compare", True)) else 0
        if g in B.GROUPS:
            n = len(B.GROUPS[g][1])
            shown = min(ch.get("top") or n, n)
            other = 1 if ch.get("top") and ch["top"] < n else 0
            return (1 if ch.get("lead") else 0) + shown + other + compare
        spec = B.FIXED[g]
        return len(spec["stacks"]) + (1 if spec.get("current") else 0) + compare
    return ({"pie": 1, "budget": 3, "scenario": 3, "bridge": 4, "movement": 2, "trend": 3, "versions": 4, "accuracy": 3, "walk": 4}.get(r)
            or len(ch["bars"]) + len(ch["lines"]))


# ---------------------------------------------------------------- register and parts

def test_register_holds_every_chart_once():
    ids = [c["id"] for c in REGISTER["charts"]]
    assert ids == [f"C{i:02d}" for i in range(1, 100)]          # C96 to C99: the Version comparison module (HFG addition)
    counts = {m["title"]: sum(1 for c in REGISTER["charts"] if c["module"] == m["key"]) for m in REGISTER["modules"]}
    assert counts == {"Income summary": 5, "Balance summary": 6, "Cash summary": 5, "Budget summary": 6, "Income report": 28,
                      "Balance report": 7, "Cash report": 12, "Budget report": 6, "Scenario report": 20, "Version comparison": 4}
    lines = set(B.LABEL)
    allowed = {"id", "module", "title", "recipe", "line", "kind", "frame", "group", "top", "lead", "total", "compare", "cumulative",
               "periods", "bars", "grouping", "lines", "by"}
    for c in REGISTER["charts"]:
        assert set(c) <= allowed and all(v is not None for v in c.values()), c["id"]      # catches an unquoted comma in a title
        assert c["recipe"] in B.RECIPES or c["recipe"] in B.V.RECIPES, c["id"]
        for key in [c.get("line"), c.get("lead"), c.get("total")] + c.get("bars", []) + c.get("lines", []):
            assert key is None or key in lines, (c["id"], key)
        if "group" in c:
            assert c["group"] in B.GROUPS or c["group"] in B.FIXED or c["group"] in ("net_assets_build", "cash_bridge"), c["id"]
        if c["recipe"] == "budget":
            assert c["line"] in dict(B.BUDGET_LINES), c["id"]


def _parts(path):
    z = zipfile.ZipFile(path)
    wbx = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels) or re.search(rf'Target="([^"]+)"[^>]*Id="{rid}"', rels)
        sheet_part = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
        srels_name = sheet_part.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if srels_name not in z.namelist():
            continue
        m = re.search(r'Target="([^"]*drawing\d+\.xml)"', z.read(srels_name).decode())
        if not m:
            continue
        dpart = "xl/drawings/" + m.group(1).split("/")[-1]
        drawing = z.read(dpart).decode()
        drels = z.read(dpart.replace("drawings/", "drawings/_rels/") + ".rels").decode()
        charts = []
        for cname, crid in re.findall(r'<(?:xdr:)?cNvPr id="\d+" name="([^"]+)"/>.*?r:id="(rId\d+)"', drawing, re.S):
            ct = re.search(rf'Id="{crid}"[^>]*Target="([^"]+)"', drels) or re.search(rf'Target="([^"]+)"[^>]*Id="{crid}"', drels)
            charts.append((cname, z.read("xl/charts/" + ct.group(1).split("/")[-1]).decode()))
        out[name] = charts
    return out


def test_workbook_carries_the_register(demo):
    parts = _parts(demo)
    assert sum(len(v) for v in parts.values()) == 99
    for m in REGISTER["modules"]:
        charts = parts[m["title"]]
        wanted = [c for c in REGISTER["charts"] if c["module"] == m["key"]]
        assert [n for n, _ in charts] == [f"{c['id']} {c['title']}" for c in wanted]      # named by register id
        for (name, xml), c in zip(charts, wanted):
            assert len(re.findall(r"<(?:c:)?ser>", xml)) == expected_series(c), name
            assert "dispNaAsBlank" in xml                                                    # #N/A draws nothing in Excel
            assert f"'{m['title']}'!$G$" in xml                                              # title follows a cell on its sheet
            refs = set(re.findall(r"<f>'([^']+)'!", xml))
            assert refs == {m["title"]}, name                                                # a chart reads only its module's rows
            r = c["recipe"]
            if r == "pie":
                assert "<pieChart>" in xml
            elif r in ("movement",) or (r == "bridge" and c["group"] == "cash_bridge"):
                assert '<barDir val="bar"/>' in xml
            elif r in ("compare", "scenario") and c.get("kind", "line") == "line":
                assert "<lineChart>" in xml and "<barChart>" not in xml
            else:
                assert "<barChart>" in xml
            if r == "budget":
                assert 'prst="upDiag"' in xml and "<lineChart>" in xml                       # forecast hatched, budget a line
            if r == "combo":
                assert ("<lineChart>" in xml) == bool(c["lines"])
                assert f'<grouping val="{c["grouping"]}"/>' in xml


# ---------------------------------------------------------------- LibreOffice

@pytest.fixture(scope="module")
def doc(demo):
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")
    from hfgmodels.verify import _prop, libreoffice
    with libreoffice() as desktop:
        d = desktop.loadComponentFromURL("file://" + str(Path(demo).resolve()), "_blank", 0, (_prop("Hidden", True),))
        d.calculateAll()
        try:
            yield d
        finally:
            d.close(True)


@pytest.fixture(scope="module")
def ref():
    inp = B.demo_inputs()
    return {"inp": inp, "s": {s: B.reference(inp, s) for s in (1, 2, 3)}, "bud": B.reference_budget(inp)}


class Sheets:
    def __init__(self, doc):
        self.doc = doc

    def cell(self, sheet, ref):
        return self.doc.Sheets.getByName(sheet).getCellRangeByName(ref)

    def set(self, sheet, ref, text):
        self.cell(sheet, ref).setString(text)
        self.doc.calculateAll()

    def table(self, cid):
        """Heading row of a chart's table on its module sheet."""
        sh = self.doc.Sheets.getByName(SHEET[cid])
        col = sh.getCellRangeByPosition(1, 0, 1, 1500).getDataArray()
        return next(i + 1 for i, (v,) in enumerate(col) if str(v).startswith(cid + "  "))

    def row(self, cid, k, n=12):
        """Values of the k-th row below a chart's header (k=0 is the first series); None for #N/A."""
        sh = self.doc.Sheets.getByName(SHEET[cid])
        r = self.table(cid) + 2 + k
        out = []
        for c in range(9, 9 + n):
            cell = sh.getCellByPosition(c, r)
            out.append(None if cell.getError() else cell.getValue())
        return out

    def labels(self, cid, k, n=12):
        sh = self.doc.Sheets.getByName(SHEET[cid])
        r = self.table(cid) + 2 + k
        return [sh.getCellByPosition(c, r).getString() for c in range(9, 9 + n)]

    def label(self, cid, k):
        sh = self.doc.Sheets.getByName(SHEET[cid])
        return sh.getCellByPosition(6, self.table(cid) + 2 + k).getString()

    def errors(self):
        out = {}
        for i in range(self.doc.Sheets.Count):
            sh = self.doc.Sheets.getByIndex(i)
            for ra in sh.queryFormulaCells(4).getRangeAddresses():
                for r in range(ra.StartRow, ra.EndRow + 1):
                    for c in range(ra.StartColumn, ra.EndColumn + 1):
                        out.setdefault(sh.Name, []).append((r + 1, c + 1, sh.getCellByPosition(c, r).getError()))
        return out

    def checks(self):
        return self.named("Chk_Errors"), self.named("Chk_Alerts")

    def named(self, nm):
        rng = self.doc.NamedRanges.getByName(nm).getReferredCells()
        return rng.getCellByPosition(0, 0).getValue()


FY = lambda y: slice(12 * (y - 1), 12 * y)        # noqa: E731


def test_statements_match_the_reference_and_checks_are_clear(doc, ref):
    s = Sheets(doc)
    assert s.errors() == {}
    assert s.checks() == (0, 0)
    st = doc.Sheets.getByName("Statements")
    data = st.getCellRangeByPosition(0, 0, B.LAST_COL - 1, 600).getDataArray()
    labels = [row[2] or row[3] for row in data]
    base = ref["s"][1]
    # The active block comes first, so the first row with a label is the active scenario's.
    for key in ("rev", "cogs", "gm", "netopex", "ebitda", "npat", "cash", "deb", "ta", "tl", "na", "te", "receipts", "payments",
                "opcf", "chg", "wc_net"):
        r = labels.index(B.LABEL[key])
        got = [data[r][B.FIRST_COL - 1 + t] for t in range(B.PERIODS)]
        assert got == pytest.approx(base[key], abs=1e-6), key


def test_chart_tables_read_the_statements(doc, ref):
    s, base, bud = Sheets(doc), ref["s"][1], ref["bud"]
    rev = base["rev"]
    assert s.row("C23", 0) == pytest.approx(rev[FY(2)])                         # Revenue: FY2026, FY2027, FY2028
    assert s.row("C23", 1) == pytest.approx(rev[FY(3)])
    assert s.row("C23", 2) == pytest.approx(rev[FY(4)])
    assert s.row("C25", 1)[-1] == pytest.approx(sum(rev[FY(3)]))                 # cumulative ends at the year total
    # Top revenue lines for the year, ranked
    totals = sorted(((sum(base[f"rev_{k + 1}"][FY(3)]), B.GROUPS["rev"][1][k]) for k in range(5)), reverse=True)
    assert [s.label("C26", k) for k in range(5)] == [n for _, n in totals]
    assert [s.row("C26", k, 1)[0] for k in range(5)] == pytest.approx([v for v, _ in totals])
    # Operating expenses make-up: top four and the rest
    opx = sorted((sum(base[f"opx_{k + 1}"][FY(3)]) for k in range(7)), reverse=True)
    pie = s.row("C43", 0, 5)
    assert pie[:4] == pytest.approx(opx[:4]) and pie[4] == pytest.approx(sum(opx[4:]))
    # Expenses: salaries, the two largest other operating expenses, the rest; lines are total operating expenses
    assert s.row("C42", 0) == pytest.approx(base["sal"][FY(3)])
    stack = [sum(s.row("C42", k)[j] for k in range(4)) for j in range(12)]
    assert stack == pytest.approx(base["netopex"][FY(3)])
    assert s.row("C42", 4) == pytest.approx(base["netopex"][FY(2)])
    # Bridges close on the statements
    m = B.MONTH_SHOWN - 1
    na = s.row("C10", 0, 5)
    assert na[-1] == pytest.approx(base["na"][m]) and na[0] == pytest.approx(base["ca"][m])
    cash = s.row("C12", 0, 7)
    assert cash[0] == pytest.approx(base["cash"][23]) and cash[-1] == pytest.approx(base["cash"][35])
    assert cash[1] == pytest.approx(sum(base["receipts"][FY(3)]))
    # Movement against a year earlier
    assert s.row("C06", 2, 12)[0] == pytest.approx(base["cash"][m] - base["cash"][m - 12])
    # Rolling window: the 12 months to September 2026
    assert s.row("C07", 0) == pytest.approx(base["cash"][m - 11:m + 1])
    assert s.labels("C07", -1)[0] == "Oct 25" and s.labels("C07", -1)[-1] == "Sep 26"
    # Budget: actual to September, forecast after, budget line
    a, f, b = s.row("C17", 0), s.row("C17", 1), s.row("C17", 2)
    assert a[:6] == pytest.approx(rev[24:30]) and a[6:] == [0] * 6
    assert f[6:] == pytest.approx(rev[30:36]) and f[:6] == [0] * 6
    assert b == pytest.approx(bud["rev"][FY(3)])
    # Scenarios by year and by month
    for sc in (1, 2, 3):
        assert s.row("C76", sc - 1, 4) == pytest.approx([sum(ref["s"][sc]["rev"][FY(y)]) for y in range(1, 5)])
        assert s.row("C79", sc - 1, 4) == pytest.approx([ref["s"][sc]["na"][12 * y - 1] for y in range(1, 5)])
        assert s.row("C84", sc - 1) == pytest.approx(ref["s"][sc]["rev"][FY(3)])
    # Cash in and out, year to September 2026: sizes of the signed flows
    ytd = slice(24, 30)
    assert s.row("C63", 1, 5) == pytest.approx([abs(sum(base[k][ytd])) for k in ("receipts", "payments", "othop", "invcf", "fincf")])
    assert s.labels("C63", -1, 2) == ["Receipts (in)", "Payments (out)"]


def test_selections_move_every_chart_on_the_sheet(doc, ref):
    s, base = Sheets(doc), ref["s"][1]
    try:
        s.set("Income report", "H4", "FY2026")
        assert s.row("C23", 1) == pytest.approx(base["rev"][FY(2)])
        assert s.row("C49", 0) == pytest.approx(base["npat"][FY(1)])
        assert s.cell("Income report", f"G{s.table('C23') + 1}").getString() == "Revenue, FY2026"
        # Ranking follows the year: the largest operating expenses in FY2025 and FY2028 differ
        tops = {}
        for y in (1, 4):
            s.set("Income report", "H4", B.fy_label(y))
            tops[y] = [s.label("C45", k) for k in range(4)]
            want = sorted(range(7), key=lambda k: -sum(base[f"opx_{k + 1}"][FY(y)]))[:4]
            assert tops[y] == [B.GROUPS["opx"][1][k] for k in want]
        assert tops[1] != tops[4]
        # The first year has no year before it: those series are blank (#N/A), raised as an alert, not an error
        s.set("Income report", "H4", "FY2025")
        assert s.row("C23", 0) == [None] * 12
        errs = s.errors()
        assert set(errs) == {"Income report"} and all(e == 32767 for _, _, e in errs["Income report"])   # 32767 is #N/A
        assert s.checks() == (0, 1)
        # Month shown moves the rolling, movement and year-to-date charts
        s.set("Income report", "H4", "FY2027")
        s.set("Balance summary", "H5", "Mar 27")
        assert s.labels("C07", -1)[0] == "Apr 26"
        assert s.row("C06", 2, 1)[0] == pytest.approx(base["cash"][35] - base["cash"][23])
        s.set("Cash report", "H5", "Mar 27")
        assert s.row("C63", 0, 1)[0] == pytest.approx(sum(base["receipts"][FY(3)]))
        # The active scenario feeds the summaries and reports
        s.set("Scenarios", "H5", "Downside")
        assert s.row("C23", 1) == pytest.approx(ref["s"][3]["rev"][FY(3)])
        assert s.row("C17", 0)[:6] == pytest.approx(base["rev"][24:30])            # actuals do not move
        assert s.checks()[0] == 0
    finally:
        for sheet, ref_, text in (("Income report", "H4", "FY2027"), ("Balance summary", "H5", "Sep 26"),
                                  ("Cash report", "H5", "Sep 26"), ("Scenarios", "H5", "Base")):
            s.set(sheet, ref_, text)
    assert s.checks() == (0, 0)


def test_contents_covers_and_links_read_right(doc):
    s = Sheets(doc)
    contents = doc.Sheets.getByName("Contents").getCellRangeByPosition(0, 0, 6, 120).getDataArray()
    shown = [[str(v) for v in row[1:5] if v not in ("", None)] for row in contents]
    assert ["Demo Building Co"] in shown and ["Budget and actuals model"] in shown        # no checks failing, so no message
    assert ["1", "Dashboards"] in shown or ["1.0", "Dashboards"] in shown
    assert ["a.", "Income summary"] in shown and ["d.", "Seasonality"] in shown and ["-", "Profile"] in shown
    model = [s.cell("Model", f"B{r}").getString() for r in (9, 10, 11, 12, 13, 14)]
    assert model == ["Financial Model", "Section 2.", "Demo Building Co", "Go to contents", "< Version comparison", "Time >"]
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        if sh.Name != "Contents":
            assert sh.getCellRangeByName("A2").getString() == "✓", sh.Name             # every sheet shows the checks clear


def test_seasonality_phases_the_revenue_budget(doc, ref):
    s, inp = Sheets(doc), ref["inp"]
    profile = [s.cell("Seasonality", f"{B.L(B.FIRST_COL + j)}14").getValue() for j in range(12)]
    assert profile == pytest.approx(B.seasonality_profile(inp))
    assert s.cell("Seasonality", "I14").getValue() == pytest.approx(1)
    annual = inp["budget"]["rev_annual"][2]
    assert s.row("C17", 2) == pytest.approx([annual * p for p in profile])                # FY2027 budget line
    try:
        s.cell("Seasonality", "H7").setValue(0)                                              # leave FY2025 out
        doc.calculateAll()
        left_out = B.seasonality_profile(inp, include=(0, 1))
        assert [s.cell("Seasonality", f"{B.L(B.FIRST_COL + j)}14").getValue() for j in range(12)] == pytest.approx(left_out)
        assert s.row("C17", 2) == pytest.approx([annual * p for p in profile])                # the saved budget does not move
        s.set("Budget summary", "H7", "Budget being built")                                  # the budget being built does
        assert s.row("C17", 2) == pytest.approx([annual * p for p in left_out])
        s.set("Budget summary", "H7", "Budget")
        assert s.checks() == (0, 0)
        s.cell("Seasonality", "J13").setValue(0.2)                                           # a typed override that breaks 100%
        doc.calculateAll()
        assert s.cell("Seasonality", "J14").getValue() == pytest.approx(0.2)
        assert s.checks()[0] == 2                                                            # profile and phased total both flagged
    finally:
        s.cell("Seasonality", "H7").setValue(1)
        s.cell("Seasonality", "J13").setString("")
        doc.calculateAll()
    assert s.checks() == (0, 0)
