"""Phase 0 proof: modules inserted into a built model link themselves in.

The pure Python tests check link resolution and the change plan. The LibreOffice
tests apply the plan to a built workbook (as the add-in's live writer would),
compare it cell by cell with the same model built from scratch, and check the
numbers against an independent calculation.
"""

import shutil
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent.parent / "prototypes" / "assembly"
sys.path.insert(0, str(HERE))

from reference import reference  # noqa: E402
from assemble import (FIRST_PERIOD_COL, TOTAL_COL, AssemblyError, Library, Model,  # noqa: E402
                      assemble, open_model, plan_change, read_metadata, write_workbook)
from demo import base_model, insert_two  # noqa: E402

PERIODS = 12


@pytest.fixture()
def lib():
    return Library.load()


def rows_of(layout, sheet):
    return {r.id: r for r in dict(layout.sheets)[sheet]}


# ---------------------------------------------------------------- resolution

def test_new_revenue_line_reaches_every_consumer(lib):
    m = base_model(lib)
    before = assemble(m)
    inst = m.insert("demo.revenue_line", base=150)
    after = assemble(m)
    st = rows_of(after, "Statements")
    uid = inst.uid
    # Statements collects it directly (is.revenue) and through the Debtors mirror (cf.receipts, bs.debtors).
    assert f"demo.statements#1/in/is.revenue/{uid}" in st
    assert f"demo.statements#1/in/cf.receipts/demo.debtors#1/{uid}" in st
    assert f"demo.statements#1/in/bs.debtors/demo.debtors#1/{uid}" in st
    # Debtors grew a block for it without anyone asking.
    assert f"demo.debtors#1/{uid}/subheading" in rows_of(after, "Working capital")
    # Total revenue now spans three rows.
    total = st["demo.statements#1/revenue"].formula
    assert total.count("«R|") == 2 and "revenue_line#1" in total and uid in total
    plan = plan_change(before, after)
    assert any(o["op"] == "insert_rows" and o["sheet"] == "Working capital" for o in plan.ops)
    rewired = [o for o in plan.ops if o["op"] == "write" and o["why"] == "rewire"]
    assert {o["sheet"] for o in rewired} == {"Statements"}


def test_second_facility_adds_interest_cash_debt_and_check_rows(lib):
    m = base_model(lib)
    inst = m.insert("demo.facility", amount=500)
    after = assemble(m)
    st = rows_of(after, "Statements")
    for link in ("is.interest", "cf.financing", "bs.debt"):
        assert f"demo.statements#1/in/{link}/{inst.uid}" in st
    assert f"demo.checks#1/in/check.error/{inst.uid}" in rows_of(after, "Checks")
    assert "Fac2_Rate" in after.names


def test_single_modules_insert_once(lib):
    m = base_model(lib)
    with pytest.raises(AssemblyError, match="only be inserted once"):
        m.insert("demo.statements")
    with pytest.raises(AssemblyError, match="no setting"):
        m.insert("demo.revenue_line", price=1)


def test_unresolved_and_orphan_links_are_reported(lib):
    m = Model(lib)
    m.insert("demo.statements")
    assert any("required link is.revenue" in w for w in assemble(m).warnings)
    m = Model(lib)
    m.insert("demo.revenue_line")
    assert any("is.revenue is not taken" in w for w in assemble(m).warnings)


def test_total_mode_adds_one_row_for_all_senders(tmp_path):
    (tmp_path / "areas.yaml").write_text("areas: [In, Out]\n")
    (tmp_path / "src.yaml").write_text(
        "id: t.src\ntitle: Source\ncode: Src\narea: In\nas_category: true\n"
        "settings: [{key: v, label: Value, default: 1}]\n"
        "rows: [{key: x, label: X, formula: '=$v'}]\noutputs: [{link: t.x, row: x}]\n")
    (tmp_path / "sink.yaml").write_text(
        "id: t.sink\ntitle: Sink\narea: Out\ninputs: [{link: t.x, mode: total}]\n"
        "rows: [{collect: t.x}, {key: y, label: Y, formula: '=[sum:t.x]*2'}]\n")
    m = Model(Library.load(tmp_path))
    m.insert("t.sink")
    for v in (1, 2, 3):
        m.insert("t.src", v=v)
    out = rows_of(assemble(m), "Out")
    total = out["t.sink#1/in/t.x"]
    assert total.formula.count("«R|") == 3
    assert out["t.sink#1/y"].formula == "=SUM(«R|t.sink#1/in/t.x»:«R|t.sink#1/in/t.x»)*2"


def test_removing_a_module_rewires_and_leaves_no_dangling_reference(lib):
    m = base_model(lib)
    before = assemble(m)
    m.remove("demo.revenue_line#1")
    after = assemble(m)
    plan = plan_change(before, after)
    assert any(o["op"] == "delete_rows" for o in plan.ops)
    assert any(o["op"] == "delete_name" and o["name"] == "Rev1_Base" for o in plan.ops)
    live_ids = {r.id for _, rows in after.sheets for r in rows}
    for _, rows in after.sheets:
        for r in rows:
            for f in (r.first, r.formula):
                for rid in __import__("re").findall(r"«[RP]\|([^»]+)»", f or ""):
                    assert rid in live_ids


def test_instance_numbers_are_never_reused(lib):
    m = base_model(lib)
    m.remove("demo.revenue_line#2")
    assert m.insert("demo.revenue_line").number == 3


def test_metadata_round_trip(lib, tmp_path):
    m = base_model(lib)
    layout = assemble(m)
    path = write_workbook(layout, tmp_path / "m.xlsx", m)
    model, meta = open_model(path, lib)
    assert model.to_dict() == m.to_dict()
    assert meta["rows"] == {s: [r.id for r in rows] for s, rows in layout.sheets}
    assert read_metadata(path)["records"] == layout.records


# ---------------------------------------------------------------- in LibreOffice

def _libreoffice():
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")


def _trim(grid):
    g = [list(r) for r in grid]
    while g and all(v in ("", None) for v in g[-1]):
        g.pop()
    width = max((max((i + 1 for i, v in enumerate(r) if v not in ("", None)), default=0) for r in g), default=0)
    return [(r + [""] * width)[:width] for r in g]


def assert_same(a, b):
    assert list(a["sheets"]) == list(b["sheets"])
    assert a["names"] == b["names"]
    assert a["charts"] == b["charts"]
    for s in a["sheets"]:
        for kind in ("formulas", "values"):
            ga, gb = _trim(a["sheets"][s][kind]), _trim(b["sheets"][s][kind])
            assert len(ga) == len(gb), (s, kind)
            for r, (ra, rb) in enumerate(zip(ga, gb), start=1):
                for c, (va, vb) in enumerate(zip(ra, rb), start=1):
                    if isinstance(va, float) and isinstance(vb, float):
                        assert va == pytest.approx(vb, abs=1e-9), (s, kind, r, c)
                    else:
                        assert va == vb, (s, kind, r, c)
        assert not a["errors"][s] and not b["errors"][s], s


def check_numbers(snap, layout, model):
    ref = reference(model)
    pos = layout.positions()
    values = snap["sheets"]["Statements"]["values"]
    for key, series in ref.items():
        _, row = pos[f"demo.statements#1/{key}"]
        got = values[row - 1][FIRST_PERIOD_COL - 1:FIRST_PERIOD_COL - 1 + model.periods]
        assert got == pytest.approx(series, abs=1e-6), key
    _, row = pos["demo.statements#1/bs_check"]
    assert all(v == 0 for v in values[row - 1][FIRST_PERIOD_COL - 1:])
    _, row = pos["demo.checks#1/errors"]
    assert snap["sheets"]["Checks"]["values"][row - 1][TOTAL_COL - 1] == 0


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    _libreoffice()
    from live import apply_plan, snapshot
    out = tmp_path_factory.mktemp("assembly")
    lib = Library.load()
    base = base_model(lib)
    write_workbook(assemble(base), out / "base.xlsx", base)

    # Insert two modules into the built file, starting from what the file says about itself.
    model, _ = open_model(out / "base.xlsx", lib)
    old = assemble(model)
    insert_two(model)
    new = assemble(model)
    apply_plan(out / "base.xlsx", plan_change(old, new, "uno"), out / "live.xlsx", model, new)
    write_workbook(new, out / "fresh.xlsx", model)

    # Then remove the first revenue line and the first facility from the live file.
    model2, _ = open_model(out / "live.xlsx", lib)
    old2 = assemble(model2)
    model2.remove("demo.revenue_line#1")
    model2.remove("demo.facility#1")
    new2 = assemble(model2)
    apply_plan(out / "live.xlsx", plan_change(old2, new2, "uno"), out / "live2.xlsx", model2, new2)
    write_workbook(new2, out / "fresh2.xlsx", model2)
    return {"snap": {k: snapshot(out / f"{k}.xlsx") for k in ("base", "live", "fresh", "live2", "fresh2")},
            "layouts": {"base": assemble(base), "live": new, "live2": new2},
            "models": {"base": base, "live": model, "live2": model2}}


def test_base_model_matches_reference(built):
    check_numbers(built["snap"]["base"], built["layouts"]["base"], built["models"]["base"])


def test_inserting_into_a_built_workbook_equals_building_fresh(built):
    assert_same(built["snap"]["live"], built["snap"]["fresh"])
    check_numbers(built["snap"]["live"], built["layouts"]["live"], built["models"]["live"])


def test_removing_from_a_built_workbook_equals_building_fresh(built):
    assert_same(built["snap"]["live2"], built["snap"]["fresh2"])
    check_numbers(built["snap"]["live2"], built["layouts"]["live2"], built["models"]["live2"])


# ---------------------------------------------------------------- module charts

def test_a_module_carries_its_chart_and_new_lines_join_it(lib):
    m = base_model(lib)
    m.insert("demo.dashboard")
    before = assemble(m)
    (chart,) = before.charts
    assert chart.sheet == "Dashboard" and chart.span == 6
    assert [k for _, k in chart.series] == ["column", "column", "line"]
    inst = m.insert("demo.revenue_line", base=150)
    after = assemble(m)
    assert f"demo.dashboard#1/rev/{inst.uid}" in [rid for rid, _ in after.charts[0].series]
    plan = plan_change(before, after)
    (op,) = [o for o in plan.ops if o["op"].endswith("_chart")]
    assert op["op"] == "set_chart" and len(op["series"]) == 4
    assert any("re-pointed (3 column series and 1 line" in line for line in plan.preview)
    m.remove("demo.dashboard#1")
    assert [o["op"] for o in plan_change(after, assemble(m)).ops if o["op"].endswith("_chart")] == []  # its sheet goes with it


def test_inserting_the_module_adds_its_chart(lib):
    m = base_model(lib)
    before = assemble(m)
    m.insert("demo.dashboard")
    plan = plan_change(before, assemble(m))
    (op,) = [o for o in plan.ops if o["op"].endswith("_chart")]
    assert op["op"] == "add_chart" and op["anchor"]["row"] == 7
    assert op["series"][0]["values"] == "Dashboard!$J$11:$O$11"   # six months, the chart window


@pytest.fixture(scope="module")
def charted(tmp_path_factory, built):
    from live import apply_plan, snapshot
    out = tmp_path_factory.mktemp("charts")
    lib = Library.load()
    base = base_model(lib)
    base.insert("demo.dashboard", first=1)
    write_workbook(assemble(base), out / "base.xlsx", base)
    steps = {}

    def step(name, src, change):
        model, _ = open_model(src, lib)
        old = assemble(model)
        change(model)
        new = assemble(model)
        apply_plan(src, plan_change(old, new, "uno"), out / f"live_{name}.xlsx", model, new)
        write_workbook(new, out / f"fresh_{name}.xlsx", model)
        steps[name] = (model, new)
        return out / f"live_{name}.xlsx"

    live1 = step("insert_line", out / "base.xlsx", lambda m: m.insert("demo.revenue_line", base=150, growth=0.02))
    step("remove_line", live1, lambda m: m.remove("demo.revenue_line#1"))
    plain = out / "plain.xlsx"
    pm = base_model(lib)
    write_workbook(assemble(pm), plain, pm)
    step("insert_dashboard", plain, lambda m: m.insert("demo.dashboard", first=1))
    later = base_model(lib)
    later.insert("demo.dashboard", first=4)
    write_workbook(assemble(later), out / "first4.xlsx", later)
    files = [f"{k}_{n}" for n in steps for k in ("live", "fresh")] + ["first4"]
    return {"snap": {f: snapshot(out / f"{f}.xlsx") for f in files}, "steps": steps,
            "first4": (later, assemble(later))}


@pytest.mark.parametrize("step", ["insert_line", "remove_line", "insert_dashboard"])
def test_charts_in_a_built_workbook_equal_building_fresh(charted, step):
    live, fresh = charted["snap"][f"live_{step}"], charted["snap"][f"fresh_{step}"]
    assert_same(live, fresh)
    (chart,) = live["charts"]["Dashboard"].values()
    model, _ = charted["steps"][step]
    lines = sum(1 for i in model.instances if i.module == "demo.revenue_line")
    assert len(chart["columns"]) == lines and len(chart["lines"]) == 1
    assert all(stacked for *_, stacked in chart["columns"])


def test_chart_window_follows_the_first_month_shown(charted):
    model, layout = charted["first4"]
    pos = layout.positions()
    values = charted["snap"]["first4"]["sheets"]["Dashboard"]["values"]
    _, row = pos["demo.dashboard#1/total"]
    got = values[row - 1][FIRST_PERIOD_COL - 1:FIRST_PERIOD_COL - 1 + 6]
    assert got == pytest.approx(reference(model)["revenue"][3:9], abs=1e-6)
    _, row = pos["demo.dashboard#1/month"]
    assert values[row - 1][FIRST_PERIOD_COL - 1] == "M4"


# ---------------------------------------------------------------- contents, section covers and navigation

def _contents_entries(layout):
    """Contents as (level, text) pairs: 1 section, 2 sheet, 3 module."""
    out = []
    for r in dict(layout.sheets)["Contents"]:
        if r.id.startswith("contents/section/"):
            out.append((1, r.id.split("/", 2)[2]))
        elif r.id.startswith("contents/sheet/"):
            out.append((2, r.id.split("/", 2)[2]))
        elif r.id.startswith("contents/module/"):
            out.append((3, r.id.split("/", 2)[2]))
    return out


def test_contents_lists_sections_sheets_and_modules(lib):
    layout = assemble(base_model(lib))
    assert [s for s, _ in layout.sheets] == ["Contents", "Model", "Revenue", "Costs", "Working capital", "Funding",
                                             "Statements", "Appendices", "Checks"]
    entries = _contents_entries(layout)
    assert entries[:4] == [(1, "Model"), (2, "Revenue"), (3, "demo.revenue_line#1"), (3, "demo.revenue_line#2")]
    assert (1, "Appendices") in entries and entries[-2:] == [(2, "Checks"), (3, "demo.checks#1")]
    rows = {r.id: r for r in dict(layout.sheets)["Contents"]}
    assert rows["contents/section/Appendices"].cells[2] == 2                       # sections numbered
    assert rows["contents/sheet/Costs"].cells[3] == "b."                           # sheets lettered in their section
    assert rows["contents/sheet/Costs"].cells[4] == '=HYPERLINK("#HL_Sheet_Costs",«S|Costs»)'
    pos = layout.positions()
    for name in ("HL_Home", "HL_Err_Chk", "HL_Sheet_Working_capital", "HL_Toc_demo_facility_1"):
        assert name in layout.names
    assert pos[layout.names["HL_Home"]] == ("Contents", 1) and layout.name_col("HL_Home") == 2
    assert pos[layout.names["HL_Toc_demo_facility_1"]][0] == "Funding"
    cover = {r.id: r for r in dict(layout.sheets)["Model"]}
    assert cover["cover/Model/number"].cells[2] == "Section 1."
    assert "HL_Sheet_Contents" in cover["cover/Model/prev"].cells[2]
    assert "HL_Sheet_Revenue" in cover["cover/Model/next"].cells[2]
    from assemble import frame_cells
    for s, _ in layout.sheets[1:]:                                                 # every sheet links home and to the checks
        cells = frame_cells(layout, s)
        assert cells[(1, 1)].startswith('=HYPERLINK("#HL_Home"') and cells[(2, 1)].startswith('=HYPERLINK("#HL_Err_Chk"')


def test_a_new_section_brings_its_cover_and_the_contents_follow(lib):
    m = base_model(lib)
    before = assemble(m)
    m.insert("demo.dashboard")
    after = assemble(m)
    plan = plan_change(before, after)
    added = [(o["sheet"], o["index"]) for o in plan.ops if o["op"] == "add_sheet"]
    assert added == [("Dashboards", 1), ("Dashboard", 2)]                          # the cover comes with the first sheet
    assert _contents_entries(after)[:3] == [(1, "Dashboards"), (2, "Dashboard"), (3, "demo.dashboard#1")]
    assert any(o["op"] == "insert_rows" and o["sheet"] == "Contents" for o in plan.ops)
    rewired = [r for o in plan.ops if o["op"] == "write" and o["why"] == "rewire" and o["sheet"] == "Contents"
               for r in [o["row"]]]
    assert len(rewired) == 2                                                       # Model and Appendices renumbered
    assert any(o["op"] == "write" and o["sheet"] == "Model" and "HL_Sheet_Dashboard" in str(o["cells"]) for o in plan.ops)
    assert {o["name"] for o in plan.ops if o["op"] == "add_name"} >= {"HL_Sheet_Dashboards", "HL_Sheet_Dashboard",
                                                                      "HL_Toc_demo_dashboard_1"}
    assert "Dashboards: new section cover (Dashboards)." in plan.preview
    m.remove("demo.dashboard#1")
    back = plan_change(after, assemble(m))
    assert sorted(o["sheet"] for o in back.ops if o["op"] == "delete_sheet") == ["Dashboard", "Dashboards"]
    assert {o["name"] for o in back.ops if o["op"] == "delete_name"} >= {"HL_Sheet_Dashboards", "HL_Toc_demo_dashboard_1"}


def test_contents_and_covers_read_right_after_a_live_insert(charted):
    snap = charted["snap"]["live_insert_dashboard"]
    contents = snap["sheets"]["Contents"]["values"]
    shown = [[str(v) for v in row[1:5] if v not in ("", None)] for row in contents]
    assert ["1.0", "Dashboards"] in shown or ["1", "Dashboards"] in shown
    assert ["a.", "Dashboard"] in shown and ["-", "Income summary"] in shown
    assert ["a.", "Revenue"] in shown and ["-", "Revenue line 1"] in shown
    cover = [r[1] for r in snap["sheets"]["Model"]["values"] if r[1] not in ("", None)]
    assert cover[:6] == ["Financial Model", "Assembly proof (demo data)", "Section 2.", "Go to contents", "< Dashboard",
                         "Revenue >"]
    for s, sheet in snap["sheets"].items():                                        # links on every sheet but the contents
        if s != "Contents":
            assert "HL_Home" in sheet["formulas"][0][0] and "HL_Err_Chk" in sheet["formulas"][1][0]
            assert sheet["values"][1][0] == "✓"                                    # no errors in the model
    assert snap["names"]["HL_Sheet_Dashboards"].endswith("$Dashboards.$A$1")
    assert snap["names"]["HL_Toc_demo_dashboard_1"].endswith("$Dashboard.$B$7")
