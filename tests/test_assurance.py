"""Phase 0 proof: model assurance (the HFG additions Kelvin adopted on 7 October 2026).

Three shared pieces in the model's metadata and one engine:
- key outputs, the change log and the before-and-after check on every structural command
  (prototypes/assurance/guard.py);
- input records, the Input register and the Group assumptions sheet (prototypes/assurance/register.py
  and the engine's _assurance_sheets);
- model compare (prototypes/assurance/compare.py), and GST return periods and due dates for the GST
  module's cash timing (prototypes/assurance/gst.py).

The pure Python tests check the layout, what a change can reach and the GST rule. The LibreOffice
tests run a chain of commands on a built model the way the add-in will (read back, plan, read the
key outputs, apply, read again, log), compare each result with the same model built from scratch,
and check the numbers against the independent reference.
"""

import shutil
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
for sub in ("assembly", "assurance"):
    sys.path.insert(0, str(ROOT / "prototypes" / sub))
sys.path.insert(0, str(ROOT))
import compare as CMP  # noqa: E402
import gst  # noqa: E402
import guard  # noqa: E402
import register as REG  # noqa: E402
import release as REL  # noqa: E402
from assemble import (FIRST_PERIOD_COL, GROUP_SHEET, RC, REGISTER_SHEET, TOTAL_COL, Library,  # noqa: E402
                      assemble, open_model, plan_change, read_metadata, write_workbook)
from reference import reference  # noqa: E402


@pytest.fixture(scope="module")
def lib():
    return Library.load()


def _libreoffice():
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")


# ---------------------------------------------------------------- layout, reach and the GST rule

def test_register_lists_every_input_once_and_key_outputs_sit_on_the_contents(lib):
    m = REG.assured_model(lib)
    lay = assemble(m)
    sheets = [s for s, _ in lay.sheets]
    assert sheets.index(GROUP_SHEET) < sheets.index("Revenue") and sheets[-1] == REGISTER_SHEET
    settings = [r.name for s, rows in lay.sheets for r in rows
                if r.kind == "setting" and s not in (REGISTER_SHEET, GROUP_SHEET)]
    reg = [r.id.split("/", 2)[2] for r in lay.sheet_rows()[REGISTER_SHEET] if r.id.startswith("register/in/")]
    assert reg == settings and len(set(reg)) == len(reg)
    m2 = m.copy()
    m2.insert("demo.dashboard", first=1)                 # a view setting is not an assumption
    lay2 = assemble(m2)
    view = next(r.name for r in lay2.sheet_rows()["Dashboard"] if r.kind == "setting")
    assert view in lay2.names and f"register/in/{view}" not in lay2.positions()
    bound = {r.id.split("/", 2)[2] for r in lay.sheet_rows()[REGISTER_SHEET]
             if str(r.cells.get(RC["source"], "")).startswith("Group assumptions")}
    assert bound == {"Fac1_Rate", "Cost1_Inflation"}
    rows = {r.id: r for _, rs in lay.sheets for r in rs}
    assert rows["demo.facility#1/set.rate"].link == "=GA_LendingRate"
    assert rows["demo.cost_line#1/set.inflation"].link == "=(1+GA_Cpi)^(1/12)-1"
    assert [h["label"] for h in lay.headlines] == ["Surplus for the period", "Closing cash", "Lowest cash",
                                                   "Closing debt", "Peak debt"]
    assert all(h["name"] in lay.names for h in lay.headlines)
    # A model without the assurance module is laid out exactly as before.
    from demo import base_model
    plain = assemble(base_model(lib))
    assert GROUP_SHEET not in plain.sheet_rows() and not any(n.startswith(("KO_", "Reg_", "GA_")) for n in plain.names)


def test_what_a_change_can_reach_is_worked_out_from_the_layout(lib):
    m = REG.assured_model(lib)
    old = assemble(m)

    def expected(change, effect="reach"):
        mm = m.copy()
        change(mm)
        new = assemble(mm)
        return guard.expected_outputs(new, plan_change(old, new, "uno"), effect)

    revenue = expected(lambda x: x.insert("demo.revenue_line", base=150, growth=0.02))
    assert revenue == ["Surplus for the period", "Closing cash", "Lowest cash"]          # revenue cannot move debt
    facility = expected(lambda x: x.insert("demo.facility", amount=500, rate=0.08, instalment=25))
    assert facility == ["Surplus for the period", "Closing cash", "Lowest cash", "Closing debt", "Peak debt"]
    assert expected(lambda x: x.insert("demo.dashboard", first=1)) == []                  # a summary reads, never feeds
    assert expected(lambda x: REG.use_set(x, 4)) == ["Surplus for the period", "Closing cash", "Lowest cash"]
    assert expected(lambda x: x.insert("demo.revenue_line", base=150, growth=0.02), effect="none") == []


def test_gst_due_dates_follow_inland_revenue_rule():
    assert gst.due_date(date(2026, 10, 31)) == date(2026, 11, 28)
    assert gst.due_date(date(2026, 11, 30)) == date(2027, 1, 15)
    assert gst.due_date(date(2026, 12, 31)) == date(2027, 1, 28)
    assert gst.due_date(date(2027, 3, 31)) == date(2027, 5, 7)
    assert gst.period_end(date(2026, 4, 30), 2, 1) == date(2026, 5, 31)     # two-monthly, odd months
    assert gst.period_end(date(2026, 4, 30), 2, 0) == date(2026, 4, 30)     # two-monthly, even months
    assert gst.period_end(date(2026, 10, 31), 6, 3) == date(2027, 3, 31)    # six-monthly, March and September
    assert gst.period_end(date(2026, 9, 30), 6, 3) == date(2026, 9, 30)


def test_bound_settings_resolve_for_the_reference(lib):
    m = REG.assured_model(lib)
    eff = REG.effective(m)
    s = {i.uid: i.settings for i in eff.instances}
    assert s["demo.facility#1"]["rate"] == pytest.approx(0.0725)
    assert s["demo.cost_line#1"]["inflation"] == pytest.approx(1.025 ** (1 / 12) - 1)


# ---------------------------------------------------------------- in LibreOffice

def _trim(grid):
    g = [list(r) for r in grid]
    while g and all(v in ("", None) for v in g[-1]):
        g.pop()
    width = max((max((i + 1 for i, v in enumerate(r) if v not in ("", None)), default=0) for r in g), default=0)
    return [(r + [""] * width)[:width] for r in g]


def assert_same(a, b):
    assert list(a["sheets"]) == list(b["sheets"])
    assert a["names"] == b["names"]
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


def _type(src: Path, dst: Path, cells: dict[tuple[str, int, int], object], lib) -> None:
    """Type values into a workbook, as someone would in Excel without the add-in.

    Excel keeps the metadata part when it saves; LibreOffice drops it, so it is put back here."""
    from hfgmodels.verify import _prop, libreoffice
    from live import _url
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL(_url(src), "_blank", 0, (_prop("Hidden", True),))
        try:
            for (sheet, row, col), v in cells.items():
                c = doc.Sheets.getByName(sheet).getCellByPosition(col - 1, row - 1)
                c.setValue(v) if isinstance(v, (int, float)) else c.setString(v)
            doc.calculateAll()
            doc.storeToURL(_url(dst), (_prop("FilterName", "Calc MS Excel 2007 XML"),))
        finally:
            doc.close(True)
    from assemble import write_metadata
    m, _ = open_model(src, lib)
    write_metadata(dst, m, assemble(m))


@pytest.fixture(scope="module")
def chain(tmp_path_factory, lib):
    """A model taken through the commands the add-in runs, each step logged."""
    _libreoffice()
    from live import snapshot
    out = tmp_path_factory.mktemp("assurance")
    m = REG.assured_model(lib)
    write_workbook(assemble(m), out / "a0.xlsx", m)
    run = lambda src, dst, *a, **k: guard.run(out / src, out / dst, lib, *a, **k)  # noqa: E731
    e = {}
    e["summary"] = run("a0.xlsx", "a1.xlsx", "Insert Income summary", lambda x: x.insert("demo.dashboard", first=1),
                       effect="none", when="7 October 2026 09:00")
    e["revenue"] = run("a1.xlsx", "a2.xlsx", "Insert Revenue line 3",
                       lambda x: x.insert("demo.revenue_line", base=150, growth=0.02), when="7 October 2026 09:05")
    e["fault"] = run("a2.xlsx", "f1.xlsx", "Rename Revenue line 3", lambda x: None, effect="none",
                     sabotage=guard.facility_repayments_stop)
    e["fault_reach"] = run("a2.xlsx", "f2.xlsx", "Insert Revenue line 4",
                           lambda x: x.insert("demo.revenue_line", base=80, growth=0.01), sabotage=guard.facility_repayments_stop)
    # Someone fills in the register in Excel: a source, an owner, a date and evidence for Rev2_Growth.
    m2, _ = open_model(out / "a2.xlsx", lib)
    row = assemble(m2).positions()["register/in/Rev2_Growth"][1]
    _type(out / "a2.xlsx", out / "a2t.xlsx", {(REGISTER_SHEET, row, RC["source"]): "Lease schedule indexation clause",
                                              (REGISTER_SHEET, row, RC["owner"]): "Property team",
                                              (REGISTER_SHEET, row, RC["updated"]): REG.serial(date(2026, 9, 1)),
                                              (REGISTER_SHEET, row, RC["evidence"]): "https://example.org/lease"}, lib)
    e["facility"] = run("a2t.xlsx", "a3.xlsx", "Insert Debt facility 2",
                        lambda x: x.insert("demo.facility", amount=500, rate=0.08, instalment=25), when="7 October 2026 09:20")
    e["latest"] = run("a3.xlsx", "a4.xlsx", "Open (version 4 published)", lambda x: REG.note_latest(x, 4), effect="none")
    e["update"] = run("a4.xlsx", "a5.xlsx", "Update to latest group assumptions", lambda x: REG.use_set(x, 4))
    e["local"] = run("a5.xlsx", "a6.xlsx", "Use a local value for Debt facility 1 rate",
                     lambda x: x.unbind("demo.facility#1", "rate", 0.08))
    e["reason"] = run("a6.xlsx", "a7.xlsx", "Give the reason for the local rate",
                      lambda x: REG.record(x, "Fac1_Rate", source="Fixed rate letter from the lender", owner="Kelvin",
                                           updated=date(2026, 10, 5), reason="Fixed rate agreed with the lender to March 2027"),
                      effect="none")
    models = {k: open_model(out / f"{k}.xlsx", lib)[0] for k in ("a3", "a4", "a5", "a6", "a7")}
    fresh = {}
    for k in ("a3", "a5", "a7"):
        write_workbook(assemble(models[k]), out / f"{k}_fresh.xlsx", models[k])
        fresh[k] = snapshot(out / f"{k}_fresh.xlsx")
    snaps = {k: snapshot(out / f"{k}.xlsx") for k in ("a3", "a4", "a5", "a6", "a7")}
    return {"out": out, "e": e, "models": models, "snaps": snaps, "fresh": fresh}


def _value(snap, layout, rid, col=TOTAL_COL):
    s, r = layout.positions()[rid]
    return snap["sheets"][s]["values"][r - 1][col - 1]


def test_a_command_that_should_not_move_the_numbers_moves_none(chain):
    e = chain["e"]["summary"]
    assert e["effect"] == "none" and e["moved"] == [] and e["unexplained"] == []
    assert guard.result_card(e) == ["No key output moved."]


def test_a_change_moves_only_the_outputs_it_can_reach(chain):
    e = chain["e"]["revenue"]
    assert e["moved"] and set(e["moved"]) <= set(e["expected"]) and e["unexplained"] == []
    assert "Closing debt" not in e["moved"]
    e = chain["e"]["facility"]
    assert {"Closing debt", "Peak debt"} <= set(e["moved"]) and e["unexplained"] == []


def test_a_fault_in_a_plan_is_caught(chain):
    e = chain["e"]["fault"]
    assert e["unexplained"] == e["moved"] and "Closing debt" in e["moved"]
    assert guard.result_card(e)[0].startswith("Check this change: ") and "should not move any key output" in guard.result_card(e)[0]
    e = chain["e"]["fault_reach"]   # revenue can move surplus and cash, but not debt
    assert e["unexplained"] == ["Closing debt"]
    assert guard.result_card(e)[0] == "Check this change: Closing debt moved, but nothing this command changed feeds it."


def test_the_change_log_is_kept_in_the_metadata(chain):
    log = read_metadata(chain["out"] / "a7.xlsx")["model"]["assurance"]["log"]
    assert [x["command"] for x in log] == ["Insert Income summary", "Insert Revenue line 3", "Insert Debt facility 2",
                                           "Open (version 4 published)", "Update to latest group assumptions",
                                           "Use a local value for Debt facility 1 rate", "Give the reason for the local rate"]
    assert log[2]["picked_up"] == ["Rev2_Growth"] and log[0]["when"] == "7 October 2026 09:00"
    assert all(set(x["outputs"]) == {"Surplus for the period", "Closing cash", "Lowest cash", "Closing debt", "Peak debt"}
               for x in log)


def test_commands_on_a_built_model_equal_building_it_fresh(chain):
    for k in ("a3", "a5", "a7"):
        assert_same(chain["snaps"][k], chain["fresh"][k])


def test_what_was_typed_into_the_register_survives_a_structural_change(chain):
    m = chain["models"]["a3"]
    assert m.assurance["inputs"]["Rev2_Growth"] == {"source": "Lease schedule indexation clause", "owner": "Property team",
                                                    "updated": REG.serial(date(2026, 9, 1)), "evidence": "https://example.org/lease"}
    lay = assemble(m)
    assert _value(chain["snaps"]["a3"], lay, "register/in/Rev2_Growth", RC["source"]) == "Lease schedule indexation clause"


def test_register_status_and_checks_match_python(chain):
    for k in ("a3", "a6", "a7"):
        m, snap = chain["models"][k], chain["snaps"][k]
        lay = assemble(m)
        want = REG.status(m, lay)
        got = {rid.split("/", 2)[2]: _value(snap, lay, rid, RC["status"])
               for rid in lay.positions() if rid.startswith("register/in/")}
        assert got == want, k
        counts = {w: list(want.values()).count(w) for w in ("No source", "Past review age", "Local, no reason", "Group")}
        for word, rid in (("No source", "nosource"), ("Past review age", "stale"), ("Local, no reason", "local"), ("Group", "group")):
            assert _value(snap, lay, f"register/count/{rid}") == counts[word], (k, word)
        raised = [_value(snap, lay, f"demo.assurance#1/{c}", FIRST_PERIOD_COL) for c in ("chk_source", "chk_stale", "chk_local", "chk_set")]
        assert raised[:3] == [int(counts["No source"] > 0), int(counts["Past review age"] > 0), int(counts["Local, no reason"] > 0)]
    a3 = REG.status(chain["models"]["a3"], assemble(chain["models"]["a3"]))
    assert a3["Rev1_Growth"] == "Past review age" and a3["Rev2_Growth"] == "OK" and a3["Fac1_Rate"] == "Group"
    assert a3["Rev3_Base"] == "No source"                                  # added by a command, not yet filled in
    assert REG.status(chain["models"]["a6"], assemble(chain["models"]["a6"]))["Fac1_Rate"] == "Local, no reason"
    assert REG.status(chain["models"]["a7"], assemble(chain["models"]["a7"]))["Fac1_Rate"] == "Local"


def test_a_newer_set_raises_an_alert_until_the_model_is_updated(chain):
    for k, want in (("a3", 0), ("a4", 1), ("a5", 0)):
        lay = assemble(chain["models"][k])
        assert _value(chain["snaps"][k], lay, "demo.assurance#1/chk_set", FIRST_PERIOD_COL) == want, k
    e = chain["e"]["update"]
    assert e["moved"] and e["unexplained"] == [] and "Closing debt" not in e["moved"]
    assert chain["e"]["latest"]["moved"] == []


def test_numbers_match_the_reference_with_group_assumptions(chain):
    for k in ("a3", "a5", "a7"):
        m, snap = chain["models"][k], chain["snaps"][k]
        lay = assemble(m)
        ref = reference(REG.effective(m))
        for key, series in ref.items():
            s, r = lay.positions()[f"demo.statements#1/{key}"]
            got = snap["sheets"][s]["values"][r - 1][FIRST_PERIOD_COL - 1:FIRST_PERIOD_COL - 1 + m.periods]
            assert got == pytest.approx(series, abs=1e-6), (k, key)
    a5 = assemble(chain["models"]["a5"])
    assert _value(chain["snaps"]["a5"], a5, "ga/item/lending_rate") == pytest.approx(0.0695)
    assert _value(chain["snaps"]["a5"], a5, "demo.facility#1/set.rate") == pytest.approx(0.0695)


def test_model_compare_lists_what_changed_by_module(chain, lib):
    out = chain["out"]
    res = CMP.compare(out / "a1.xlsx", out / "a7.xlsx", lib)
    assert res["modules_added"] == ["Revenue line 3", "Debt facility 2"] and res["modules_removed"] == []
    assert "Financial statements" in res["rows"] and "Debtors" in res["rows"]
    assert res["assumptions"]["a"] == 3 and res["assumptions"]["b"] == 4
    assert {i["item"] for i in res["assumptions"]["items"]} == {"Lending rate, development facilities", "CPI per year"}
    assert [b["name"] for b in res["bindings"]] == ["Fac1_Rate"] and res["bindings"][0]["now"] == "local"
    assert {r["name"] for r in res["records"]} == {"Rev2_Growth", "Fac1_Rate"}
    assert [x["command"] for x in res["log"]][0] == "Insert Revenue line 3" and len(res["log"]) == 6
    changed = {i["name"]: (i["a"], i["b"]) for i in res["inputs"]}
    assert set(changed) == {"Fac1_Rate", "Cost1_Inflation"}
    assert changed["Fac1_Rate"] == (pytest.approx(0.0725), pytest.approx(0.08))
    lines = CMP.summary(res)
    assert lines[0] == "Modules added: Revenue line 3, Debt facility 2."
    same = CMP.compare(out / "a7.xlsx", out / "a7.xlsx", lib)
    assert CMP.summary(same) == ["No differences."]


def test_release_profiles_keep_only_what_the_recipient_needs(chain, lib):
    from live import snapshot
    out = chain["out"]
    full = chain["snaps"]["a7"]
    m = chain["models"]["a7"]
    lay = assemble(m)
    assert REL.plan_release(lay, "auditor") == ["Contents", GROUP_SHEET, "Revenue", "Costs", "Working capital", "Funding",
                                                "Statements", REGISTER_SHEET]
    assert REL.plan_release(lay, "lender") == ["Contents", "Funding", "Statements", "Checks"]
    assert REL.plan_release(lay, "board") == ["Contents", "Dashboard", "Statements"]
    import zipfile
    for profile in REL.PROFILES:
        res = REL.release(out / "a7.xlsx", out / f"{profile}.xlsx", lib, profile, when="7 October 2026")
        snap = snapshot(out / f"{profile}.xlsx")
        assert list(snap["sheets"]) == res["sheets"]
        assert not any(snap["errors"][s] for s in snap["sheets"]), profile
        stmt, ref = _trim(snap["sheets"]["Statements"]["values"]), _trim(full["sheets"]["Statements"]["values"])
        assert len(stmt) == len(ref)
        for ra, rb in zip(stmt[5:], ref[5:]):   # every row from the period row down, as in the full model
            for a, b in zip(ra, rb):
                assert a == (pytest.approx(b, abs=1e-9) if isinstance(b, float) else b), profile
        formulas = [(s, v) for s in snap["sheets"] if s != "Contents"
                    for row in snap["sheets"][s]["formulas"] for v in row if isinstance(v, str) and v.startswith("=")]
        if REL.PROFILES[profile]["values"]:
            assert formulas == [], profile
        else:
            assert any(s == "Statements" for s, _ in formulas)
        assert all(not n.startswith(("HL_", "KO_")) and "#REF" not in c.upper() for n, c in snap["names"].items())
        for s in snap["sheets"]:
            assert snap["sheets"][s]["values"][1][1] in ("", None) or s == "Contents", (profile, s)   # no subtitles
        with zipfile.ZipFile(out / f"{profile}.xlsx") as z:
            assert not any(n.startswith("customXml/") for n in z.namelist())
        toc = [r[1] for r in snap["sheets"]["Contents"]["values"]]
        assert REL.PROFILES[profile]["title"] in toc and "Key outputs" in toc
        assert all(s in toc for s in res["sheets"][1:])
    reg = snapshot(out / "auditor.xlsx")["sheets"][REGISTER_SHEET]["values"]
    heads = next(r for r in reg if "Source" in r)
    assert "Owner" not in heads and "Evidence" not in heads and "Reason for a local value" not in heads
    assert "Status" in heads and "Updated" in heads


@pytest.mark.parametrize("months,cycle,lag", [(1, 0, 0), (2, 0, 1), (2, 1, 0), (6, 3, 1), (6, 9 % 6, 0)])
def test_gst_timing_formulas_match_the_reference(tmp_path, months, cycle, lag):
    _libreoffice()
    from live import snapshot
    start, n = date(2026, 4, 30), 30
    net = [((-1) ** t) * 100.0 + (t % 5) * 37.5 - (400.0 if t in (3, 4, 11) else 0.0) for t in range(n)]
    path = gst.write_workbook(tmp_path / "gst.xlsx", start, n, months, cycle, lag, net)
    vals = snapshot(path)["sheets"]["GST timing"]["values"]
    assert not snapshot(path)["errors"]["GST timing"]
    ref = gst.reference(start, n, months, cycle, lag, net)
    for t, want in enumerate(ref):
        row = vals[gst.FIRST - 1 + t]
        assert REG.as_date(row[2]) == want["end"]
        assert REG.as_date(row[3]) == want["period_end"], (t, months, cycle)
        assert REG.as_date(row[4]) == want["due"]
        assert row[5] == want["due_month"]
        assert row[9] == pytest.approx(want["cash"], abs=1e-9), (t, months, cycle, lag)
