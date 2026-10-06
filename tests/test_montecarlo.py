"""Phase 0 proof: a native, seeded Monte Carlo over a one-way cash flow model.

The workbook is recalculated in LibreOffice (which runs the Excel data table as a
table operation) and every trial is compared with the Python reference.
"""

import shutil
import sys
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "prototypes" / "montecarlo"))

from simulation import OUTPUTS, Inputs, cash_model, hdr, simulate  # noqa: E402

N = 200


def test_generator_is_a_pure_function_of_its_counters():
    assert hdr(1, 101, 1) == hdr(1, 101, 1)
    u = np.array([hdr(t, 101, 1) for t in range(1, 5001)])
    assert 0 < u.min() and u.max() < 1
    assert abs(u.mean() - 0.5) < 0.01
    assert hdr(1, 101, 1) != hdr(1, 101, 2)  # the seed changes the stream


def test_latin_hypercube_puts_one_draw_in_each_stratum():
    sim = simulate(Inputs(correlation={}), N)
    from scipy import stats
    for j, d in enumerate(Inputs().drivers):
        if d.dist in ("PERT", "Normal", "Triangular"):
            lo, mode, hi = d.p1, d.p2, d.p3
            if d.dist == "PERT":
                a, b = 1 + 4 * (mode - lo) / (hi - lo), 1 + 4 * (hi - mode) / (hi - lo)
                u = stats.beta.cdf(sim["draws"][:, j], a, b, loc=lo, scale=hi - lo)
            elif d.dist == "Normal":
                u = stats.norm.cdf(sim["draws"][:, j], d.p1, d.p2)
            else:
                c = (mode - lo) / (hi - lo)
                x = sim["draws"][:, j]
                u = np.where(x < mode, (x - lo) ** 2 / ((hi - lo) * (mode - lo)), 1 - (hi - x) ** 2 / ((hi - lo) * (hi - mode)))
                assert c > 0
            assert sorted(np.floor(u * N).astype(int)) == list(range(N))


def test_correlation_reaches_the_target_roughly():
    d = simulate(Inputs(), 2000)["draws"]
    assert np.corrcoef(d[:, 0], d[:, 3])[0, 1] == pytest.approx(-0.4, abs=0.06)


def test_base_case_events():
    inp = Inputs()
    base = cash_model(inp, {x.name: x.base for x in inp.drivers})
    assert base["breach"] == 0 and base["peak_debt"] > 0


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")
    from simulation import build
    path = tmp_path_factory.mktemp("mc") / "mc.xlsx"
    info = build(Inputs(), N, path)
    sys.path.insert(0, str(ROOT))
    from hfgmodels.verify import recalculate
    calc = recalculate(path, ["Run", "Checks", "Cash", "Stats", "Sensitivity", "Control"])
    return {"path": path, "info": info, "calc": calc}


def test_every_trial_matches_the_reference(built):
    table = built["info"]["reference"]["table"]
    grid = built["calc"].values["Run"]
    got = np.array([row[2:2 + table.shape[1]] for row in grid[6:6 + N]], dtype=float)
    assert got.shape == table.shape
    assert np.allclose(got, table, rtol=1e-7, atol=1e-6)


def test_workbook_checks_pass(built):
    calc = built["calc"]
    assert not any(calc.errors.values()), calc.errors
    checks = calc.values["Checks"]
    labels = {row[2]: row[8] for row in checks if len(row) > 8 and isinstance(row[2], str)}
    assert labels["Error checks failing"] == 0
    assert labels["Results differ from the build's reference"] == 0
    assert labels["Simulation not run (press F9)"] == 0


def test_rank_sensitivity_matches_the_reference(built):
    rho = built["info"]["reference"]["rho"]
    sen = built["calc"].values["Sensitivity"]
    got = [sen[6 + j][2] for j in range(len(rho))]
    assert np.allclose(got, rho, atol=1e-9)


def test_inspect_trial_replays_that_trial_through_the_visible_model(built):
    from hfgmodels.verify import _prop, libreoffice
    rows = built["info"]["rows"]
    table = built["info"]["reference"]["table"]
    trial = 37
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL("file://" + str(Path(built["path"]).resolve()), "_blank", 0,
                                           (_prop("Hidden", True),))
        try:
            doc.Sheets.getByName("Control").getCellRangeByName("I8").setValue(trial)
            doc.calculateAll()
            cash = doc.Sheets.getByName("Cash")
            for i, (key, _, _) in enumerate(OUTPUTS):
                v = cash.getCellRangeByName(f"I{rows['cash_results'][key]}").getValue()
                assert v == pytest.approx(table[trial - 1, i], rel=1e-7, abs=1e-6), key
        finally:
            doc.close(True)
