"""Phase 0 proof: native report charts and controls written by the package writer."""

import shutil
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "prototypes" / "charts"))
sys.path.insert(0, str(ROOT))

from build import build  # noqa: E402


@pytest.fixture(scope="module")
def demo(tmp_path_factory):
    return build(tmp_path_factory.mktemp("charts") / "charts_demo.xlsx")


def test_charts_are_native_and_styled(demo):
    with zipfile.ZipFile(demo) as z:
        charts = [z.read(n).decode() for n in sorted(z.namelist()) if n.startswith("xl/charts/chart")]
    assert len(charts) == 5
    z_chart = charts[0]
    assert "pattFill" in z_chart and 'prst="upDiag"' in z_chart        # forecast bars hatched, as in the template
    assert z_chart.count("<barChart>") == 2                            # budget bars behind, actual and forecast in front
    assert "lineChart" in z_chart and "scatterChart" in z_chart        # lines, and helpers placed between months
    assert 'prstDash val="dash"' in z_chart                            # forecast cumulative dashed
    assert "<legend>" not in z_chart                                   # labelled lines instead of a legend
    assert "'Z chart'!$D$5" in z_chart                                 # title follows the chosen line
    assert z_chart.count('<showSerName val="1"/>') >= 11               # MAT, AC+FC, BUD, AC, FC, deltas, variances
    assert 'cap="flat"' in z_chart                                     # variance boxes
    assert all("dispNaAsBlank" in c for c in charts)                   # #N/A draws nothing in Excel
    ibcs = charts[3]
    assert "pattFill" in ibcs and "<a:noFill/>" in ibcs                # forecast hatched, plan outlined


def test_form_control_is_written_as_excel_writes_it(demo):
    with zipfile.ZipFile(demo) as z:
        names = z.namelist()
        ctrl = z.read("xl/ctrlProps/ctrlProp1.xml").decode()
        sheet = next(z.read(n).decode() for n in names if n.startswith("xl/worksheets/sheet") and "rIdCtlP" in z.read(n).decode())
        ct = z.read("[Content_Types].xml").decode()
    assert 'objectType="Drop"' in ctrl and 'fmlaLink="DD_Chart_Line"' in ctrl and 'fmlaRange="LU_Chart_Lines"' in ctrl
    assert "xl/drawings/vmlDrawing1.vml" in names
    assert sheet.index("<legacyDrawing") < sheet.index("<controls>") < sheet.index("</worksheet>")
    assert "controlproperties+xml" in ct and 'Extension="vml"' in ct


def _libreoffice():
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    try:
        import uno  # noqa: F401
    except ImportError:
        pytest.skip("LibreOffice Python bridge not available")


def test_selections_drive_the_charts_and_checks_stay_clear(demo):
    _libreoffice()
    from hfgmodels.verify import _prop, libreoffice
    with libreoffice() as desktop:
        doc = desktop.loadComponentFromURL("file://" + str(Path(demo).resolve()), "_blank", 0, (_prop("Hidden", True),))
        try:
            data, zs, ib, ck = (doc.Sheets.getByName(n) for n in ("Data", "Z chart", "IBCS", "Checks"))
            doc.calculateAll()
            z = lambda ref: zs.getCellRangeByName(ref)                              # noqa: E731
            hidden = lambda ref: z(ref).getError() != 0                             # noqa: E731  #N/A hides a helper
            # Revenue: the template's Z-Chart numbers
            assert z("K19").getValue() == pytest.approx(82)                         # full year AC then FC
            assert z("I19").getValue() == pytest.approx(76)                         # full year budget
            assert z("J19").getValue() == pytest.approx(82)                         # MAT meets the year at March
            assert [z(f"I{r}").getValue() for r in (8, 11, 13, 14, 17)] == [6, 27, 38, 46, 66]
            assert z("F26").getValue() == pytest.approx(6)                          # YTD variance to budget
            assert z("F29").getValue() == pytest.approx(6)                          # full year variance
            assert z("M52").getString() == "+6"
            assert z("C38").getValue() == pytest.approx(6.5)                        # divider between Sep and Oct
            assert not hidden("H45") and hidden("H46")                              # YTD box green, not red
            assert ck.getCellRangeByName("H11").getValue() == 0
            data.getCellRangeByName("H4").setValue(2)                               # the combo box picks Cost of sales
            doc.calculateAll()
            assert z("J4").getValue() == 1                                          # a cost line: over budget is bad
            assert hidden("H48") and not hidden("H49")                              # FY box red
            data.getCellRangeByName("H5").setValue(12)                              # a full year of actuals
            doc.calculateAll()
            assert hidden("C38") and hidden("H38") and hidden("C45")                # no divider and no YTD box
            data.getCellRangeByName("H5").setValue(6)
            data.getCellRangeByName("H4").setValue(3)                               # Gross profit
            doc.calculateAll()
            assert z("D4").getString() == "Gross profit"
            assert z("K19").getValue() == pytest.approx(data.getCellRangeByName("H22").getValue())
            assert ck.getCellRangeByName("H11").getValue() == 0
            ib.getCellRangeByName("E4").setString("Prior year")                     # the in-cell drop-down
            doc.calculateAll()
            assert ib.getCellRangeByName("I7").getValue() == pytest.approx(1085 - 920)
            assert ib.getCellRangeByName("J7").getValue() == pytest.approx(165)     # revenue up: good
            assert ck.getCellRangeByName("H11").getValue() == 0
        finally:
            doc.close(True)
