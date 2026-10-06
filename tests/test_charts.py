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
    assert "pattFill" in z_chart and 'prst="wdUpDiag"' in z_chart      # forecast bars hatched
    assert "barChart" in z_chart and "lineChart" in z_chart            # bars and lines in one chart
    assert 'prstDash val="dash"' in z_chart                            # forecast cumulative dashed
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
            assert zs.getCellRangeByName("G18").getValue() == pytest.approx(82)       # full year AC then FC, revenue
            assert zs.getCellRangeByName("F20").getValue() == pytest.approx(6)        # YTD variance to budget
            assert ck.getCellRangeByName("H10").getValue() == 0
            data.getCellRangeByName("H4").setValue(3)                               # the combo box picks Gross profit
            doc.calculateAll()
            assert zs.getCellRangeByName("D4").getString() == "Gross profit"
            assert zs.getCellRangeByName("G18").getValue() == pytest.approx(data.getCellRangeByName("H22").getValue())
            assert ck.getCellRangeByName("H10").getValue() == 0
            ib.getCellRangeByName("E4").setString("Prior year")                     # the in-cell drop-down
            doc.calculateAll()
            assert ib.getCellRangeByName("I7").getValue() == pytest.approx(1085 - 920)
            assert ib.getCellRangeByName("J7").getValue() == pytest.approx(165)     # revenue up: good
            assert ck.getCellRangeByName("H10").getValue() == 0
        finally:
            doc.close(True)
