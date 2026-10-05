"""Recalculate a built workbook in LibreOffice (headless) and read results.

Built workbooks carry formulas without cached values. Excel recalculates on
open; for tests and for cached values we use LibreOffice through UNO.
"""

from __future__ import annotations

import os
import shutil
import socket
import subprocess
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path


def _free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


@contextmanager
def libreoffice():
    """Start a headless LibreOffice and yield a UNO desktop."""
    import uno  # noqa: F401  (provided by the LibreOffice Python bridge)
    from com.sun.star.connection import NoConnectException

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) is not installed")
    port = _free_port()
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    proc = subprocess.Popen(
        [soffice, "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
         f"-env:UserInstallation=file://{profile}",
         f"--accept=socket,host=127.0.0.1,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = None
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except NoConnectException:
            time.sleep(0.5)
    if ctx is None:
        proc.kill()
        raise RuntimeError("could not connect to LibreOffice")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    try:
        yield desktop
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=20)
        except Exception:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


def _prop(name, value):
    import uno
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


class CalculatedWorkbook:
    """Values of selected sheets after a full recalculation."""

    def __init__(self, values: dict[str, list[list]], errors: dict[str, list[str]]):
        self.values = values  # sheet -> 2D list (row 1 at index 0)
        self.errors = errors  # sheet -> list of A1 ranges with formula errors

    def get(self, sheet: str, row: int, col: int):
        grid = self.values[sheet]
        if row - 1 >= len(grid) or col - 1 >= len(grid[row - 1]):
            return None
        return grid[row - 1][col - 1]


def recalculate(path: str | Path, sheets: list[str]) -> CalculatedWorkbook:
    path = Path(path).resolve()
    with libreoffice() as desktop:
        url = "file://" + str(path)
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (_prop("Hidden", True), _prop("ReadOnly", True)))
        try:
            doc.calculateAll()
            values, errors = {}, {}
            for name in sheets:
                sh = doc.Sheets.getByName(name)
                cur = sh.createCursor()
                cur.gotoEndOfUsedArea(False)
                last_col = cur.RangeAddress.EndColumn
                last_row = cur.RangeAddress.EndRow
                rng = sh.getCellRangeByPosition(0, 0, last_col, last_row)
                values[name] = [list(r) for r in rng.getDataArray()]
                err = sh.queryFormulaCells(4)  # com.sun.star.sheet.FormulaResult.ERROR
                errors[name] = [a for a in err.getRangeAddressesAsString().split(";") if a]
            return CalculatedWorkbook(values, errors)
        finally:
            doc.close(True)
