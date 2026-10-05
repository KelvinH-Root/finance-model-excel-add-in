"""Compare a recalculated workbook with the reference calculation."""

from __future__ import annotations

import numpy as np

from .. import frame
from ..layout import Model
from ..verify import CalculatedWorkbook

# reference keys whose workbook row lives under a different key
ALIASES = {
    ".d.total_budget": ".d.total_budget",
}

RET_MAP = {
    "ret.revenue": "revenue", "ret.dev_cost": "dev_cost", "ret.finance": "finance", "ret.profit": "profit",
    "ret.margin": "margin", "ret.poc": "poc", "ret.peak_debt": "peak_debt", "ret.peak_equity": "peak_equity",
    "ret.eq_irr": "eq_irr", "ret.proj_irr": "proj_irr", "ret.multiple": "multiple",
    "ret.hhlp_margin": "hhlp_margin", "ret.group_margin": "group_margin",
}


def _num(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return float("nan")


def compare(model: Model, cw: CalculatedWorkbook, ref: dict, sites: list, tol: float = 0.01) -> list[str]:
    """Return a list of mismatch descriptions (empty when everything agrees)."""
    problems = []
    n = model.n
    for key, expected in ref.items():
        site = key.split(".")[0]
        if key in model.index:
            sheet, row = model.index[key]
            if isinstance(expected, np.ndarray):
                got = np.array([_num(cw.get(sheet, row, frame.ts_col(t))) for t in range(1, n + 1)])
                diff = np.abs(got - expected)
                scale = np.maximum(1.0, np.abs(expected))
                bad = np.nonzero(diff > tol * np.maximum(1.0, scale * 1e-6))[0]
                if len(bad):
                    i = bad[0]
                    problems.append(f"{key} ({sheet} row {row}) period {i + 1}: workbook {got[i]:.4f}, reference {expected[i]:.4f}"
                                    f" ({len(bad)} periods differ)")
            else:
                got = _num(cw.get(sheet, row, frame.UNIT_COL))
                if abs(got - float(expected)) > tol * max(1.0, abs(float(expected)) * 1e-6):
                    problems.append(f"{key} ({sheet}!H{row}): workbook {got}, reference {expected}")
        elif ".ret." in key:
            short = key.split(".ret.")[1]
            rk = f"ret.{short}"
            if rk not in model.index:
                continue
            sheet, row = model.index[rk]
            col = 8 + [s.code for s in sites].index(site)
            got = _num(cw.get(sheet, row, col))
            exp = float(expected)
            t = 1e-6 if short in ("margin", "poc", "eq_irr", "proj_irr", "multiple") else tol
            if abs(got - exp) > max(t, abs(exp) * 1e-7):
                problems.append(f"{key} ({sheet} row {row}): workbook {got}, reference {exp}")
    return problems
