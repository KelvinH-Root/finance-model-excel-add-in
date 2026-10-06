"""Input records, the Input register and group assumptions.

The engine (prototypes/assembly/assemble.py) writes two sheets when the Model assurance module is
in a model: Group assumptions (the set the model uses, item by item, with GA_ names) and the
Input register (every named input with its value, source, owner, date updated, evidence, age,
the group assumption it draws on and a status). This module holds what sits around them:

- loading a published set and putting it in a model (a library file in Phase 1, Home Hub in Phase 2);
- noting the latest version published, which raises an alert until the model is updated;
- reading back what people typed into the register, so a structural change never loses it
  (the add-in does this on open and before every structural command, as it reconciles metadata);
- the settings a model resolves to, for the independent reference calculation.
"""

from __future__ import annotations

import copy
import sys
from datetime import date, datetime
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "assembly"))
from assemble import (RC, REGISTER_SHEET, TOTAL_COL, Layout, Model)  # noqa: E402

SETS = HERE / "sets"
EPOCH = date(1899, 12, 30)


def serial(d: date | datetime | int | float) -> int:
    """An Excel date number."""
    if isinstance(d, (int, float)):
        return int(d)
    if isinstance(d, datetime):
        d = d.date()
    return (d - EPOCH).days


def as_date(n: int | float) -> date:
    from datetime import timedelta
    return EPOCH + timedelta(days=int(n))


def load_set(version: int) -> dict:
    """A published group assumptions set, with its date as an Excel date number."""
    s = yaml.safe_load((SETS / f"hfg-group-v{version}.yaml").read_text())
    s["published"] = serial(s["published"])
    return s


def latest_version() -> int:
    return max(int(p.stem.rsplit("v", 1)[1]) for p in SETS.glob("hfg-group-v*.yaml"))


def use_set(model: Model, version: int) -> None:
    """Put a set in the model (Update to latest set, or the set a new model starts on)."""
    model.assurance["set"] = load_set(version)
    model.assurance["latest"] = max(version, model.assurance.get("latest", 0))


def note_latest(model: Model, version: int | None = None) -> None:
    """What the add-in does on open: record the latest version published, so the check can compare."""
    model.assurance["latest"] = version if version is not None else latest_version()


def record(model: Model, name: str, **fields) -> None:
    """Set an input's record (source, owner, updated, evidence, reason), as the Inputs pane does."""
    rec = model.assurance.setdefault("inputs", {}).setdefault(name, {})
    for k, v in fields.items():
        rec[k] = serial(v) if k == "updated" and isinstance(v, (date, datetime)) else v


def read_back(path: Path, model: Model, layout: Layout) -> dict:
    """Copy what people typed into the register (and its date and review age) into the model's records.

    Columns the engine owns (the link, value, age, group assumption and status, and the source,
    owner and date of an input drawn from the group set) are left alone. Returns the records
    that changed, so the pane can say what it picked up."""
    from openpyxl import load_workbook

    if REGISTER_SHEET not in layout.sheet_rows():
        return {}
    ws = load_workbook(path)[REGISTER_SHEET]
    pos = layout.positions()
    changed: dict[str, dict] = {}
    for key, rid in (("register_date", "register/date"), ("review_age", "register/age")):
        v = ws.cell(pos[rid][1], TOTAL_COL).value
        if isinstance(v, (datetime, date)):
            v = serial(v)
        if v not in (None, "") and v != model.assurance.get(key):
            model.assurance[key] = v
            changed[key] = {"value": v}
    recs = model.assurance.setdefault("inputs", {})
    for r in layout.sheet_rows()[REGISTER_SHEET]:
        if not r.id.startswith("register/in/"):
            continue
        name = r.id.split("/", 2)[2]
        if str(r.cells.get(RC["source"]) or "").startswith("Group assumptions"):
            continue  # drawn from the group set: the engine owns its source, owner and date
        row = pos[r.id][1]
        got = {}
        for k in ("source", "owner", "updated", "evidence", "reason"):
            v = ws.cell(row, RC[k]).value
            if isinstance(v, (datetime, date)):
                v = serial(v)
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            if v not in (None, ""):
                got[k] = v
        if got != {k: v for k, v in recs.get(name, {}).items() if v not in (None, "")}:
            recs[name] = got
            changed[name] = got
    return changed


def effective(model: Model) -> Model:
    """The model with every bound setting replaced by the number it resolves to (for the reference)."""
    m = model.copy()
    items = {i["key"]: i["value"] for i in (m.assurance.get("set") or {}).get("items", [])}
    for inst in m.instances:
        specs = {s["key"]: s for s in m.lib.modules[inst.module].get("settings", [])}
        for k, v in list(inst.settings.items()):
            if isinstance(v, dict) and "group" in v:
                tpl = specs[k]["group"].get("formula", "={item}")
                expr = tpl.lstrip("=").replace("{item}", repr(items[v["group"]])).replace("^", "**")
                inst.settings[k] = eval(expr, {"__builtins__": {}}, {})  # digits and arithmetic only
    return m


def status(model: Model, layout: Layout) -> dict[str, str]:
    """The status each register row should show, worked out in Python (the test compares the sheet with it)."""
    a = model.assurance
    as_at = as_date(a.get("register_date"))
    max_age = a.get("review_age")
    out = {}
    for r in layout.sheet_rows().get(REGISTER_SHEET, []):
        if not r.id.startswith("register/in/"):
            continue
        name = r.id.split("/", 2)[2]
        group = r.cells.get(RC["group"]) or ""
        source = r.cells.get(RC["source"]) or ""
        upd = r.cells.get(RC["updated"])
        if upd == "=GA_Published":
            upd = (a.get("set") or {}).get("published")
        rec = a.get("inputs", {}).get(name, {})
        if group.startswith("Local") and not rec.get("reason"):
            out[name] = "Local, no reason"
        elif source in ("", None):
            out[name] = "No source"
        elif upd not in (None, "") and (as_at.year - as_date(upd).year) * 12 + as_at.month - as_date(upd).month > max_age:
            out[name] = "Past review age"
        elif not group:
            out[name] = "OK"
        else:
            out[name] = "Local" if group.startswith("Local") else "Group"
    return out


def assured_model(lib, version: int = 3) -> Model:
    """The assembly demo with the Model assurance module, two inputs drawn from the group set and a
    few records filled in, as a recipe would start it."""
    sys.path.insert(0, str(HERE.parent / "assembly"))
    from demo import base_model

    m = base_model(lib)
    m.insert("demo.assurance")
    m.assurance = {"register_date": serial(date(2026, 10, 1)), "review_age": 12}
    use_set(m, version)
    m.bind("demo.facility#1", "rate")
    m.bind("demo.cost_line#1", "inflation")
    record(m, "Rev1_Base", source="Board-approved budget FY2027", owner="Kelvin", updated=date(2026, 6, 30),
           evidence="https://example.org/budget-fy2027")
    record(m, "Rev1_Growth", source="Board-approved budget FY2027", owner="Kelvin", updated=date(2025, 6, 30))
    record(m, "Rev2_Base", source="Signed lease schedule", owner="Property team", updated=date(2026, 9, 15))
    record(m, "Cost1_Amount", source="Operating budget", owner="Kelvin", updated=date(2026, 8, 1))
    record(m, "Fac1_Amount", source="Facility agreement", owner="Kelvin", updated=date(2026, 3, 1),
           evidence="https://example.org/facility-agreement")
    record(m, "Fac1_Instalment", source="Facility agreement", owner="Kelvin", updated=date(2026, 3, 1))
    return m
