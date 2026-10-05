"""Load and validate a development model recipe (YAML)."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import yaml

PROFILES = ("Lump sum", "Flat", "S-curve")
FUNDING_METHODS = ("LTC ceiling", "Equity first")
LIMIT_BASES = ("Override", "LTC on budget", "LVR on As Complete", "LVR on BTR")
GST_TREATMENTS = ("Standard-rated", "Zero-rated")


def yes_no(v) -> str:
    if isinstance(v, bool):
        return "Yes" if v else "No"
    s = str(v).strip().capitalize()
    if s not in ("Yes", "No"):
        raise ValueError(f"expected Yes or No, got {v!r}")
    return s


def as_date(v) -> dt.date:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v))


@dataclass
class Line:
    key: str
    label: str
    budget: float
    profile: str
    start: dt.date
    months: int
    gst: str
    counterparty: str = ""


@dataclass
class Site:
    code: str
    name: str
    entity: str
    include: str
    gst_registered: str
    units: float
    opening_wip: float
    opening_debt: float
    alpha: float
    beta: float
    lines: list[Line]
    sale_date: dt.date
    sale_price: float
    sale_gst: str
    via_spv: str
    btr_value: float
    method: str
    equity_commitment: float
    limit_basis: str
    limit_override: float
    max_ltc: float
    max_lvr: float
    debt_start: dt.date
    land_loan: float
    rate: float
    line_fee: float
    estab_fee: float
    capitalise: str
    release_equity: str
    cap_borrowing: str


@dataclass
class Recipe:
    name: str
    periods: int
    last_actual: int
    insert_after: str
    checks_after: str
    template_env: str
    gst_rate: float
    gst_lag: int
    fund_lvr: float
    line_items: list[tuple[str, str]]
    sites: list[Site] = field(default_factory=list)
    theme: dict = field(default_factory=dict)
    start_date: dt.date = dt.date(2025, 4, 1)  # template Ts_Start_Date

    def month_index(self, d: dt.date) -> int:
        """Timeline index of the month containing d (1 = first model month)."""
        return (d.year - self.start_date.year) * 12 + d.month - self.start_date.month + 1


def _check(cond, msg):
    if not cond:
        raise ValueError(msg)


def load(path: str | Path) -> Recipe:
    raw = yaml.safe_load(Path(path).read_text())
    m, g = raw["model"], raw["global"]
    items = [(x["key"], x["label"]) for x in raw["line_items"]]
    r = Recipe(name=m["name"], periods=int(m["periods"]), last_actual=int(m["last_actual_month"]),
               insert_after=m["insert_after"], checks_after=m.get("checks_after", "Checks"),
               template_env=m.get("template_env", "HFG_TEMPLATE"),
               gst_rate=float(g["gst_rate"]), gst_lag=int(g["gst_lag_months"]), fund_lvr=float(g["fund_lvr"]),
               line_items=items, theme={k: str(v) for k, v in (m.get("theme") or {}).items()})
    _check(1 <= r.last_actual < r.periods, "last_actual_month must be between 1 and periods - 1")
    codes = set()
    for s in raw["sites"]:
        code = s["code"]
        _check(code not in codes, f"duplicate site code {code}")
        codes.add(code)
        lines = []
        for key, label in items:
            x = s["lines"][key]
            prof = x["profile"]
            _check(prof in PROFILES, f"{code}.{key}: profile must be one of {PROFILES}")
            lines.append(Line(key=key, label=label, budget=float(x["budget"]), profile=prof,
                              start=as_date(x["start"]), months=max(1, int(x["months"])), gst=yes_no(x["gst"]),
                              counterparty=str(x.get("counterparty", "") or "")))
        f, sale, acc = s["funding"], s["sale"], s["accounting"]
        _check(f["method"] in FUNDING_METHODS, f"{code}: funding method must be one of {FUNDING_METHODS}")
        _check(f["limit_basis"] in LIMIT_BASES, f"{code}: limit basis must be one of {LIMIT_BASES}")
        _check(sale["gst_treatment"] in GST_TREATMENTS, f"{code}: GST treatment must be one of {GST_TREATMENTS}")
        r.sites.append(Site(
            code=code, name=s["name"], entity=s["entity"], include=yes_no(s["include"]),
            gst_registered=yes_no(s["gst_registered"]), units=float(s["units"]),
            opening_wip=float(s["opening"]["wip"]), opening_debt=float(s["opening"]["debt"]),
            alpha=float(s["scurve"]["alpha"]), beta=float(s["scurve"]["beta"]), lines=lines,
            sale_date=as_date(sale["date"]), sale_price=float(sale["price"]), sale_gst=sale["gst_treatment"],
            via_spv=yes_no(sale["via_spv"]), btr_value=float(sale["btr_value"]),
            method=f["method"], equity_commitment=float(f["equity_commitment"]), limit_basis=f["limit_basis"],
            limit_override=float(f["limit_override"]), max_ltc=float(f["max_ltc"]), max_lvr=float(f["max_lvr"]),
            debt_start=as_date(f["debt_start"]), land_loan=float(f["land_loan"]), rate=float(f["rate"]),
            line_fee=float(f["line_fee"]), estab_fee=float(f["estab_fee"]), capitalise=yes_no(f["capitalise"]),
            release_equity=yes_no(f["release_equity"]), cap_borrowing=yes_no(acc["capitalise_borrowing_costs"]),
        ))
    return r
