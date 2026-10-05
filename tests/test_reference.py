"""Properties the development example's reference calculation must always satisfy."""

import numpy as np
import pytest

from examples.development.reference import irr, shares


@pytest.mark.parametrize("profile", ["S-curve", "Flat", "Lump sum"])
def test_profile_shares_sum_to_one(profile):
    s = shares(profile, start=5, months=14, alpha=2.0, beta=2.5, n=40)
    assert s.sum() == pytest.approx(1.0)
    assert s[:4].sum() == 0  # nothing before the start month


def test_symmetric_scurve_is_symmetric():
    s = shares("S-curve", start=1, months=10, alpha=2.0, beta=2.0, n=10)
    assert np.allclose(s, s[::-1])


def test_irr_known_answer():
    flows = np.array([-100.0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 110.0])
    r = irr(flows)
    assert (1 + r) ** 11 == pytest.approx(1.1)


def _sites(recipe):
    return [s.code for s in recipe.sites]


def test_balance_sheet_balances(recipe, reference):
    for k in _sites(recipe):
        assert np.abs(reference[f"{k}.fs.bs_check"]).max() < 1e-6


def test_budget_equals_opening_wip_plus_forecast(recipe, reference):
    for s in recipe.sites:
        k = s.code
        total = s.opening_wip + reference[f"{k}.cost.total"].sum()
        assert total == pytest.approx(reference[f"{k}.d.total_budget"])


def test_facility_within_limit_and_never_negative(recipe, reference):
    for k in _sites(recipe):
        close = reference[f"{k}.fund.close"]
        assert close.max() <= reference[f"{k}.d.limit"] + 1e-6
        assert close.min() >= -1e-6


def test_facility_repaid_and_equity_returns_the_surplus(recipe, reference):
    for k in _sites(recipe):
        assert reference[f"{k}.fund.close"][-1] == pytest.approx(0, abs=1e-6)
        # once everything is settled, net equity invested equals minus the surplus
        assert reference[f"{k}.fund.eq_bal"][-1] == pytest.approx(-reference[f"{k}.ret.profit"], abs=1e-4)


def test_gst_settles_in_full(recipe, reference):
    for k in _sites(recipe):
        assert reference[f"{k}.gst.net"].sum() == pytest.approx(reference[f"{k}.gst.cash"].sum(), abs=1e-6)
        assert reference[f"{k}.gst.balance"][-1] == pytest.approx(0, abs=1e-6)


def test_loan_to_cost_within_maximum(recipe, reference):
    for s in recipe.sites:
        assert reference[f"{s.code}.fund.ltc"].max() <= s.max_ltc + 1e-9


def test_cash_never_negative(recipe, reference):
    for k in _sites(recipe):
        assert reference[f"{k}.fund.cash"].min() >= -1e-9
