from fractions import Fraction
import math

import pandas as pd
import pytest

from support import load_helpers


@pytest.mark.parametrize("rate", [0, 10, 15, 20, 30, 50, 99.5])
@pytest.mark.parametrize("cost", [0, 1, 20, 250, 10000, 18055, 10000.25])
def test_gross_margin_meets_independent_fraction_oracle(rate, cost):
    h = load_helpers()
    exact = Fraction(str(cost)) / (1 - Fraction(str(rate)) / 100)
    expected = math.ceil(exact / 10) * 10
    actual = h.proposal_price(cost, rate, "gross_margin")
    assert actual == expected
    assert 0 <= Fraction(str(actual)) - exact < 10
    if actual and cost:
        assert (Fraction(str(actual)) - Fraction(str(cost))) / Fraction(str(actual)) >= Fraction(str(rate)) / 100


@pytest.mark.parametrize("rate", [0, 8.5, 10, 15, 20, 30, 50])
def test_legacy_scalar_and_series_rounding_are_unchanged(rate):
    h = load_helpers()
    costs = pd.Series([0., 1., 250., 650., 10000., 18055., 10000.25], index=list("abcdefg"))
    pd.testing.assert_series_equal(h.proposal_price(costs, rate), (costs * (1 + rate / 100)).round(-1))
    for cost in costs:
        assert h.proposal_price(cost, rate) == round(cost * (1 + rate / 100), -1)
    pd.testing.assert_series_equal(h.proposal_price(costs, rate, "gross_margin"),
                                   costs.map(lambda c: h.proposal_price(c, rate, "gross_margin")))


@pytest.mark.parametrize("cost,rate,basis", [(10000, 100, "gross_margin"), (10000, 101, "gross_margin"),
    (10000, -1, "gross_margin"), (10000, float("nan"), "gross_margin"),
    (10000, float("inf"), "gross_margin"), (-1, 20, "gross_margin"),
    (float("nan"), 20, "gross_margin"), (float("inf"), 20, "gross_margin"), (1, 20, "unknown")])
def test_invalid_automatic_pricing_rejected(cost, rate, basis):
    with pytest.raises(ValueError):
        load_helpers().proposal_price(cost, rate, basis)


def test_existing_session_is_not_reinitialized():
    h = load_helpers()
    state = h.st.session_state
    del state["q_pricing_basis"]
    state.q_express_margin = 20.0
    state.q_warehouse_df.loc[0, "제안단가"] = 12345
    before = state.q_warehouse_df.copy(deep=True)
    h.init_quote_state()
    assert state.q_pricing_basis == "markup"
    assert state.q_express_margin == 20
    pd.testing.assert_frame_equal(state.q_warehouse_df, before)
