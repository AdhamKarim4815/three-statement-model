from __future__ import annotations

from dataclasses import replace

import pytest

from three_statement import SCENARIOS, northwind_inputs
from three_statement.inputs import canonical_scenario


def test_default_inputs_validate(inputs):
    inputs.validate()


def test_timeline_is_three_actual_and_five_forecast_years(inputs):
    assert inputs.n_hist == 3
    assert inputs.n_forecast == 5
    assert inputs.years == [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]


def test_three_segments_and_three_scenarios(inputs):
    assert len(inputs.segments) == 3
    assert set(inputs.scenarios) == set(SCENARIOS)


@pytest.mark.parametrize("name", ["base", "BASE", " Upside ", "downside"])
def test_canonical_scenario_is_case_insensitive(name):
    assert canonical_scenario(name) in SCENARIOS


def test_unknown_scenario_rejected():
    with pytest.raises(ValueError, match="Unknown scenario"):
        canonical_scenario("moonshot")


def test_unbalanced_history_rejected(inputs):
    bad_hist = replace(inputs.hist, cash=(118.2, 92.6, 90.0))
    with pytest.raises(ValueError, match="does not balance"):
        replace(inputs, hist=bad_hist).validate()


def test_wrong_driver_length_rejected(inputs):
    with pytest.raises(ValueError, match="needs 5 values"):
        replace(inputs, dpo=(48.0,) * 4).validate()


def test_wrong_scenario_driver_length_rejected(inputs):
    drv = inputs.scenarios["Base"]
    bad = replace(drv, dso=(52.0,) * 3)
    with pytest.raises(ValueError, match="dso"):
        replace(inputs, scenarios={**inputs.scenarios, "Base": bad}).validate()


def test_negative_minimum_cash_rejected(inputs):
    with pytest.raises(ValueError, match="min_cash"):
        replace(inputs, min_cash=-1.0).validate()


def test_inputs_are_immutable(inputs):
    with pytest.raises(AttributeError):
        inputs.min_cash = 1.0  # type: ignore[misc]


def test_historical_series_lengths_checked(inputs):
    with pytest.raises(ValueError, match="historical series"):
        replace(inputs, hist=replace(inputs.hist, ap=(1.0, 2.0))).validate()


def test_fresh_inputs_are_equal_each_call():
    assert northwind_inputs() == northwind_inputs()
