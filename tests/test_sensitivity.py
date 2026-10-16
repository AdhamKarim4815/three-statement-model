from __future__ import annotations

import pytest

from three_statement import SCENARIOS, run_model
from three_statement.report import format_sensitivity
from three_statement.sensitivity import DEFAULT_GROWTH_SHIFTS, DEFAULT_MARGIN_SHIFTS, sensitivity_grid, shifted_inputs


@pytest.fixture(scope="module")
def grid(inputs):
    return sensitivity_grid(inputs, "Base")


def test_grid_shape(grid):
    assert len(grid.ebitda) == len(DEFAULT_MARGIN_SHIFTS)
    assert all(len(row) == len(DEFAULT_GROWTH_SHIFTS) for row in grid.ebitda)
    assert len(grid.fcf) == len(DEFAULT_MARGIN_SHIFTS)
    assert grid.year == 2030


def test_centre_cell_equals_unshifted_model(inputs, grid, results):
    c = len(DEFAULT_MARGIN_SHIFTS) // 2
    assert grid.ebitda[c][c] == pytest.approx(results["Base"].last("ebitda"))
    assert grid.fcf[c][c] == pytest.approx(results["Base"].last("fcf"))


def test_every_cell_is_a_full_model_rerun(inputs, grid):
    res = run_model(shifted_inputs(inputs, "Base", 0.01, -0.02), "Base")
    assert grid.ebitda[0][3] == pytest.approx(res.last("ebitda"))
    assert grid.fcf[0][3] == pytest.approx(res.last("fcf"))


def test_ebitda_rises_with_growth_and_margin(grid):
    for row in grid.ebitda:
        assert row == sorted(row)
    for j in range(len(grid.growth_shifts)):
        column = [grid.ebitda[i][j] for i in range(len(grid.margin_shifts))]
        assert column == sorted(column)


def test_fcf_rises_with_margin(grid):
    for j in range(len(grid.growth_shifts)):
        column = [grid.fcf[i][j] for i in range(len(grid.margin_shifts))]
        assert column == sorted(column)


def test_one_point_of_margin_is_worth_about_one_point_of_revenue(grid, results):
    c = len(DEFAULT_MARGIN_SHIFTS) // 2
    step = grid.ebitda[c + 1][c] - grid.ebitda[c][c]
    assert step == pytest.approx(0.01 * results["Base"].last("revenue"))


def test_shifted_inputs_do_not_mutate_original(inputs):
    before = inputs.scenarios["Base"].gross_margin
    shifted = shifted_inputs(inputs, "base", 0.05, 0.05)
    assert inputs.scenarios["Base"].gross_margin == before
    assert shifted.scenarios["Base"].gross_margin[0] == pytest.approx(before[0] + 0.05)
    assert shifted.scenarios["Upside"] == inputs.scenarios["Upside"]


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_grid_scenario_label_and_centre(inputs, scenario):
    g = sensitivity_grid(inputs, scenario, growth_shifts=(0.0,), margin_shifts=(0.0,))
    assert g.scenario == scenario
    assert g.ebitda[0][0] == pytest.approx(run_model(inputs, scenario).last("ebitda"))


def test_text_rendering(grid):
    text = format_sensitivity(grid)
    assert "FY2030E EBITDA" in text and "free cash flow" in text
    assert text.count("+2.0pp") == 2
