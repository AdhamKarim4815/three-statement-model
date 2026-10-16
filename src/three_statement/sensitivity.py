"""Two-way sensitivity of final-year EBITDA and free cash flow.

The grid shifts every segment's revenue growth (in every forecast year) by a
number of percentage points, and the gross margin (in every forecast year) by a
number of percentage points, then re-runs the full integrated model for each
cell. The result is exact (not a linear approximation).
"""

from __future__ import annotations

from dataclasses import dataclass, replace

from .engine import run_model
from .inputs import Inputs, ScenarioDrivers, canonical_scenario

DEFAULT_GROWTH_SHIFTS: tuple[float, ...] = (-0.02, -0.01, 0.0, 0.01, 0.02)
DEFAULT_MARGIN_SHIFTS: tuple[float, ...] = (-0.02, -0.01, 0.0, 0.01, 0.02)


@dataclass(frozen=True)
class SensitivityGrid:
    """Result grids indexed ``[margin_shift_row][growth_shift_col]``."""

    scenario: str
    year: int
    growth_shifts: tuple[float, ...]
    margin_shifts: tuple[float, ...]
    ebitda: list[list[float]]
    fcf: list[list[float]]


def shifted_inputs(inputs: Inputs, scenario: str, growth_shift: float, margin_shift: float) -> Inputs:
    """Copy of ``inputs`` with the chosen scenario's growth and gross margin shifted."""
    scenario = canonical_scenario(scenario)
    drv = inputs.scenarios[scenario]
    new = ScenarioDrivers(
        growth=tuple(tuple(g + growth_shift for g in seg) for seg in drv.growth),
        gross_margin=tuple(m + margin_shift for m in drv.gross_margin),
        sga_pct=drv.sga_pct,
        capex_pct=drv.capex_pct,
        dso=drv.dso,
        dio=drv.dio,
    )
    return replace(inputs, scenarios={**inputs.scenarios, scenario: new})


def sensitivity_grid(
    inputs: Inputs,
    scenario: str = "Base",
    growth_shifts: tuple[float, ...] = DEFAULT_GROWTH_SHIFTS,
    margin_shifts: tuple[float, ...] = DEFAULT_MARGIN_SHIFTS,
    interest_basis: str = "opening",
) -> SensitivityGrid:
    """Compute final-year EBITDA and FCF over a growth-shift by margin-shift grid."""
    scenario = canonical_scenario(scenario)
    ebitda: list[list[float]] = []
    fcf: list[list[float]] = []
    for m in margin_shifts:
        row_e: list[float] = []
        row_f: list[float] = []
        for g in growth_shifts:
            result = run_model(shifted_inputs(inputs, scenario, g, m), scenario, interest_basis)
            row_e.append(result.last("ebitda"))
            row_f.append(result.last("fcf"))
        ebitda.append(row_e)
        fcf.append(row_f)
    return SensitivityGrid(
        scenario=scenario,
        year=inputs.years[-1],
        growth_shifts=tuple(growth_shifts),
        margin_shifts=tuple(margin_shifts),
        ebitda=ebitda,
        fcf=fcf,
    )
