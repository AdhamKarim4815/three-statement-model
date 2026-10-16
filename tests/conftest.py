from __future__ import annotations

import pytest

from three_statement import SCENARIOS, northwind_inputs, run_model


@pytest.fixture(scope="session")
def inputs():
    return northwind_inputs()


@pytest.fixture(scope="session")
def results(inputs):
    """Opening-basis results for every scenario."""
    return {s: run_model(inputs, s) for s in SCENARIOS}


@pytest.fixture(scope="session")
def avg_results(inputs):
    """Average-basis results for every scenario."""
    return {s: run_model(inputs, s, "average") for s in SCENARIOS}
