"""Integrated three-statement financial model: Python engine and Excel generator."""

__version__ = "1.0.0"

from .engine import ModelResult, run_model
from .inputs import SCENARIOS, Inputs
from .northwind import northwind_inputs

__all__ = ["SCENARIOS", "Inputs", "ModelResult", "__version__", "northwind_inputs", "run_model"]
