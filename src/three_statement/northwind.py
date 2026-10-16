"""Fictional company data: Northwind Components Inc.

Northwind Components Inc. is a made-up manufacturer of industrial and automotive
components. Nothing here describes a real company; the figures are illustrative.
"""

from __future__ import annotations

from .inputs import Historicals, Inputs, ScenarioDrivers

COMPANY = "Northwind Components Inc."
SEGMENTS = ("Industrial Controls", "Automotive Sensors", "Aftermarket Services")


def northwind_inputs() -> Inputs:
    """Return the default (fictional) Northwind inputs: 3 historical and 5 forecast years."""
    hist = Historicals(
        years=(2023, 2024, 2025),
        segment_revenue=((310.0, 332.5, 356.8), (255.0, 281.4, 311.2), (118.0, 129.9, 142.7)),
        cogs=(437.1, 472.3, 510.7),
        sga=(99.0, 106.4, 114.3),
        rnd=(34.2, 37.2, 41.3),
        da=(30.5, 32.8, 35.4),
        int_exp=(14.6, 13.9, 13.1),
        int_inc=(1.9, 2.3, 2.6),
        tax=(17.4, 20.9, 24.6),
        cash=(118.2, 92.6, 88.6),
        ar=(95.4, 106.0, 115.5),
        inventory=(92.2, 98.3, 104.9),
        oca=(17.1, 18.6, 20.3),
        ppe=(226.0, 241.3, 258.0),
        goodwill=(140.0, 140.0, 140.0),
        ap=(56.3, 62.1, 67.2),
        accrued=(37.6, 40.9, 44.6),
        revolver=(15.0, 0.0, 0.0),
        term_loan=(210.0, 195.0, 180.0),
        other_ltl=(30.0, 30.0, 30.0),
        apic=(120.0, 120.0, 120.0),
        retained=(240.0, 283.8, 335.5),
        treasury=(-20.0, -35.0, -50.0),
        shares_end=(41.5, 40.8, 40.0),
        shares_wavg=(41.8, 41.1, 40.4),
    )
    base = ScenarioDrivers(
        growth=(
            (0.065, 0.060, 0.055, 0.050, 0.045),
            (0.090, 0.085, 0.080, 0.070, 0.060),
            (0.070, 0.070, 0.065, 0.060, 0.055),
        ),
        gross_margin=(0.373, 0.375, 0.378, 0.380, 0.380),
        sga_pct=(0.140, 0.139, 0.138, 0.137, 0.136),
        capex_pct=(0.065, 0.063, 0.060, 0.058, 0.055),
        dso=(52.0, 52.0, 52.0, 51.5, 51.0),
        dio=(75.0, 74.0, 74.0, 73.0, 73.0),
    )
    upside = ScenarioDrivers(
        growth=(
            (0.085, 0.080, 0.075, 0.070, 0.060),
            (0.120, 0.110, 0.100, 0.090, 0.080),
            (0.090, 0.090, 0.085, 0.080, 0.070),
        ),
        gross_margin=(0.378, 0.384, 0.389, 0.393, 0.395),
        sga_pct=(0.138, 0.136, 0.134, 0.132, 0.131),
        capex_pct=(0.068, 0.066, 0.063, 0.060, 0.058),
        dso=(51.0, 50.0, 50.0, 49.5, 49.0),
        dio=(74.0, 72.0, 71.0, 70.0, 70.0),
    )
    downside = ScenarioDrivers(
        growth=(
            (-0.020, 0.005, 0.025, 0.035, 0.040),
            (-0.060, -0.010, 0.030, 0.050, 0.055),
            (0.010, 0.020, 0.030, 0.035, 0.040),
        ),
        gross_margin=(0.355, 0.353, 0.358, 0.363, 0.365),
        sga_pct=(0.148, 0.149, 0.146, 0.143, 0.141),
        capex_pct=(0.050, 0.050, 0.052, 0.055, 0.055),
        dso=(56.0, 58.0, 57.0, 56.0, 55.0),
        dio=(82.0, 86.0, 84.0, 81.0, 79.0),
    )
    inputs = Inputs(
        company=COMPANY,
        segments=SEGMENTS,
        hist=hist,
        scenarios={"Base": base, "Upside": upside, "Downside": downside},
        rnd_pct=(0.051,) * 5,
        oca_pct=(0.025,) * 5,
        accrued_pct=(0.055,) * 5,
        dpo=(48.0,) * 5,
        da_rate=(0.145,) * 5,
        tax_rate=(0.25,) * 5,
        tl_rate=(0.0675,) * 5,
        revolver_rate=(0.0725,) * 5,
        cash_rate=(0.025,) * 5,
        tl_amort=(25.0,) * 5,
        payout=(0.30,) * 5,
        buybacks=(30.0, 30.0, 30.0, 35.0, 35.0),
        issuance=(5.0,) * 5,
        share_price=(28.0, 29.5, 31.0, 32.5, 34.0),
        min_cash=60.0,
        revolver_limit=150.0,
    )
    inputs.validate()
    return inputs
