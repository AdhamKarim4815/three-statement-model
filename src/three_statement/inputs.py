"""Typed input containers for the three-statement model.

Everything the engine and the Excel builder need lives in :class:`Inputs`.
All monetary amounts are in millions of US dollars unless stated otherwise.
"""

from __future__ import annotations

from dataclasses import dataclass

SCENARIOS: tuple[str, ...] = ("Base", "Upside", "Downside")


def canonical_scenario(name: str) -> str:
    """Return the canonical scenario name for a case-insensitive input."""
    for scenario in SCENARIOS:
        if scenario.lower() == str(name).strip().lower():
            return scenario
    raise ValueError(f"Unknown scenario {name!r}; choose one of {', '.join(SCENARIOS)}")


@dataclass(frozen=True)
class Historicals:
    """Reported (fictional) results. Every tuple has one entry per historical year."""

    years: tuple[int, ...]
    segment_revenue: tuple[tuple[float, ...], ...]  # [segment][year]
    cogs: tuple[float, ...]  # cost of goods sold, excluding D&A
    sga: tuple[float, ...]
    rnd: tuple[float, ...]
    da: tuple[float, ...]
    int_exp: tuple[float, ...]
    int_inc: tuple[float, ...]
    tax: tuple[float, ...]
    cash: tuple[float, ...]
    ar: tuple[float, ...]
    inventory: tuple[float, ...]
    oca: tuple[float, ...]  # prepaid and other current assets
    ppe: tuple[float, ...]  # net PP&E
    goodwill: tuple[float, ...]  # goodwill and intangibles
    ap: tuple[float, ...]
    accrued: tuple[float, ...]
    revolver: tuple[float, ...]
    term_loan: tuple[float, ...]
    other_ltl: tuple[float, ...]
    apic: tuple[float, ...]
    retained: tuple[float, ...]
    treasury: tuple[float, ...]  # negative balance
    shares_end: tuple[float, ...]  # millions
    shares_wavg: tuple[float, ...]  # millions


@dataclass(frozen=True)
class ScenarioDrivers:
    """Drivers that change between Base, Upside and Downside (one value per forecast year)."""

    growth: tuple[tuple[float, ...], ...]  # [segment][year]
    gross_margin: tuple[float, ...]
    sga_pct: tuple[float, ...]
    capex_pct: tuple[float, ...]
    dso: tuple[float, ...]
    dio: tuple[float, ...]


@dataclass(frozen=True)
class Inputs:
    """Complete model inputs: history, scenario drivers and scenario-independent drivers."""

    company: str
    segments: tuple[str, ...]
    hist: Historicals
    scenarios: dict[str, ScenarioDrivers]
    # Scenario-independent drivers, one value per forecast year.
    rnd_pct: tuple[float, ...]
    oca_pct: tuple[float, ...]
    accrued_pct: tuple[float, ...]
    dpo: tuple[float, ...]
    da_rate: tuple[float, ...]  # D&A as % of opening net PP&E
    tax_rate: tuple[float, ...]
    tl_rate: tuple[float, ...]
    revolver_rate: tuple[float, ...]
    cash_rate: tuple[float, ...]
    tl_amort: tuple[float, ...]  # mandatory term loan amortization, $m
    payout: tuple[float, ...]  # dividend payout ratio on positive net income
    buybacks: tuple[float, ...]  # $m
    issuance: tuple[float, ...]  # $m
    share_price: tuple[float, ...]  # $ per share used for buybacks and issuance
    # Scalars
    days_in_year: float = 365.0
    min_cash: float = 40.0
    revolver_limit: float = 150.0
    tolerance: float = 0.001

    @property
    def n_hist(self) -> int:
        return len(self.hist.years)

    @property
    def n_forecast(self) -> int:
        return len(self.rnd_pct)

    @property
    def years(self) -> list[int]:
        last = self.hist.years[-1]
        return list(self.hist.years) + [last + k for k in range(1, self.n_forecast + 1)]

    def validate(self) -> None:
        """Raise ValueError when shapes are inconsistent or history does not balance."""
        nh, nf, ns = self.n_hist, self.n_forecast, len(self.segments)
        h = self.hist
        for fname, seq in vars(h).items():
            if fname in ("years", "segment_revenue"):
                continue
            if len(seq) != nh:
                raise ValueError(f"historical series {fname!r} needs {nh} values, got {len(seq)}")
        if len(h.segment_revenue) != ns or any(len(s) != nh for s in h.segment_revenue):
            raise ValueError("segment_revenue must be [segment][historical year]")
        flat = (
            "rnd_pct", "oca_pct", "accrued_pct", "dpo", "da_rate", "tax_rate", "tl_rate",
            "revolver_rate", "cash_rate", "tl_amort", "payout", "buybacks", "issuance",
            "share_price",
        )  # fmt: skip
        for fname in flat:
            if len(getattr(self, fname)) != nf:
                raise ValueError(f"driver {fname!r} needs {nf} values")
        for scenario in SCENARIOS:
            if scenario not in self.scenarios:
                raise ValueError(f"missing scenario {scenario!r}")
            drv = self.scenarios[scenario]
            if len(drv.growth) != ns or any(len(g) != nf for g in drv.growth):
                raise ValueError(f"{scenario}: growth must be [segment][forecast year]")
            for fname in ("gross_margin", "sga_pct", "capex_pct", "dso", "dio"):
                if len(getattr(drv, fname)) != nf:
                    raise ValueError(f"{scenario}: driver {fname!r} needs {nf} values")
        for i in range(nh):
            assets = h.cash[i] + h.ar[i] + h.inventory[i] + h.oca[i] + h.ppe[i] + h.goodwill[i]
            claims = (
                h.ap[i] + h.accrued[i] + h.revolver[i] + h.term_loan[i] + h.other_ltl[i]
                + h.apic[i] + h.retained[i] + h.treasury[i]
            )  # fmt: skip
            if abs(assets - claims) > 1e-6:
                raise ValueError(f"historical balance sheet {h.years[i]} does not balance")
        if self.min_cash < 0 or self.revolver_limit < 0 or self.days_in_year <= 0:
            raise ValueError("min_cash and revolver_limit must be >= 0; days_in_year > 0")
