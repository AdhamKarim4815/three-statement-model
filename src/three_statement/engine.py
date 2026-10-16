"""Integrated three-statement forecasting engine.

The engine mirrors, line for line, the formulas written into the Excel workbook
(see :mod:`three_statement.excel`). Two interest conventions are supported:

``opening``
    Interest accrues on opening balances. No circularity; this is what the
    Excel workbook uses.
``average``
    Interest accrues on the average of opening and closing balances. Closing
    balances depend on net income, which depends on interest, so each forecast
    year is solved by fixed-point iteration until the balances stop moving.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .inputs import Inputs, canonical_scenario

INTEREST_BASES = ("opening", "average")


@dataclass
class ModelResult:
    """Output of :func:`run_model`: every line item for every year (history + forecast)."""

    scenario: str
    interest_basis: str
    years: list[int]
    n_hist: int
    lines: dict[str, list[float | None]]
    iterations: list[int] = field(default_factory=list)
    residuals: list[float] = field(default_factory=list)

    def __getitem__(self, key: str) -> list[float | None]:
        return self.lines[key]

    def forecast(self, key: str) -> list[float]:
        """Forecast-period values of ``key`` (never None)."""
        return [v for v in self.lines[key][self.n_hist :] if v is not None]

    def last(self, key: str) -> float:
        """Final forecast year value of ``key``."""
        value = self.lines[key][-1]
        if value is None:
            raise KeyError(f"{key} has no value in the final year")
        return value


def _income_subtotals(p: dict[str, float], n_seg: int) -> None:
    p["revenue"] = sum(p[f"rev_s{k + 1}"] for k in range(n_seg))
    p["gross_profit"] = p["revenue"] - p["cogs"]
    p["ebitda"] = p["gross_profit"] - p["sga"] - p["rnd"]
    p["ebit"] = p["ebitda"] - p["da"]
    p["ebt"] = p["ebit"] - p["int_exp"] + p["int_inc"]
    p["net_income"] = p["ebt"] - p["tax"]


def _balance_subtotals(p: dict[str, float]) -> None:
    p["total_current_assets"] = p["cash"] + p["ar"] + p["inventory"] + p["oca"]
    p["total_assets"] = p["total_current_assets"] + p["ppe"] + p["goodwill"]
    p["total_liabilities"] = p["ap"] + p["accrued"] + p["revolver"] + p["term_loan"] + p["other_ltl"]
    p["total_equity"] = p["apic"] + p["retained"] + p["treasury"]
    p["total_le"] = p["total_liabilities"] + p["total_equity"]
    p["bs_check"] = p["total_assets"] - p["total_le"]


def _historical_periods(inputs: Inputs) -> list[dict[str, float]]:
    h = inputs.hist
    periods: list[dict[str, float]] = []
    for i in range(inputs.n_hist):
        p: dict[str, float] = {}
        for k in range(len(inputs.segments)):
            p[f"rev_s{k + 1}"] = h.segment_revenue[k][i]
        for name in (
            "cogs", "sga", "rnd", "da", "int_exp", "int_inc", "tax", "cash", "ar", "inventory",
            "oca", "ppe", "goodwill", "ap", "accrued", "revolver", "term_loan", "other_ltl",
            "apic", "retained", "treasury",
        ):  # fmt: skip
            p[name] = getattr(h, name)[i]
        p["shares_end"] = h.shares_end[i]
        p["shares_wavg"] = h.shares_wavg[i]
        _income_subtotals(p, len(inputs.segments))
        _balance_subtotals(p)
        periods.append(p)
    return periods


def _forecast_year(
    inputs: Inputs, scenario: str, t: int, prev: dict[str, float], basis: str, tol: float, max_iter: int
) -> tuple[dict[str, float], int, float]:
    """Build forecast year ``t`` (0-based) from the prior period ``prev``."""
    drv = inputs.scenarios[scenario]
    n_seg = len(inputs.segments)
    p: dict[str, float] = {}

    # Income statement above the interest line
    for k in range(n_seg):
        p[f"rev_s{k + 1}"] = prev[f"rev_s{k + 1}"] * (1.0 + drv.growth[k][t])
    revenue = sum(p[f"rev_s{k + 1}"] for k in range(n_seg))
    p["cogs"] = revenue * (1.0 - drv.gross_margin[t])
    p["sga"] = revenue * drv.sga_pct[t]
    p["rnd"] = revenue * inputs.rnd_pct[t]

    # PP&E roll-forward
    p["capex"] = drv.capex_pct[t] * revenue
    p["da"] = inputs.da_rate[t] * prev["ppe"]
    p["ppe"] = prev["ppe"] + p["capex"] - p["da"]
    p["goodwill"] = prev["goodwill"]
    p["other_ltl"] = prev["other_ltl"]

    # Working capital (days based)
    days = inputs.days_in_year
    p["ar"] = drv.dso[t] / days * revenue
    p["inventory"] = drv.dio[t] / days * p["cogs"]
    p["oca"] = inputs.oca_pct[t] * revenue
    p["ap"] = inputs.dpo[t] / days * p["cogs"]
    p["accrued"] = inputs.accrued_pct[t] * revenue

    # Term loan: mandatory amortization is independent of cash
    p["tl_amort"] = min(prev["term_loan"], inputs.tl_amort[t])
    p["term_loan"] = prev["term_loan"] - p["tl_amort"]
    int_tl = inputs.tl_rate[t] * (
        prev["term_loan"] if basis == "opening" else 0.5 * (prev["term_loan"] + p["term_loan"])
    )

    # Equity and share count
    p["issuance"] = inputs.issuance[t]
    p["buybacks"] = inputs.buybacks[t]
    p["sh_open"] = prev["shares_end"]
    p["sh_rep"] = p["buybacks"] / inputs.share_price[t]
    p["sh_iss"] = p["issuance"] / inputs.share_price[t]
    p["shares_end"] = p["sh_open"] - p["sh_rep"] + p["sh_iss"]
    p["shares_wavg"] = 0.5 * (p["sh_open"] + p["shares_end"])
    p["apic"] = prev["apic"] + p["issuance"]
    p["treasury"] = prev["treasury"] - p["buybacks"]

    d_nwc_items = {
        "ar": p["ar"] - prev["ar"],
        "inventory": p["inventory"] - prev["inventory"],
        "oca": p["oca"] - prev["oca"],
        "ap": p["ap"] - prev["ap"],
        "accrued": p["accrued"] - prev["accrued"],
    }
    d_nwc = (
        d_nwc_items["ar"] + d_nwc_items["inventory"] + d_nwc_items["oca"]
        - d_nwc_items["ap"] - d_nwc_items["accrued"]
    )  # fmt: skip

    # Fixed-point solve for the interest-sensitive block (one pass for opening basis)
    rev_close_guess = prev["revolver"]
    cash_close_guess = prev["cash"]
    iterations = 0
    residual = 0.0
    while True:
        iterations += 1
        if basis == "opening":
            int_rev = inputs.revolver_rate[t] * prev["revolver"]
            int_inc = inputs.cash_rate[t] * prev["cash"]
        else:
            int_rev = inputs.revolver_rate[t] * 0.5 * (prev["revolver"] + rev_close_guess)
            int_inc = inputs.cash_rate[t] * 0.5 * (prev["cash"] + cash_close_guess)
        p["int_tl"], p["int_rev"] = int_tl, int_rev
        p["int_exp"], p["int_inc"] = int_tl + int_rev, int_inc
        p["ebitda"] = revenue - p["cogs"] - p["sga"] - p["rnd"]
        p["ebt"] = p["ebitda"] - p["da"] - p["int_exp"] + p["int_inc"]
        p["tax"] = max(0.0, p["ebt"]) * inputs.tax_rate[t]
        ni = p["ebt"] - p["tax"]
        p["dividends"] = inputs.payout[t] * max(0.0, ni)
        cfo = ni + p["da"] - d_nwc
        cash_pre = (
            prev["cash"] + cfo - p["capex"] - p["tl_amort"] + p["issuance"]
            - p["dividends"] - p["buybacks"]
        )  # fmt: skip
        p["cash_pre"] = cash_pre
        shortfall = inputs.min_cash - cash_pre
        p["rev_change"] = max(-prev["revolver"], min(shortfall, inputs.revolver_limit - prev["revolver"]))
        rev_close = prev["revolver"] + p["rev_change"]
        cash_close = cash_pre + p["rev_change"]
        residual = abs(rev_close - rev_close_guess) + abs(cash_close - cash_close_guess)
        rev_close_guess, cash_close_guess = rev_close, cash_close
        if basis == "opening":
            residual = 0.0
            break
        if residual < tol or iterations >= max_iter:
            break
    if basis == "average" and residual >= tol:
        raise RuntimeError(f"interest iteration did not converge in year index {t}")
    p["revolver"] = rev_close
    p["cash"] = cash_close
    _income_subtotals(p, len(inputs.segments))
    p["retained"] = prev["retained"] + p["net_income"] - p["dividends"]
    _balance_subtotals(p)

    # Cash flow statement
    p["cf_ni"] = p["net_income"]
    p["cf_da"] = p["da"]
    p["cf_ar"] = -d_nwc_items["ar"]
    p["cf_inv"] = -d_nwc_items["inventory"]
    p["cf_oca"] = -d_nwc_items["oca"]
    p["cf_ap"] = d_nwc_items["ap"]
    p["cf_accr"] = d_nwc_items["accrued"]
    p["cfo"] = p["cf_ni"] + p["cf_da"] + p["cf_ar"] + p["cf_inv"] + p["cf_oca"] + p["cf_ap"] + p["cf_accr"]
    p["cf_capex"] = -p["capex"]
    p["cfi"] = p["cf_capex"]
    p["cf_tl"] = -p["tl_amort"]
    p["cf_rev"] = p["rev_change"]
    p["cf_issue"] = p["issuance"]
    p["cf_div"] = -p["dividends"]
    p["cf_bb"] = -p["buybacks"]
    p["cff"] = p["cf_tl"] + p["cf_rev"] + p["cf_issue"] + p["cf_div"] + p["cf_bb"]
    p["net_change"] = p["cfo"] + p["cfi"] + p["cff"]
    p["cash_begin"] = prev["cash"]
    p["cash_end"] = p["cash_begin"] + p["net_change"]
    p["cash_close"] = cash_close
    return p, iterations, residual


def _derive(periods: list[dict[str, float]], n_hist: int, n_seg: int, days: float) -> dict[str, list[float | None]]:
    """Add implied flows, schedule openings and ratio lines; return line -> series."""
    n = len(periods)
    for i in range(1, n):
        prev, p = periods[i - 1], periods[i]
        p["ppe_open"] = prev["ppe"]
        p["tl_open"] = prev["term_loan"]
        p["rev_open"] = prev["revolver"]
        p["re_open"] = prev["retained"]
        p["apic_open"] = prev["apic"]
        p["tre_open"] = prev["treasury"]
        if i < n_hist:  # implied historical flows
            p["capex"] = p["ppe"] - prev["ppe"] + p["da"]
            p["tl_amort"] = prev["term_loan"] - p["term_loan"]
            p["rev_change"] = p["revolver"] - prev["revolver"]
            p["dividends"] = prev["retained"] + p["net_income"] - p["retained"]
            p["buybacks"] = prev["treasury"] - p["treasury"]
            p["issuance"] = p["apic"] - prev["apic"]
        else:
            p["cash_open"] = prev["cash"]
    for p in periods:
        p["nwc"] = p["ar"] + p["inventory"] + p["oca"] - p["ap"] - p["accrued"]
        p["total_debt"] = p["revolver"] + p["term_loan"]
        p["net_debt"] = p["total_debt"] - p["cash"]
        p["gross_margin"] = p["gross_profit"] / p["revenue"]
        p["ebitda_margin"] = p["ebitda"] / p["revenue"]
        p["ebit_margin"] = p["ebit"] / p["revenue"]
        p["net_margin"] = p["net_income"] / p["revenue"]
        p["eff_tax"] = p["tax"] / p["ebt"] if p["ebt"] > 0 else 0.0
        p["eps"] = p["net_income"] / p["shares_wavg"]
        p["net_debt_ebitda"] = p["net_debt"] / p["ebitda"]
        p["int_cov"] = p["ebit"] / p["int_exp"] if p["int_exp"] > 0 else None
        p["invested_capital"] = p["total_debt"] + p["total_equity"] - p["cash"]
        p["sga_pct"] = p["sga"] / p["revenue"]
        p["rnd_pct"] = p["rnd"] / p["revenue"]
        p["oca_pct"] = p["oca"] / p["revenue"]
        p["accrued_pct"] = p["accrued"] / p["revenue"]
        p["dso"] = p["ar"] / p["revenue"] * days
        p["dio"] = p["inventory"] / p["cogs"] * days
        p["dpo"] = p["ap"] / p["cogs"] * days
    for i in range(n):
        p = periods[i]
        if i >= 1:
            prev = periods[i - 1]
            p["rev_growth"] = p["revenue"] / prev["revenue"] - 1.0
            for k in range(n_seg):
                p[f"g_s{k + 1}"] = p[f"rev_s{k + 1}"] / prev[f"rev_s{k + 1}"] - 1.0
            p["capex_pct"] = p["capex"] / p["revenue"]
            p["da_rate"] = p["da"] / prev["ppe"]
            p["payout"] = p["dividends"] / p["net_income"] if p["net_income"] > 0 else 0.0
            p["d_nwc"] = p["nwc"] - prev["nwc"]
            p["dps"] = p["dividends"] / p["shares_wavg"]
            avg_equity = 0.5 * (prev["total_equity"] + p["total_equity"])
            p["roe"] = p["net_income"] / avg_equity
            avg_ic = 0.5 * (p["invested_capital"] + prev["invested_capital"])
            p["roic"] = p["ebit"] * (1.0 - p["eff_tax"]) / avg_ic
        if i >= n_hist:
            p["fcf"] = p["cfo"] + p["cf_capex"]
            p["fcf_margin"] = p["fcf"] / p["revenue"]
            p["fcf_conv"] = p["fcf"] / p["ebitda"]
    keys: list[str] = []
    for p in periods:
        for key in p:
            if key not in keys:
                keys.append(key)
    return {key: [p.get(key) for p in periods] for key in keys}


def run_model(
    inputs: Inputs, scenario: str = "Base", interest_basis: str = "opening", tol: float = 1e-10, max_iter: int = 200
) -> ModelResult:
    """Run the integrated model for one scenario.

    Parameters
    ----------
    inputs:
        Model inputs (see :func:`three_statement.northwind.northwind_inputs`).
    scenario:
        ``Base``, ``Upside`` or ``Downside`` (case-insensitive).
    interest_basis:
        ``opening`` (default, no circularity) or ``average`` (iterative).
    tol, max_iter:
        Convergence tolerance on the sum of absolute balance changes between
        iterations, and the iteration cap, used only for ``average``.
    """
    scenario = canonical_scenario(scenario)
    if interest_basis not in INTEREST_BASES:
        raise ValueError(f"interest_basis must be one of {INTEREST_BASES}")
    inputs.validate()
    periods = _historical_periods(inputs)
    iterations: list[int] = []
    residuals: list[float] = []
    for t in range(inputs.n_forecast):
        p, its, res = _forecast_year(inputs, scenario, t, periods[-1], interest_basis, tol, max_iter)
        periods.append(p)
        iterations.append(its)
        residuals.append(res)
    lines = _derive(periods, inputs.n_hist, len(inputs.segments), inputs.days_in_year)
    return ModelResult(
        scenario=scenario,
        interest_basis=interest_basis,
        years=inputs.years,
        n_hist=inputs.n_hist,
        lines=lines,
        iterations=iterations,
        residuals=residuals,
    )
