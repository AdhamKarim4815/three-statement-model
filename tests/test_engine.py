from __future__ import annotations

from dataclasses import replace

import pytest

from three_statement import SCENARIOS, run_model
from three_statement.engine import INTEREST_BASES
from three_statement.inputs import ScenarioDrivers

BASES = ("opening", "average")


def _stress(inputs, **changes):
    return replace(inputs, **changes)


def _scenario(inputs, name="Base", **driver_changes):
    drv = replace(inputs.scenarios[name], **driver_changes)
    return replace(inputs, scenarios={**inputs.scenarios, name: drv})


# ------------------------------------------------------------------ integrity


@pytest.mark.parametrize("scenario", SCENARIOS)
@pytest.mark.parametrize("basis", BASES)
def test_balance_sheet_balances_every_year(inputs, scenario, basis):
    res = run_model(inputs, scenario, basis)
    assert all(abs(v) < 1e-8 for v in res["bs_check"])
    for a, le in zip(res["total_assets"], res["total_le"], strict=True):
        assert a == pytest.approx(le, abs=1e-8)


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_cash_on_balance_sheet_ties_to_cash_flow(results, scenario):
    res = results[scenario]
    n = res.n_hist
    assert res["cash"][n:] == pytest.approx(res["cash_end"][n:], abs=1e-9)
    for i in range(n, len(res.years)):
        assert res["cash_begin"][i] == pytest.approx(res["cash"][i - 1])
        assert res["cash_end"][i] == pytest.approx(res["cash_begin"][i] + res["net_change"][i])


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_retained_earnings_roll_forward(results, scenario):
    res = results[scenario]
    for i in range(res.n_hist, len(res.years)):
        expected = res["retained"][i - 1] + res["net_income"][i] - res["dividends"][i]
        assert res["retained"][i] == pytest.approx(expected)


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_debt_roll_forward(results, scenario):
    res = results[scenario]
    for i in range(res.n_hist, len(res.years)):
        assert res["term_loan"][i] == pytest.approx(res["term_loan"][i - 1] + res["cf_tl"][i])
        assert res["revolver"][i] == pytest.approx(res["revolver"][i - 1] + res["cf_rev"][i])


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_cash_flow_statement_sums(results, scenario):
    res = results[scenario]
    for i in range(res.n_hist, len(res.years)):
        wc = sum(res[k][i] for k in ("cf_ar", "cf_inv", "cf_oca", "cf_ap", "cf_accr"))
        assert res["cfo"][i] == pytest.approx(res["net_income"][i] + res["da"][i] + wc)
        assert wc == pytest.approx(-res["d_nwc"][i])
        assert res["fcf"][i] == pytest.approx(res["cfo"][i] + res["cf_capex"][i])


def test_historical_subtotals(results):
    res = results["Base"]
    assert res["revenue"][:3] == pytest.approx([683.0, 743.8, 810.7])
    assert res["ebitda"][2] == pytest.approx(144.4)
    assert res["net_income"][2] == pytest.approx(73.9)


# ------------------------------------------------------------------ drivers


def test_segment_revenue_compounds_at_scenario_growth(inputs, results):
    res = results["Base"]
    for k in range(3):
        growth = inputs.scenarios["Base"].growth[k]
        level = res[f"rev_s{k + 1}"][2]
        for t, g in enumerate(growth):
            level *= 1 + g
            assert res[f"rev_s{k + 1}"][3 + t] == pytest.approx(level)


def test_total_revenue_is_sum_of_segments(results):
    res = results["Upside"]
    for i in range(8):
        assert res["revenue"][i] == pytest.approx(sum(res[f"rev_s{k}"][i] for k in (1, 2, 3)))


def test_gross_margin_and_opex_percent_drivers(inputs, results):
    res = results["Base"]
    drv = inputs.scenarios["Base"]
    for t in range(5):
        i = 3 + t
        assert res["gross_margin"][i] == pytest.approx(drv.gross_margin[t])
        assert res["sga"][i] == pytest.approx(res["revenue"][i] * drv.sga_pct[t])
        assert res["rnd"][i] == pytest.approx(res["revenue"][i] * inputs.rnd_pct[t])
        assert res["ebitda"][i] == pytest.approx(res["gross_profit"][i] - res["sga"][i] - res["rnd"][i])


def test_ppe_roll_forward(inputs, results):
    res = results["Base"]
    drv = inputs.scenarios["Base"]
    for t in range(5):
        i = 3 + t
        capex = drv.capex_pct[t] * res["revenue"][i]
        da = inputs.da_rate[t] * res["ppe"][i - 1]
        assert res["capex"][i] == pytest.approx(capex)
        assert res["da"][i] == pytest.approx(da)
        assert res["ppe"][i] == pytest.approx(res["ppe"][i - 1] + capex - da)


def test_working_capital_days_math(inputs, results):
    res = results["Downside"]
    drv = inputs.scenarios["Downside"]
    for t in range(5):
        i = 3 + t
        assert res["ar"][i] == pytest.approx(drv.dso[t] / 365 * res["revenue"][i])
        assert res["inventory"][i] == pytest.approx(drv.dio[t] / 365 * res["cogs"][i])
        assert res["ap"][i] == pytest.approx(inputs.dpo[t] / 365 * res["cogs"][i])
        assert res["dso"][i] == pytest.approx(drv.dso[t])
        assert res["dio"][i] == pytest.approx(drv.dio[t])
        assert res["dpo"][i] == pytest.approx(inputs.dpo[t])


def test_days_in_year_input_is_respected(inputs):
    res = run_model(replace(inputs, days_in_year=360.0))
    assert res["ar"][3] == pytest.approx(inputs.scenarios["Base"].dso[0] / 360 * res["revenue"][3])


def test_longer_receivable_days_absorb_cash(inputs):
    slow = run_model(_scenario(inputs, dso=(70.0,) * 5))
    base = run_model(inputs)
    assert slow["cfo"][3] < base["cfo"][3]
    assert slow["ar"][3] > base["ar"][3]


def test_tax_is_zero_on_pre_tax_loss(inputs):
    res = run_model(_scenario(inputs, gross_margin=(0.05,) * 5))
    assert min(res["ebt"][3:]) < 0
    assert all(t == 0 for t in res["tax"][3:])


def test_no_dividend_when_net_income_negative(inputs):
    res = run_model(_scenario(inputs, gross_margin=(0.05,) * 5))
    assert all(ni < 0 for ni in res["net_income"][3:])
    assert all(d == 0 for d in res["dividends"][3:])


def test_dividends_follow_payout_ratio(inputs, results):
    res = results["Base"]
    for t in range(5):
        assert res["dividends"][3 + t] == pytest.approx(inputs.payout[t] * res["net_income"][3 + t])


def test_buybacks_and_issuance_move_equity_and_shares(inputs, results):
    res = results["Base"]
    for t in range(5):
        i = 3 + t
        assert res["treasury"][i] == pytest.approx(res["treasury"][i - 1] - inputs.buybacks[t])
        assert res["apic"][i] == pytest.approx(res["apic"][i - 1] + inputs.issuance[t])
        shares = res["shares_end"][i - 1] - inputs.buybacks[t] / inputs.share_price[t]
        shares += inputs.issuance[t] / inputs.share_price[t]
        assert res["shares_end"][i] == pytest.approx(shares)


def test_issuance_switched_off(inputs):
    res = run_model(replace(inputs, issuance=(0.0,) * 5))
    assert res["apic"][3:] == pytest.approx([120.0] * 5)
    assert res["cf_issue"][3:] == [0.0] * 5


# ------------------------------------------------------------------ debt: term loan and revolver


def test_term_loan_amortizes_mandatory_amount(inputs, results):
    assert results["Base"]["term_loan"][2:] == pytest.approx([180.0, 155.0, 130.0, 105.0, 80.0, 55.0])


def test_term_loan_never_below_zero(inputs):
    res = run_model(replace(inputs, tl_amort=(100.0,) * 5))
    assert res["term_loan"][3:] == pytest.approx([80.0, 0.0, 0.0, 0.0, 0.0])
    assert res["tl_amort"][4] == pytest.approx(80.0)


def test_revolver_draws_when_cash_is_short(inputs):
    res = run_model(replace(inputs, buybacks=(150.0,) * 5))
    assert res["revolver"][3] > 0
    # cash is held at the minimum while the revolver funds the gap
    assert res["cash"][3] == pytest.approx(inputs.min_cash)
    assert res["cash_pre"][3] < inputs.min_cash
    assert res["rev_change"][3] == pytest.approx(inputs.min_cash - res["cash_pre"][3])


def test_revolver_repays_when_cash_is_in_excess(inputs):
    hist = replace(inputs.hist, revolver=(15.0, 0.0, 50.0), cash=(118.2, 92.6, 138.6))
    res = run_model(replace(inputs, hist=hist), "Upside")
    assert res["revolver"][2] == 50.0
    assert res["rev_change"][3] < 0
    assert res["revolver"][3] == pytest.approx(0.0, abs=1e-9) or res["revolver"][3] < 50.0
    # excess cash above the minimum only ever repays up to the opening balance
    for i in range(3, 8):
        assert res["revolver"][i] >= 0
        assert res["cash"][i] >= inputs.min_cash - 1e-9


def test_revolver_repays_only_the_excess_when_balance_is_large(inputs):
    hist = replace(inputs.hist, revolver=(15.0, 0.0, 200.0), cash=(118.2, 92.6, 288.6))
    res = run_model(replace(inputs, hist=hist, min_cash=250.0, revolver_limit=300.0), "Upside")
    repay = -res["rev_change"][3]
    assert 0 < repay < 200
    assert repay == pytest.approx(res["cash_pre"][3] - 250.0)
    assert res["cash"][3] == pytest.approx(250.0)
    assert res["revolver"][3] == pytest.approx(200.0 - repay)


def test_revolver_repayment_limited_to_opening_balance(inputs):
    hist = replace(inputs.hist, revolver=(15.0, 0.0, 20.0), cash=(118.2, 92.6, 108.6))
    res = run_model(replace(inputs, hist=hist), "Upside")
    assert res["rev_change"][3] == pytest.approx(-20.0)
    assert res["revolver"][3] == pytest.approx(0.0)
    assert res["cash"][3] > inputs.min_cash


def test_revolver_is_zero_when_cash_is_ample(inputs):
    res = run_model(replace(inputs, min_cash=0.0), "Upside")
    assert all(v == 0.0 for v in res["revolver"][2:])


@pytest.mark.parametrize("scenario", SCENARIOS)
@pytest.mark.parametrize("basis", BASES)
def test_revolver_never_negative_and_cash_at_least_minimum(inputs, scenario, basis):
    res = run_model(inputs, scenario, basis)
    assert min(res["revolver"]) >= 0
    assert min(res["cash"][3:]) >= inputs.min_cash - 1e-9


def test_revolver_never_negative_under_stress(inputs):
    for buyback in (0.0, 5.0, 60.0, 400.0):
        res = run_model(replace(inputs, buybacks=(buyback,) * 5))
        assert min(res["revolver"]) >= 0


def test_revolver_capped_at_commitment(inputs):
    res = run_model(replace(inputs, buybacks=(300.0,) * 5, revolver_limit=40.0))
    assert max(res["revolver"][3:]) == pytest.approx(40.0)
    assert min(res["cash"][3:]) < inputs.min_cash  # cannot be defended once the facility is full
    assert all(abs(v) < 1e-8 for v in res["bs_check"])


def test_downside_uses_revolver_and_base_repays_it(results):
    assert max(results["Downside"].forecast("revolver")) > 20
    base = results["Base"].forecast("revolver")
    assert max(base) > 0 and base[-1] == pytest.approx(0.0)  # drawn, then repaid


def test_downside_revolver_draws_in_a_cash_crunch_year(results):
    res = results["Downside"]
    drawn = [i for i in range(3, 8) if res["rev_change"][i] > 1e-9]
    assert drawn
    for i in drawn:
        assert res["cash"][i] == pytest.approx(60.0)


# ------------------------------------------------------------------ interest


def test_opening_basis_interest_uses_opening_balances(inputs, results):
    res = results["Downside"]
    for t in range(5):
        i = 3 + t
        assert res["int_tl"][i] == pytest.approx(inputs.tl_rate[t] * res["term_loan"][i - 1])
        assert res["int_rev"][i] == pytest.approx(inputs.revolver_rate[t] * res["revolver"][i - 1])
        assert res["int_inc"][i] == pytest.approx(inputs.cash_rate[t] * res["cash"][i - 1])
        assert res["int_exp"][i] == pytest.approx(res["int_tl"][i] + res["int_rev"][i])


def test_opening_basis_needs_one_pass(results):
    assert results["Base"].iterations == [1] * 5


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_average_interest_converges(avg_results, scenario):
    res = avg_results[scenario]
    assert all(1 < n < 40 for n in res.iterations)
    assert max(res.residuals) < 1e-10


@pytest.mark.parametrize("scenario", SCENARIOS)
def test_average_interest_is_a_true_fixed_point(inputs, avg_results, scenario):
    res = avg_results[scenario]
    for t in range(5):
        i = 3 + t
        tl = inputs.tl_rate[t] * 0.5 * (res["term_loan"][i - 1] + res["term_loan"][i])
        rev = inputs.revolver_rate[t] * 0.5 * (res["revolver"][i - 1] + res["revolver"][i])
        inc = inputs.cash_rate[t] * 0.5 * (res["cash"][i - 1] + res["cash"][i])
        assert res["int_exp"][i] == pytest.approx(tl + rev, abs=1e-8)
        assert res["int_inc"][i] == pytest.approx(inc, abs=1e-8)


def test_average_interest_is_lower_when_debt_amortizes(results, avg_results):
    assert avg_results["Base"]["int_exp"][3] < results["Base"]["int_exp"][3]
    assert avg_results["Base"]["net_income"][3] > results["Base"]["net_income"][3]


def test_average_and_opening_agree_before_interest(results, avg_results):
    for key in ("revenue", "ebitda", "ebit", "capex", "da"):
        assert avg_results["Base"][key] == pytest.approx(results["Base"][key])


def test_average_interest_tighter_tolerance_gives_same_answer(inputs):
    loose = run_model(inputs, "Downside", "average", tol=1e-6)
    tight = run_model(inputs, "Downside", "average", tol=1e-12)
    assert loose["net_income"] == pytest.approx(tight["net_income"], abs=1e-4)
    assert sum(tight.iterations) >= sum(loose.iterations)


def test_average_interest_nonconvergence_is_reported(inputs):
    with pytest.raises(RuntimeError, match="did not converge"):
        run_model(inputs, "Base", "average", tol=1e-14, max_iter=2)


def test_invalid_interest_basis_rejected(inputs):
    with pytest.raises(ValueError, match="interest_basis"):
        run_model(inputs, "Base", "monthly")
    assert INTEREST_BASES == ("opening", "average")


# ------------------------------------------------------------------ scenarios and ratios


def test_scenario_ordering_on_final_year_outputs(results):
    for key in ("revenue", "ebitda", "net_income", "fcf"):
        assert results["Upside"].last(key) > results["Base"].last(key) > results["Downside"].last(key)


def test_scenario_name_is_case_insensitive(inputs, results):
    assert run_model(inputs, "downside").last("cash") == pytest.approx(results["Downside"].last("cash"))


def test_net_debt_and_leverage_ratio(results):
    res = results["Base"]
    for i in range(8):
        assert res["net_debt"][i] == pytest.approx(res["term_loan"][i] + res["revolver"][i] - res["cash"][i])
        assert res["net_debt_ebitda"][i] == pytest.approx(res["net_debt"][i] / res["ebitda"][i])


def test_interest_coverage_and_roe_and_roic(results):
    res = results["Base"]
    assert res["int_cov"][3] == pytest.approx(res["ebit"][3] / res["int_exp"][3])
    eq = 0.5 * (res["total_equity"][2] + res["total_equity"][3])
    assert res["roe"][3] == pytest.approx(res["net_income"][3] / eq)
    ic = lambda i: res["total_debt"][i] + res["total_equity"][i] - res["cash"][i]  # noqa: E731
    nopat = res["ebit"][3] * (1 - res["eff_tax"][3])
    assert res["roic"][3] == pytest.approx(nopat / (0.5 * (ic(2) + ic(3))))


def test_interest_coverage_undefined_without_debt(inputs):
    hist = replace(inputs.hist, term_loan=(0.0, 0.0, 0.0), revolver=(0.0, 0.0, 0.0), int_exp=(0.0, 0.0, 0.0))
    hist = replace(hist, retained=(465.0, 478.8, 515.5))
    res = run_model(replace(inputs, hist=hist, tl_amort=(0.0,) * 5), "Upside")
    assert res["int_cov"][3] is None


def test_eps_uses_weighted_average_shares(results):
    res = results["Base"]
    assert res["eps"][3] == pytest.approx(res["net_income"][3] / res["shares_wavg"][3])
    assert res["shares_wavg"][3] == pytest.approx(0.5 * (res["shares_end"][2] + res["shares_end"][3]))


def test_free_cash_flow_definition(results):
    res = results["Base"]
    assert res["fcf_margin"][3] == pytest.approx(res["fcf"][3] / res["revenue"][3])
    assert res["fcf_conv"][3] == pytest.approx(res["fcf"][3] / res["ebitda"][3])


def test_extreme_growth_inputs_still_balance(inputs):
    drv = ScenarioDrivers(
        growth=((0.4,) * 5, (-0.3,) * 5, (0.0,) * 5),
        gross_margin=(0.30,) * 5,
        sga_pct=(0.2,) * 5,
        capex_pct=(0.1,) * 5,
        dso=(90.0,) * 5,
        dio=(120.0,) * 5,
    )
    res = run_model(replace(inputs, scenarios={**inputs.scenarios, "Base": drv}))
    assert all(abs(v) < 1e-8 for v in res["bs_check"])
    assert min(res["revolver"]) >= 0
