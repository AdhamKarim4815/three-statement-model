"""Row specifications for the workbook sheets.

Each sheet is described declaratively as a list of :class:`Row` objects. The
workbook writer (:mod:`three_statement.excel`) first assigns a worksheet row
number to every spec (so formulas can reference rows on any sheet) and then
renders values and formulas. Every row that has a counterpart in the Python
engine carries an ``eng`` key, which is how cached values are produced and how
the Excel-versus-Python reconciliation test finds the cells to compare.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import Any

from xlsxwriter.utility import xl_col_to_name

from .engine import ModelResult
from .inputs import SCENARIOS, Inputs

SHEET_NAMES = {
    "cover": "Cover",
    "summary": "Summary",
    "assump": "Assumptions",
    "is": "Income Statement",
    "bs": "Balance Sheet",
    "cf": "Cash Flow",
    "sched": "Schedules",
    "checks": "Checks",
    "sens": "Sensitivity",
}

# Workbook-level defined names (all point at Assumptions scalars or the Checks flag)
NAME_SCENARIO = "Scenario"
NAME_INDEX = "Scenario_Idx"
NAME_DAYS = "Days_In_Year"
NAME_MIN_CASH = "Min_Cash"
NAME_REV_LIMIT = "Revolver_Limit"
NAME_TOL = "Tolerance"
NAME_CHECKS_OK = "Model_Checks_OK"
SCALAR_NAMES = (NAME_DAYS, NAME_MIN_CASH, NAME_REV_LIMIT, NAME_TOL, NAME_INDEX, NAME_CHECKS_OK)

HEADER_ROW = 5  # 1-based worksheet row holding the year headers
FIRST_DATA_ROW = 6
FIRST_COL = 2  # column C


@dataclass(frozen=True)
class Col:
    """A timeline column: ``i`` runs 0..(n_hist + n_forecast - 1)."""

    i: int
    n_hist: int

    @property
    def letter(self) -> str:
        return xl_col_to_name(FIRST_COL + self.i)

    @property
    def prev(self) -> str:
        if self.i == 0:
            raise ValueError("first column has no previous column")
        return xl_col_to_name(FIRST_COL + self.i - 1)

    @property
    def first(self) -> bool:
        return self.i == 0

    @property
    def hist(self) -> bool:
        return self.i < self.n_hist

    @property
    def t(self) -> int:
        return self.i - self.n_hist


@dataclass
class Row:
    """One worksheet row. ``hist``/``fc`` give per-column content, ``both`` applies to all."""

    key: str | None
    label: str = ""
    unit: str = ""
    fmt: str = "num"
    kind: str = "calc"  # calc, total, grand, section, blank, text, scalar, sub
    hist: Callable[[Col], Any] | Sequence[float] | None = None
    fc: Callable[[Col], Any] | Sequence[float] | None = None
    both: Callable[[Col], Any] | None = None
    scalar: Any = None  # for kind="scalar": value, or callable(R) -> formula/value
    eng: str | None = None  # engine line used for cached values and reconciliation
    cache: Callable[[Col], Any] | None = None  # cached value when no engine line applies
    indent: int = 0
    note: str = ""
    name: str | None = None  # defined name (scalar rows)
    italic: bool = False
    xl: int = 0  # assigned worksheet row (1-based)

    @property
    def engine_key(self) -> str | None:
        return self.eng if self.eng is not None else self.key


class Refs:
    """Resolves row keys to worksheet references; tracks the sheet being written."""

    def __init__(self) -> None:
        self.rows: dict[str, dict[str, int]] = {}
        self.current = ""

    def row(self, sheet: str, key: str) -> int:
        return self.rows[sheet][key]

    def x(self, sheet: str, key: str, c: Col, prev: bool = False) -> str:
        ref = f"{c.prev if prev else c.letter}{self.row(sheet, key)}"
        return ref if sheet == self.current else f"{quote(sheet)}!{ref}"

    def span(self, sheet: str, key: str) -> str:
        """Absolute-column range across the whole timeline for ``key`` (used by charts)."""
        r = self.row(sheet, key)
        return f"{quote(sheet)}!$C${r}:$J${r}"


def quote(sheet: str) -> str:
    name = SHEET_NAMES[sheet]
    return f"'{name}'" if " " in name else name


@dataclass
class Ctx:
    inp: Inputs
    res: ModelResult
    scenario: str


# --------------------------------------------------------------------------- Assumptions


def _scenario_block(
    R: Refs,
    key: str,
    label: str,
    unit: str,
    fmt: str,
    values: Sequence[Sequence[float]],
    hist: Callable[[Col], Any] | None,
    eng: str,
    note: str,
) -> list[Row]:
    def live(c: Col) -> str:
        cells = ",".join(f"{c.letter}{R.row('assump', f'{key}_{s}')}" for s in ("base", "up", "down"))
        return f"=CHOOSE({NAME_INDEX},{cells})"

    rows = [Row(key, label, unit, fmt, hist=hist, fc=live, eng=eng, note=note)]
    for suffix, name, vals in zip(("base", "up", "down"), SCENARIOS, values, strict=True):
        rows.append(Row(f"{key}_{suffix}", name, unit, fmt, kind="sub", fc=vals, indent=2))
    return rows


def assumptions_rows(R: Refs, ctx: Ctx) -> list[Row]:
    inp = ctx.inp
    x = R.x
    idx = SCENARIOS.index(ctx.scenario) + 1
    rows: list[Row] = [
        Row(None, "Model controls", kind="section"),
        Row(
            "scenario",
            "Active scenario (choose from list)",
            "list",
            "text",
            kind="scalar",
            scalar=ctx.scenario,
            name=NAME_SCENARIO,
            note="Single dropdown: every sheet, chart and ratio recalculates.",
        ),
        Row(
            "scen_idx",
            "Scenario index",
            "#",
            "int",
            kind="scalar",
            scalar=lambda R: (
                f"=MATCH(C{R.row('assump', 'scenario')},C{R.row('assump', 'sc1')}:C{R.row('assump', 'sc3')},0)"
            ),
            name=NAME_INDEX,
            cache=lambda c: idx,
            note="Position of the active scenario in the list below; drives CHOOSE().",
        ),
        Row("sc1", "Scenario list (1)", "", "text", kind="scalar", scalar=SCENARIOS[0], indent=1),
        Row("sc2", "Scenario list (2)", "", "text", kind="scalar", scalar=SCENARIOS[1], indent=1),
        Row("sc3", "Scenario list (3)", "", "text", kind="scalar", scalar=SCENARIOS[2], indent=1),
        Row(
            "days",
            "Days in year (working capital)",
            "days",
            "int",
            kind="scalar",
            scalar=inp.days_in_year,
            name=NAME_DAYS,
        ),
        Row(
            "min_cash",
            "Minimum cash balance",
            "$m",
            "num",
            kind="scalar",
            scalar=inp.min_cash,
            name=NAME_MIN_CASH,
            note="The revolver draws to defend this level and repays from any excess.",
        ),
        Row(
            "rev_limit",
            "Revolver commitment",
            "$m",
            "num",
            kind="scalar",
            scalar=inp.revolver_limit,
            name=NAME_REV_LIMIT,
        ),
        Row("tol", "Check tolerance", "$m", "tol", kind="scalar", scalar=inp.tolerance, name=NAME_TOL),
        Row(
            None,
            "Interest basis: opening balances (no circular references). The Python "
            "engine also offers an average-balance mode solved iteratively.",
            kind="text",
        ),
        Row(None, kind="blank"),
        Row(None, "Scenario drivers (first row of each block is live; edit the three rows beneath)", kind="section"),
    ]
    for k, seg in enumerate(inp.segments, 1):
        rows += _scenario_block(
            R,
            f"g_s{k}",
            f"Revenue growth: {seg}",
            "%",
            "pct",
            [inp.scenarios[s].growth[k - 1] for s in SCENARIOS],
            hist=lambda c, k=k: (
                None if c.first else (f"={x('is', f'rev_s{k}', c)}/{x('is', f'rev_s{k}', c, prev=True)}-1")
            ),
            eng=f"g_s{k}",
            note="Year-over-year growth for the segment.",
        )
    rows += _scenario_block(
        R,
        "gm",
        "Gross margin (before D&A)",
        "%",
        "pct",
        [inp.scenarios[s].gross_margin for s in SCENARIOS],
        hist=lambda c: f"={x('is', 'gross_margin', c)}",
        eng="gross_margin",
        note="COGS excludes depreciation and amortization.",
    )
    rows += _scenario_block(
        R,
        "sga_pct",
        "SG&A as % of revenue",
        "%",
        "pct",
        [inp.scenarios[s].sga_pct for s in SCENARIOS],
        hist=lambda c: f"={x('is', 'sga', c)}/{x('is', 'revenue', c)}",
        eng="sga_pct",
        note="Downside deleverages: fixed costs do not flex with revenue.",
    )
    rows += _scenario_block(
        R,
        "capex_pct",
        "Capex as % of revenue",
        "%",
        "pct",
        [inp.scenarios[s].capex_pct for s in SCENARIOS],
        hist=lambda c: None if c.first else f"={x('sched', 'capex', c)}/{x('is', 'revenue', c)}",
        eng="capex_pct",
        note="Historical capex is implied from the PP&E roll-forward.",
    )
    rows += _scenario_block(
        R,
        "dso",
        "Days sales outstanding (DSO)",
        "days",
        "days",
        [inp.scenarios[s].dso for s in SCENARIOS],
        hist=lambda c: f"={x('bs', 'ar', c)}/{x('is', 'revenue', c)}*{NAME_DAYS}",
        eng="dso",
        note="Receivables = DSO / days in year x revenue.",
    )
    rows += _scenario_block(
        R,
        "dio",
        "Days inventory outstanding (DIO)",
        "days",
        "days",
        [inp.scenarios[s].dio for s in SCENARIOS],
        hist=lambda c: f"={x('bs', 'inventory', c)}/{x('is', 'cogs', c)}*{NAME_DAYS}",
        eng="dio",
        note="Inventory = DIO / days in year x COGS.",
    )
    rows += [
        Row(None, kind="blank"),
        Row(None, "Other operating drivers (all scenarios)", kind="section"),
        Row(
            "rnd_pct",
            "R&D as % of revenue",
            "%",
            "pct",
            hist=lambda c: f"={x('is', 'rnd', c)}/{x('is', 'revenue', c)}",
            fc=inp.rnd_pct,
            eng="rnd_pct",
        ),
        Row(
            "oca_pct",
            "Prepaid and other current assets, % of revenue",
            "%",
            "pct",
            hist=lambda c: f"={x('bs', 'oca', c)}/{x('is', 'revenue', c)}",
            fc=inp.oca_pct,
            eng="oca_pct",
        ),
        Row(
            "accrued_pct",
            "Accrued liabilities, % of revenue",
            "%",
            "pct",
            hist=lambda c: f"={x('bs', 'accrued', c)}/{x('is', 'revenue', c)}",
            fc=inp.accrued_pct,
            eng="accrued_pct",
        ),
        Row(
            "dpo",
            "Days payables outstanding (DPO)",
            "days",
            "days",
            hist=lambda c: f"={x('bs', 'ap', c)}/{x('is', 'cogs', c)}*{NAME_DAYS}",
            fc=inp.dpo,
            eng="dpo",
            note="Payables = DPO / days in year x COGS.",
        ),
        Row(
            "da_rate",
            "D&A as % of opening net PP&E",
            "%",
            "pct",
            hist=lambda c: None if c.first else f"={x('is', 'da', c)}/{x('bs', 'ppe', c, prev=True)}",
            fc=inp.da_rate,
            eng="da_rate",
        ),
        Row(
            "tax_rate",
            "Tax rate (on positive pre-tax income)",
            "%",
            "pct",
            hist=lambda c: f"={x('is', 'eff_tax', c)}",
            fc=inp.tax_rate,
            eng="eff_tax",
            note="Simplified: no deferred taxes, no loss carry-forwards.",
        ),
        Row(None, kind="blank"),
        Row(None, "Financing", kind="section"),
        Row("tl_rate", "Term loan interest rate", "%", "pct2", fc=inp.tl_rate, note="Applied to the opening balance."),
        Row(
            "revolver_rate",
            "Revolver interest rate",
            "%",
            "pct2",
            fc=inp.revolver_rate,
            note="Applied to the opening balance.",
        ),
        Row(
            "cash_rate",
            "Interest rate earned on cash",
            "%",
            "pct2",
            fc=inp.cash_rate,
            note="Applied to the opening balance.",
        ),
        Row(
            "tl_amort",
            "Term loan mandatory amortization",
            "$m",
            "num",
            hist=lambda c: None if c.first else f"={x('sched', 'tl_amort', c)}",
            fc=inp.tl_amort,
            eng="tl_amort",
        ),
        Row(None, kind="blank"),
        Row(None, "Capital returns and issuance", kind="section"),
        Row(
            "payout",
            "Dividend payout ratio (of positive net income)",
            "%",
            "pct",
            hist=lambda c: None if c.first else (f"={x('sched', 'dividends', c)}/{x('is', 'net_income', c)}"),
            fc=inp.payout,
            eng="payout",
        ),
        Row(
            "buybacks",
            "Share buybacks",
            "$m",
            "num",
            hist=lambda c: None if c.first else f"={x('sched', 'buybacks', c)}",
            fc=inp.buybacks,
            eng="buybacks",
        ),
        Row(
            "issuance",
            "Stock issuance (optional; set to zero to switch off)",
            "$m",
            "num",
            hist=lambda c: None if c.first else f"={x('sched', 'issuance', c)}",
            fc=inp.issuance,
            eng="issuance",
        ),
        Row("share_price", "Average price for buybacks and issuance", "$/sh", "usd", fc=inp.share_price),
    ]
    return rows


# --------------------------------------------------------------------------- Income statement


def income_statement_rows(R: Refs, ctx: Ctx) -> list[Row]:
    inp, h = ctx.inp, ctx.inp.hist
    x = R.x
    n_seg = len(inp.segments)
    rows: list[Row] = [Row(None, "Revenue", kind="section")]
    for k, seg in enumerate(inp.segments, 1):
        rows.append(
            Row(
                f"rev_s{k}",
                seg,
                "$m",
                hist=h.segment_revenue[k - 1],
                fc=lambda c, k=k: f"={x('is', f'rev_s{k}', c, prev=True)}*(1+{x('assump', f'g_s{k}', c)})",
                indent=1,
            )
        )
    first, last = "rev_s1", f"rev_s{n_seg}"
    rows += [
        Row(
            "revenue",
            "Total revenue",
            "$m",
            kind="total",
            both=lambda c: f"=SUM({c.letter}{R.row('is', first)}:{c.letter}{R.row('is', last)})",
        ),
        Row(
            "rev_growth",
            "% growth",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: None if c.first else (f"={x('is', 'revenue', c)}/{x('is', 'revenue', c, prev=True)}-1"),
        ),
        Row(None, kind="blank"),
        Row(
            "cogs",
            "Cost of goods sold (excl. D&A)",
            "$m",
            hist=h.cogs,
            fc=lambda c: f"={x('is', 'revenue', c)}*(1-{x('assump', 'gm', c)})",
        ),
        Row(
            "gross_profit",
            "Gross profit",
            "$m",
            kind="total",
            both=lambda c: f"={x('is', 'revenue', c)}-{x('is', 'cogs', c)}",
        ),
        Row(
            "gross_margin",
            "% margin",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: f"={x('is', 'gross_profit', c)}/{x('is', 'revenue', c)}",
        ),
        Row(
            "sga",
            "Selling, general and administrative",
            "$m",
            hist=h.sga,
            fc=lambda c: f"={x('is', 'revenue', c)}*{x('assump', 'sga_pct', c)}",
        ),
        Row(
            "rnd",
            "Research and development",
            "$m",
            hist=h.rnd,
            fc=lambda c: f"={x('is', 'revenue', c)}*{x('assump', 'rnd_pct', c)}",
        ),
        Row(
            "ebitda",
            "EBITDA",
            "$m",
            kind="total",
            both=lambda c: f"={x('is', 'gross_profit', c)}-{x('is', 'sga', c)}-{x('is', 'rnd', c)}",
        ),
        Row(
            "ebitda_margin",
            "% margin",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: f"={x('is', 'ebitda', c)}/{x('is', 'revenue', c)}",
        ),
        Row("da", "Depreciation and amortization", "$m", hist=h.da, fc=lambda c: f"={x('sched', 'da', c)}"),
        Row(
            "ebit",
            "EBIT (operating income)",
            "$m",
            kind="total",
            both=lambda c: f"={x('is', 'ebitda', c)}-{x('is', 'da', c)}",
        ),
        Row(
            "ebit_margin",
            "% margin",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: f"={x('is', 'ebit', c)}/{x('is', 'revenue', c)}",
        ),
        Row("int_exp", "Interest expense", "$m", hist=h.int_exp, fc=lambda c: f"={x('sched', 'int_exp', c)}"),
        Row("int_inc", "Interest income", "$m", hist=h.int_inc, fc=lambda c: f"={x('sched', 'int_inc', c)}"),
        Row(
            "ebt",
            "Pre-tax income",
            "$m",
            kind="total",
            both=lambda c: f"={x('is', 'ebit', c)}-{x('is', 'int_exp', c)}+{x('is', 'int_inc', c)}",
        ),
        Row(
            "tax",
            "Income tax expense",
            "$m",
            hist=h.tax,
            fc=lambda c: f"=MAX(0,{x('is', 'ebt', c)})*{x('assump', 'tax_rate', c)}",
        ),
        Row("net_income", "Net income", "$m", kind="grand", both=lambda c: f"={x('is', 'ebt', c)}-{x('is', 'tax', c)}"),
        Row(
            "net_margin",
            "% margin",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: f"={x('is', 'net_income', c)}/{x('is', 'revenue', c)}",
        ),
        Row(
            "eff_tax",
            "Effective tax rate",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            both=lambda c: f"=IF({x('is', 'ebt', c)}>0,{x('is', 'tax', c)}/{x('is', 'ebt', c)},0)",
        ),
        Row(None, kind="blank"),
        Row(None, "Per share data", kind="section"),
        Row(
            "shares_wavg",
            "Weighted average shares outstanding",
            "m",
            "num",
            both=lambda c: f"={x('sched', 'shares_wavg', c)}",
        ),
        Row(
            "eps",
            "Earnings per share",
            "$",
            "usd",
            both=lambda c: f"={x('is', 'net_income', c)}/{x('is', 'shares_wavg', c)}",
        ),
        Row(
            "dps",
            "Dividends per share",
            "$",
            "usd",
            both=lambda c: None if c.first else (f"={x('sched', 'dividends', c)}/{x('is', 'shares_wavg', c)}"),
        ),
    ]
    return rows


# --------------------------------------------------------------------------- Balance sheet


def balance_sheet_rows(R: Refs, ctx: Ctx) -> list[Row]:
    h = ctx.inp.hist
    x = R.x

    def sched(key: str, skey: str | None = None) -> Callable[[Col], str]:
        return lambda c: f"={x('sched', skey or key, c)}"

    def hold(key: str) -> Callable[[Col], str]:
        return lambda c: f"={x('bs', key, c, prev=True)}"

    return [
        Row(None, "Assets", kind="section"),
        Row("cash", "Cash and equivalents", "$m", hist=h.cash, fc=lambda c: f"={x('cf', 'cash_end', c)}", indent=1),
        Row("ar", "Accounts receivable", "$m", hist=h.ar, fc=sched("ar"), indent=1),
        Row("inventory", "Inventory", "$m", hist=h.inventory, fc=sched("inventory"), indent=1),
        Row("oca", "Prepaid and other current assets", "$m", hist=h.oca, fc=sched("oca"), indent=1),
        Row(
            "total_current_assets",
            "Total current assets",
            "$m",
            kind="total",
            both=lambda c: f"=SUM({c.letter}{R.row('bs', 'cash')}:{c.letter}{R.row('bs', 'oca')})",
        ),
        Row("ppe", "Net property, plant and equipment", "$m", hist=h.ppe, fc=sched("ppe", "ppe_close"), indent=1),
        Row("goodwill", "Goodwill and intangibles", "$m", hist=h.goodwill, fc=hold("goodwill"), indent=1),
        Row(
            "total_assets",
            "Total assets",
            "$m",
            kind="grand",
            both=lambda c: f"={x('bs', 'total_current_assets', c)}+{x('bs', 'ppe', c)}+{x('bs', 'goodwill', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Liabilities", kind="section"),
        Row("ap", "Accounts payable", "$m", hist=h.ap, fc=sched("ap"), indent=1),
        Row("accrued", "Accrued liabilities", "$m", hist=h.accrued, fc=sched("accrued"), indent=1),
        Row(
            "revolver", "Revolving credit facility", "$m", hist=h.revolver, fc=sched("revolver", "rev_close"), indent=1
        ),
        Row("term_loan", "Term loan", "$m", hist=h.term_loan, fc=sched("term_loan", "tl_close"), indent=1),
        Row("other_ltl", "Other long-term liabilities", "$m", hist=h.other_ltl, fc=hold("other_ltl"), indent=1),
        Row(
            "total_liabilities",
            "Total liabilities",
            "$m",
            kind="total",
            both=lambda c: f"=SUM({c.letter}{R.row('bs', 'ap')}:{c.letter}{R.row('bs', 'other_ltl')})",
        ),
        Row(None, kind="blank"),
        Row(None, "Shareholders' equity", kind="section"),
        Row("apic", "Common stock and paid-in capital", "$m", hist=h.apic, fc=sched("apic", "apic_close"), indent=1),
        Row("retained", "Retained earnings", "$m", hist=h.retained, fc=sched("retained", "re_close"), indent=1),
        Row("treasury", "Treasury stock", "$m", hist=h.treasury, fc=sched("treasury", "tre_close"), indent=1),
        Row(
            "total_equity",
            "Total shareholders' equity",
            "$m",
            kind="total",
            both=lambda c: f"=SUM({c.letter}{R.row('bs', 'apic')}:{c.letter}{R.row('bs', 'treasury')})",
        ),
        Row(
            "total_le",
            "Total liabilities and equity",
            "$m",
            kind="grand",
            both=lambda c: f"={x('bs', 'total_liabilities', c)}+{x('bs', 'total_equity', c)}",
        ),
        Row(None, kind="blank"),
        Row(
            "bs_check",
            "Balance check (assets less liabilities and equity)",
            "$m",
            "num",
            italic=True,
            both=lambda c: f"=ROUND({x('bs', 'total_assets', c)}-{x('bs', 'total_le', c)},6)",
        ),
    ]


# --------------------------------------------------------------------------- Cash flow


def cash_flow_rows(R: Refs, ctx: Ctx) -> list[Row]:
    x = R.x

    def d(key: str, sign: str, bs_key: str) -> Callable[[Col], str]:
        return lambda c: f"={sign}({x('bs', bs_key, c)}-{x('bs', bs_key, c, prev=True)})"

    return [
        Row(None, "Operating activities", kind="section"),
        Row("cf_ni", "Net income", "$m", fc=lambda c: f"={x('is', 'net_income', c)}", indent=1),
        Row("cf_da", "Depreciation and amortization", "$m", fc=lambda c: f"={x('is', 'da', c)}", indent=1),
        Row("cf_ar", "(Increase) / decrease in accounts receivable", "$m", fc=d("cf_ar", "-", "ar"), indent=1),
        Row("cf_inv", "(Increase) / decrease in inventory", "$m", fc=d("cf_inv", "-", "inventory"), indent=1),
        Row("cf_oca", "(Increase) / decrease in other current assets", "$m", fc=d("cf_oca", "-", "oca"), indent=1),
        Row("cf_ap", "Increase / (decrease) in accounts payable", "$m", fc=d("cf_ap", "+", "ap"), indent=1),
        Row("cf_accr", "Increase / (decrease) in accrued liabilities", "$m", fc=d("cf_accr", "+", "accrued"), indent=1),
        Row(
            "cfo",
            "Cash flow from operations",
            "$m",
            kind="total",
            fc=lambda c: f"=SUM({c.letter}{R.row('cf', 'cf_ni')}:{c.letter}{R.row('cf', 'cf_accr')})",
        ),
        Row(None, kind="blank"),
        Row(None, "Investing activities", kind="section"),
        Row("cf_capex", "Capital expenditures", "$m", fc=lambda c: f"=-{x('sched', 'capex', c)}", indent=1),
        Row("cfi", "Cash flow from investing", "$m", kind="total", fc=lambda c: f"={x('cf', 'cf_capex', c)}"),
        Row(None, kind="blank"),
        Row(None, "Financing activities", kind="section"),
        Row("cf_tl", "Term loan mandatory amortization", "$m", fc=lambda c: f"=-{x('sched', 'tl_amort', c)}", indent=1),
        Row("cf_rev", "Revolver draw / (repayment)", "$m", fc=lambda c: f"={x('sched', 'rev_change', c)}", indent=1),
        Row("cf_issue", "Stock issuance", "$m", fc=lambda c: f"={x('sched', 'issuance', c)}", indent=1),
        Row("cf_div", "Dividends paid", "$m", fc=lambda c: f"=-{x('sched', 'dividends', c)}", indent=1),
        Row("cf_bb", "Share buybacks", "$m", fc=lambda c: f"=-{x('sched', 'buybacks', c)}", indent=1),
        Row(
            "cff",
            "Cash flow from financing",
            "$m",
            kind="total",
            fc=lambda c: f"=SUM({c.letter}{R.row('cf', 'cf_tl')}:{c.letter}{R.row('cf', 'cf_bb')})",
        ),
        Row(None, kind="blank"),
        Row(
            "net_change",
            "Net change in cash",
            "$m",
            kind="total",
            fc=lambda c: f"={x('cf', 'cfo', c)}+{x('cf', 'cfi', c)}+{x('cf', 'cff', c)}",
        ),
        Row("cash_begin", "Cash, beginning of year", "$m", fc=lambda c: f"={x('bs', 'cash', c, prev=True)}"),
        Row(
            "cash_end",
            "Cash, end of year",
            "$m",
            kind="grand",
            fc=lambda c: f"={x('cf', 'cash_begin', c)}+{x('cf', 'net_change', c)}",
        ),
        Row(None, kind="blank"),
        Row(
            "fcf",
            "Free cash flow (CFO less capex)",
            "$m",
            kind="total",
            fc=lambda c: f"={x('cf', 'cfo', c)}+{x('cf', 'cf_capex', c)}",
        ),
        Row(None, "Historical cash flows are not presented: the model starts from three balance sheets.", kind="text"),
    ]


# --------------------------------------------------------------------------- Schedules


def schedules_rows(R: Refs, ctx: Ctx) -> list[Row]:
    inp, h = ctx.inp, ctx.inp.hist
    x = R.x
    min_cash = inp.min_cash

    def open_(close: str) -> Callable[[Col], str | None]:
        return lambda c: None if c.first else f"={x('sched', close, c, prev=True)}"

    def link(sheet: str, key: str) -> Callable[[Col], str]:
        return lambda c: f"={x(sheet, key, c)}"

    return [
        Row(None, "PP&E roll-forward", kind="section"),
        Row("ppe_open", "Opening net PP&E", "$m", both=open_("ppe_close"), eng="ppe_open"),
        Row(
            "capex",
            "Plus: capital expenditures",
            "$m",
            indent=1,
            hist=lambda c: (
                None
                if c.first
                else (f"={x('sched', 'ppe_close', c)}-{x('sched', 'ppe_open', c)}+{x('sched', 'da', c)}")
            ),
            fc=lambda c: f"={x('assump', 'capex_pct', c)}*{x('is', 'revenue', c)}",
        ),
        Row(
            "da",
            "Less: depreciation and amortization",
            "$m",
            indent=1,
            hist=link("is", "da"),
            fc=lambda c: f"={x('assump', 'da_rate', c)}*{x('sched', 'ppe_open', c)}",
        ),
        Row(
            "ppe_close",
            "Closing net PP&E",
            "$m",
            kind="total",
            eng="ppe",
            hist=link("bs", "ppe"),
            fc=lambda c: f"={x('sched', 'ppe_open', c)}+{x('sched', 'capex', c)}-{x('sched', 'da', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Working capital (days based)", kind="section"),
        Row(
            "ar",
            "Accounts receivable",
            "$m",
            indent=1,
            hist=link("bs", "ar"),
            fc=lambda c: f"={x('assump', 'dso', c)}/{NAME_DAYS}*{x('is', 'revenue', c)}",
        ),
        Row(
            "inventory",
            "Inventory",
            "$m",
            indent=1,
            hist=link("bs", "inventory"),
            fc=lambda c: f"={x('assump', 'dio', c)}/{NAME_DAYS}*{x('is', 'cogs', c)}",
        ),
        Row(
            "oca",
            "Prepaid and other current assets",
            "$m",
            indent=1,
            hist=link("bs", "oca"),
            fc=lambda c: f"={x('assump', 'oca_pct', c)}*{x('is', 'revenue', c)}",
        ),
        Row(
            "ap",
            "Accounts payable",
            "$m",
            indent=1,
            hist=link("bs", "ap"),
            fc=lambda c: f"={x('assump', 'dpo', c)}/{NAME_DAYS}*{x('is', 'cogs', c)}",
        ),
        Row(
            "accrued",
            "Accrued liabilities",
            "$m",
            indent=1,
            hist=link("bs", "accrued"),
            fc=lambda c: f"={x('assump', 'accrued_pct', c)}*{x('is', 'revenue', c)}",
        ),
        Row(
            "nwc",
            "Net working capital",
            "$m",
            kind="total",
            both=lambda c: (
                f"={x('sched', 'ar', c)}+{x('sched', 'inventory', c)}+{x('sched', 'oca', c)}"
                f"-{x('sched', 'ap', c)}-{x('sched', 'accrued', c)}"
            ),
        ),
        Row(
            "d_nwc",
            "Increase / (decrease) in net working capital",
            "$m",
            indent=1,
            both=lambda c: None if c.first else (f"={x('sched', 'nwc', c)}-{x('sched', 'nwc', c, prev=True)}"),
        ),
        Row(None, kind="blank"),
        Row(None, "Term loan", kind="section"),
        Row("tl_open", "Opening balance", "$m", both=open_("tl_close")),
        Row(
            "tl_amort",
            "Less: mandatory amortization",
            "$m",
            indent=1,
            hist=lambda c: None if c.first else (f"={x('sched', 'tl_open', c)}-{x('sched', 'tl_close', c)}"),
            fc=lambda c: f"=MIN({x('sched', 'tl_open', c)},{x('assump', 'tl_amort', c)})",
        ),
        Row(
            "tl_close",
            "Closing balance",
            "$m",
            kind="total",
            eng="term_loan",
            hist=link("bs", "term_loan"),
            fc=lambda c: f"={x('sched', 'tl_open', c)}-{x('sched', 'tl_amort', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Cash available before revolver", kind="section"),
        Row("cash_open", "Opening cash", "$m", indent=1, fc=lambda c: f"={x('bs', 'cash', c, prev=True)}"),
        Row("s_cfo", "Plus: cash flow from operations", "$m", indent=1, eng="cfo", fc=link("cf", "cfo")),
        Row("s_cfi", "Plus: cash flow from investing", "$m", indent=1, eng="cfi", fc=link("cf", "cfi")),
        Row(
            "s_amort",
            "Less: term loan amortization",
            "$m",
            indent=1,
            eng="cf_tl",
            fc=lambda c: f"=-{x('sched', 'tl_amort', c)}",
        ),
        Row(
            "s_issue",
            "Plus: stock issuance",
            "$m",
            indent=1,
            eng="issuance",
            fc=lambda c: f"={x('sched', 'issuance', c)}",
        ),
        Row("s_div", "Less: dividends", "$m", indent=1, eng="cf_div", fc=lambda c: f"=-{x('sched', 'dividends', c)}"),
        Row("s_bb", "Less: share buybacks", "$m", indent=1, eng="cf_bb", fc=lambda c: f"=-{x('sched', 'buybacks', c)}"),
        Row(
            "cash_pre",
            "Cash before revolver",
            "$m",
            kind="total",
            fc=lambda c: f"=SUM({c.letter}{R.row('sched', 'cash_open')}:{c.letter}{R.row('sched', 's_bb')})",
        ),
        Row(
            "s_min", "Minimum cash balance", "$m", indent=1, fc=lambda c: f"={NAME_MIN_CASH}", cache=lambda c: min_cash
        ),
        Row(
            "s_surplus",
            "Surplus / (shortfall) versus minimum",
            "$m",
            indent=1,
            fc=lambda c: f"={x('sched', 'cash_pre', c)}-{x('sched', 's_min', c)}",
            cache=lambda c: ctx.res["cash_pre"][c.i] - min_cash,
        ),
        Row(None, kind="blank"),
        Row(None, "Revolving credit facility", kind="section"),
        Row("rev_open", "Opening balance", "$m", both=open_("rev_close")),
        Row(
            "rev_change",
            "Draw / (repayment)",
            "$m",
            indent=1,
            hist=lambda c: None if c.first else (f"={x('sched', 'rev_close', c)}-{x('sched', 'rev_open', c)}"),
            fc=lambda c: (
                f"=MAX(-{x('sched', 'rev_open', c)},MIN({x('sched', 's_min', c)}"
                f"-{x('sched', 'cash_pre', c)},{NAME_REV_LIMIT}-{x('sched', 'rev_open', c)}))"
            ),
        ),
        Row(
            "rev_close",
            "Closing balance",
            "$m",
            kind="total",
            eng="revolver",
            hist=link("bs", "revolver"),
            fc=lambda c: f"={x('sched', 'rev_open', c)}+{x('sched', 'rev_change', c)}",
        ),
        Row(
            "cash_close",
            "Cash after revolver",
            "$m",
            kind="total",
            fc=lambda c: f"={x('sched', 'cash_pre', c)}+{x('sched', 'rev_change', c)}",
        ),
        Row(None, kind="blank"),
        Row(
            "total_debt",
            "Total debt",
            "$m",
            both=lambda c: f"={x('sched', 'tl_close', c)}+{x('sched', 'rev_close', c)}",
        ),
        Row("net_debt", "Net debt", "$m", both=lambda c: f"={x('sched', 'total_debt', c)}-{x('bs', 'cash', c)}"),
        Row(None, kind="blank"),
        Row(None, "Interest (opening balances)", kind="section"),
        Row(
            "int_tl",
            "Term loan interest expense",
            "$m",
            indent=1,
            fc=lambda c: f"={x('assump', 'tl_rate', c)}*{x('sched', 'tl_open', c)}",
        ),
        Row(
            "int_rev",
            "Revolver interest expense",
            "$m",
            indent=1,
            fc=lambda c: f"={x('assump', 'revolver_rate', c)}*{x('sched', 'rev_open', c)}",
        ),
        Row(
            "int_exp",
            "Total interest expense",
            "$m",
            kind="total",
            fc=lambda c: f"={x('sched', 'int_tl', c)}+{x('sched', 'int_rev', c)}",
        ),
        Row(
            "int_inc",
            "Interest income on cash",
            "$m",
            fc=lambda c: f"={x('assump', 'cash_rate', c)}*{x('sched', 'cash_open', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Retained earnings", kind="section"),
        Row("re_open", "Opening balance", "$m", both=open_("re_close")),
        Row(
            "re_ni",
            "Plus: net income",
            "$m",
            indent=1,
            eng="net_income",
            both=lambda c: None if c.first else f"={x('is', 'net_income', c)}",
        ),
        Row(
            "dividends",
            "Less: dividends",
            "$m",
            indent=1,
            hist=lambda c: (
                None
                if c.first
                else (f"={x('sched', 're_open', c)}+{x('sched', 're_ni', c)}-{x('sched', 're_close', c)}")
            ),
            fc=lambda c: f"={x('assump', 'payout', c)}*MAX(0,{x('sched', 're_ni', c)})",
        ),
        Row(
            "re_close",
            "Closing balance",
            "$m",
            kind="total",
            eng="retained",
            hist=link("bs", "retained"),
            fc=lambda c: f"={x('sched', 're_open', c)}+{x('sched', 're_ni', c)}-{x('sched', 'dividends', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Paid-in capital and treasury stock", kind="section"),
        Row("apic_open", "Opening paid-in capital", "$m", both=open_("apic_close")),
        Row(
            "issuance",
            "Plus: stock issuance",
            "$m",
            indent=1,
            hist=lambda c: None if c.first else (f"={x('sched', 'apic_close', c)}-{x('sched', 'apic_open', c)}"),
            fc=lambda c: f"={x('assump', 'issuance', c)}",
        ),
        Row(
            "apic_close",
            "Closing paid-in capital",
            "$m",
            kind="total",
            eng="apic",
            hist=link("bs", "apic"),
            fc=lambda c: f"={x('sched', 'apic_open', c)}+{x('sched', 'issuance', c)}",
        ),
        Row("tre_open", "Opening treasury stock", "$m", both=open_("tre_close")),
        Row(
            "buybacks",
            "Less: share buybacks",
            "$m",
            indent=1,
            hist=lambda c: None if c.first else (f"={x('sched', 'tre_open', c)}-{x('sched', 'tre_close', c)}"),
            fc=lambda c: f"={x('assump', 'buybacks', c)}",
        ),
        Row(
            "tre_close",
            "Closing treasury stock",
            "$m",
            kind="total",
            eng="treasury",
            hist=link("bs", "treasury"),
            fc=lambda c: f"={x('sched', 'tre_open', c)}-{x('sched', 'buybacks', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Share count (millions)", kind="section"),
        Row(
            "sh_open",
            "Opening shares outstanding",
            "m",
            indent=1,
            eng="sh_open",
            fc=lambda c: f"={x('sched', 'shares_end', c, prev=True)}",
        ),
        Row(
            "sh_rep",
            "Less: shares repurchased",
            "m",
            indent=1,
            eng="sh_rep",
            fc=lambda c: f"={x('sched', 'buybacks', c)}/{x('assump', 'share_price', c)}",
        ),
        Row(
            "sh_iss",
            "Plus: shares issued",
            "m",
            indent=1,
            eng="sh_iss",
            fc=lambda c: f"={x('sched', 'issuance', c)}/{x('assump', 'share_price', c)}",
        ),
        Row(
            "shares_end",
            "Closing shares outstanding",
            "m",
            kind="total",
            hist=h.shares_end,
            fc=lambda c: f"={x('sched', 'sh_open', c)}-{x('sched', 'sh_rep', c)}+{x('sched', 'sh_iss', c)}",
        ),
        Row(
            "shares_wavg",
            "Weighted average shares",
            "m",
            hist=h.shares_wavg,
            fc=lambda c: f"=AVERAGE({x('sched', 'sh_open', c)},{x('sched', 'shares_end', c)})",
        ),
    ]


# --------------------------------------------------------------------------- Summary


def summary_rows(R: Refs, ctx: Ctx) -> list[Row]:
    x = R.x

    def link(sheet: str, key: str) -> Callable[[Col], str]:
        return lambda c: f"={x(sheet, key, c)}"

    def link1(sheet: str, key: str) -> Callable[[Col], str | None]:
        return lambda c: None if c.first else f"={x(sheet, key, c)}"

    return [
        Row(None, "Income statement", kind="section"),
        Row("revenue", "Revenue", "$m", both=link("is", "revenue")),
        Row("rev_growth", "% growth", "%", "pct", kind="sub", italic=True, indent=1, both=link1("is", "rev_growth")),
        Row("gross_profit", "Gross profit", "$m", both=link("is", "gross_profit")),
        Row("gross_margin", "% margin", "%", "pct", kind="sub", italic=True, indent=1, both=link("is", "gross_margin")),
        Row("ebitda", "EBITDA", "$m", both=link("is", "ebitda")),
        Row(
            "ebitda_margin", "% margin", "%", "pct", kind="sub", italic=True, indent=1, both=link("is", "ebitda_margin")
        ),
        Row("ebit_margin", "EBIT margin", "%", "pct", both=link("is", "ebit_margin")),
        Row("net_income", "Net income", "$m", both=link("is", "net_income")),
        Row("net_margin", "% margin", "%", "pct", kind="sub", italic=True, indent=1, both=link("is", "net_margin")),
        Row("eps", "Earnings per share", "$", "usd", both=link("is", "eps")),
        Row(None, kind="blank"),
        Row(None, "Cash flow", kind="section"),
        Row("cfo", "Cash flow from operations", "$m", fc=link("cf", "cfo")),
        Row("cf_capex", "Capital expenditures", "$m", fc=link("cf", "cf_capex")),
        Row("fcf", "Free cash flow", "$m", kind="total", fc=link("cf", "fcf")),
        Row(
            "fcf_margin",
            "FCF margin (% of revenue)",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            fc=lambda c: f"={x('summary', 'fcf', c)}/{x('is', 'revenue', c)}",
        ),
        Row(
            "fcf_conv",
            "FCF conversion (% of EBITDA)",
            "%",
            "pct",
            kind="sub",
            italic=True,
            indent=1,
            fc=lambda c: f"={x('summary', 'fcf', c)}/{x('summary', 'ebitda', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Balance sheet and credit", kind="section"),
        Row("cash", "Cash and equivalents", "$m", both=link("bs", "cash")),
        Row("total_debt", "Total debt", "$m", both=link("sched", "total_debt")),
        Row("net_debt", "Net debt", "$m", both=link("sched", "net_debt")),
        Row(
            "net_debt_ebitda",
            "Net debt / EBITDA",
            "x",
            "mult",
            both=lambda c: f"={x('summary', 'net_debt', c)}/{x('summary', 'ebitda', c)}",
        ),
        Row(
            "int_cov",
            "Interest coverage (EBIT / interest expense)",
            "x",
            "mult",
            both=lambda c: f'=IF({x("is", "int_exp", c)}>0,{x("is", "ebit", c)}/{x("is", "int_exp", c)},"n.m.")',
        ),
        Row(
            "invested_capital",
            "Invested capital (debt + equity - cash)",
            "$m",
            both=lambda c: f"={x('summary', 'total_debt', c)}+{x('bs', 'total_equity', c)}-{x('summary', 'cash', c)}",
        ),
        Row(None, kind="blank"),
        Row(None, "Returns (on average balances)", kind="section"),
        Row(
            "roe",
            "Return on equity",
            "%",
            "pct",
            both=lambda c: (
                None
                if c.first
                else (
                    f"={x('is', 'net_income', c)}/AVERAGE({x('bs', 'total_equity', c, prev=True)},"
                    f"{x('bs', 'total_equity', c)})"
                )
            ),
        ),
        Row(
            "roic",
            "Return on invested capital (NOPAT / invested capital)",
            "%",
            "pct",
            both=lambda c: (
                None
                if c.first
                else (
                    f"={x('is', 'ebit', c)}*(1-{x('is', 'eff_tax', c)})/AVERAGE("
                    f"{x('summary', 'invested_capital', c, prev=True)},"
                    f"{x('summary', 'invested_capital', c)})"
                )
            ),
        ),
    ]


# --------------------------------------------------------------------------- Checks

CHECK_LABELS = {
    "chk_bs": "Balance sheet balances (assets less liabilities and equity)",
    "chk_cash": "Balance sheet cash ties to cash flow statement ending cash",
    "chk_cash2": "Cash flow ending cash ties to debt schedule cash after revolver",
    "chk_re": "Retained earnings roll-forward (prior + net income - dividends)",
    "chk_debt": "Debt roll-forward (prior debt + financing flows)",
    "chk_ppe": "PP&E roll-forward (prior + capex - D&A)",
    "chk_equity": "Total equity roll-forward (net income, issuance, dividends, buybacks)",
    "chk_nwc": "Cash flow working capital lines tie to schedule change in NWC",
    "chk_mincash": "Minimum cash maintained (shortfall)",
    "chk_limit": "Revolver within commitment (excess)",
    "chk_negdebt": "No negative debt balances",
}


def checks_rows(R: Refs, ctx: Ctx) -> list[Row]:
    x = R.x

    def rnd(expr: Callable[[Col], str]) -> Callable[[Col], str]:
        return lambda c: f"=ROUND({expr(c)},6)"

    rows = [
        Row(None, "Integrity checks (each row should be zero in every year)", kind="section"),
        Row(
            "chk_bs",
            CHECK_LABELS["chk_bs"],
            "$m",
            "chk",
            both=rnd(lambda c: f"{x('bs', 'total_assets', c)}-{x('bs', 'total_le', c)}"),
        ),
        Row(
            "chk_cash",
            CHECK_LABELS["chk_cash"],
            "$m",
            "chk",
            fc=rnd(lambda c: f"{x('bs', 'cash', c)}-{x('cf', 'cash_end', c)}"),
        ),
        Row(
            "chk_cash2",
            CHECK_LABELS["chk_cash2"],
            "$m",
            "chk",
            fc=rnd(lambda c: f"{x('cf', 'cash_end', c)}-{x('sched', 'cash_close', c)}"),
        ),
        Row(
            "chk_re",
            CHECK_LABELS["chk_re"],
            "$m",
            "chk",
            fc=rnd(
                lambda c: (
                    f"{x('bs', 'retained', c)}-({x('bs', 'retained', c, prev=True)}"
                    f"+{x('is', 'net_income', c)}+{x('cf', 'cf_div', c)})"
                )
            ),
        ),
        Row(
            "chk_debt",
            CHECK_LABELS["chk_debt"],
            "$m",
            "chk",
            fc=rnd(
                lambda c: (
                    f"({x('bs', 'term_loan', c)}+{x('bs', 'revolver', c)})"
                    f"-({x('bs', 'term_loan', c, prev=True)}+{x('bs', 'revolver', c, prev=True)}"
                    f"+{x('cf', 'cf_tl', c)}+{x('cf', 'cf_rev', c)})"
                )
            ),
        ),
        Row(
            "chk_ppe",
            CHECK_LABELS["chk_ppe"],
            "$m",
            "chk",
            fc=rnd(
                lambda c: (
                    f"{x('bs', 'ppe', c)}-({x('bs', 'ppe', c, prev=True)}-{x('cf', 'cf_capex', c)}-{x('is', 'da', c)})"
                )
            ),
        ),
        Row(
            "chk_equity",
            CHECK_LABELS["chk_equity"],
            "$m",
            "chk",
            fc=rnd(
                lambda c: (
                    f"{x('bs', 'total_equity', c)}-({x('bs', 'total_equity', c, prev=True)}"
                    f"+{x('is', 'net_income', c)}+{x('cf', 'cf_issue', c)}"
                    f"+{x('cf', 'cf_div', c)}+{x('cf', 'cf_bb', c)})"
                )
            ),
        ),
        Row(
            "chk_nwc",
            CHECK_LABELS["chk_nwc"],
            "$m",
            "chk",
            fc=rnd(
                lambda c: f"SUM({x('cf', 'cf_ar', c)}:{x('cf', 'cf_accr', c).split('!')[-1]})+{x('sched', 'd_nwc', c)}"
            ),
        ),
        Row(
            "chk_mincash",
            CHECK_LABELS["chk_mincash"],
            "$m",
            "chk",
            fc=rnd(lambda c: f"MAX(0,{NAME_MIN_CASH}-{x('bs', 'cash', c)})"),
        ),
        Row(
            "chk_limit",
            CHECK_LABELS["chk_limit"],
            "$m",
            "chk",
            fc=rnd(lambda c: f"MAX(0,{x('bs', 'revolver', c)}-{NAME_REV_LIMIT})"),
        ),
        Row(
            "chk_negdebt",
            CHECK_LABELS["chk_negdebt"],
            "$m",
            "chk",
            both=rnd(lambda c: f"MAX(0,-{x('bs', 'revolver', c)})+MAX(0,-{x('bs', 'term_loan', c)})"),
        ),
    ]
    return rows
