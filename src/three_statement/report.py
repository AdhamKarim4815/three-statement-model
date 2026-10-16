"""Plain-text and machine-readable renderings of a :class:`ModelResult`."""

from __future__ import annotations

import csv
import io
import json

from .engine import ModelResult
from .inputs import Inputs
from .sensitivity import SensitivityGrid

# (label, engine key, format) with None key marking a section heading
Layout = list[tuple[str, str | None, str]]

STATEMENT_LAYOUT: Layout = [
    ("INCOME STATEMENT ($m)", None, ""),
    ("Total revenue", "revenue", "num"),
    ("  % growth", "rev_growth", "pct"),
    ("Gross profit", "gross_profit", "num"),
    ("  % margin", "gross_margin", "pct"),
    ("EBITDA", "ebitda", "num"),
    ("  % margin", "ebitda_margin", "pct"),
    ("Depreciation and amortization", "da", "num"),
    ("EBIT", "ebit", "num"),
    ("Interest expense", "int_exp", "num"),
    ("Interest income", "int_inc", "num"),
    ("Pre-tax income", "ebt", "num"),
    ("Income tax expense", "tax", "num"),
    ("Net income", "net_income", "num"),
    ("Earnings per share ($)", "eps", "usd"),
    ("BALANCE SHEET ($m)", None, ""),
    ("Cash", "cash", "num"),
    ("Accounts receivable", "ar", "num"),
    ("Inventory", "inventory", "num"),
    ("Other current assets", "oca", "num"),
    ("Net PP&E", "ppe", "num"),
    ("Goodwill and intangibles", "goodwill", "num"),
    ("Total assets", "total_assets", "num"),
    ("Accounts payable", "ap", "num"),
    ("Accrued liabilities", "accrued", "num"),
    ("Revolver", "revolver", "num"),
    ("Term loan", "term_loan", "num"),
    ("Other long-term liabilities", "other_ltl", "num"),
    ("Total liabilities", "total_liabilities", "num"),
    ("Total equity", "total_equity", "num"),
    ("Total liabilities and equity", "total_le", "num"),
    ("CASH FLOW STATEMENT ($m)", None, ""),
    ("Net income", "cf_ni", "num"),
    ("Depreciation and amortization", "cf_da", "num"),
    ("Change in working capital", "cf_nwc", "num"),
    ("Cash flow from operations", "cfo", "num"),
    ("Capital expenditures", "cf_capex", "num"),
    ("Term loan amortization", "cf_tl", "num"),
    ("Revolver draw / (repayment)", "cf_rev", "num"),
    ("Stock issuance", "cf_issue", "num"),
    ("Dividends", "cf_div", "num"),
    ("Share buybacks", "cf_bb", "num"),
    ("Net change in cash", "net_change", "num"),
    ("Ending cash", "cash_end", "num"),
    ("Free cash flow", "fcf", "num"),
    ("KEY RATIOS", None, ""),
    ("Net debt / EBITDA", "net_debt_ebitda", "mult"),
    ("Interest coverage (EBIT / interest)", "int_cov", "mult"),
    ("Return on equity", "roe", "pct"),
    ("Return on invested capital", "roic", "pct"),
]


def _value(result: ModelResult, key: str, i: int) -> float | None:
    if key == "cf_nwc":  # presentation line: sum of the five working capital items
        parts = [result[k][i] for k in ("cf_ar", "cf_inv", "cf_oca", "cf_ap", "cf_accr")]
        return None if any(p is None for p in parts) else float(sum(p for p in parts if p))
    return result[key][i]


def _fmt(value: float | None, kind: str) -> str:
    if value is None:
        return ""
    if kind == "pct":
        v = round(value * 100, 1) + 0.0
        return f"{v:.1f}%" if v >= 0 else f"({abs(v):.1f}%)"
    if kind == "mult":
        v = round(value, 2) + 0.0
        return f"{v:.2f}x" if v >= 0 else f"({abs(v):.2f}x)"
    if kind == "usd":
        return f"{round(value, 2) + 0.0:.2f}"
    v = round(value, 1) + 0.0
    return f"{v:,.1f}" if v >= 0 else f"({abs(v):,.1f})"


def format_statements(
    result: ModelResult, title: str = "Northwind Components Inc. (fictional)", forecast_only: bool = False
) -> str:
    """Render the three statements and key ratios as an aligned text table."""
    start = result.n_hist if forecast_only else 0
    idx = list(range(start, len(result.years)))
    heads = [f"FY{result.years[i]}{'A' if i < result.n_hist else 'E'}" for i in idx]
    width = 36
    col = 11
    lines = [
        f"{title}: scenario {result.scenario}, interest on {result.interest_basis} balances",
        "=" * (width + col * len(idx)),
        f"{'':<{width}}" + "".join(f"{h:>{col}}" for h in heads),
        "-" * (width + col * len(idx)),
    ]
    for label, key, kind in STATEMENT_LAYOUT:
        if key is None:
            lines.append("")
            lines.append(label)
            continue
        cells = [_fmt(_value(result, key, i), kind) for i in idx]
        if not any(cells):
            continue
        lines.append(f"{label:<{width}}" + "".join(f"{c:>{col}}" for c in cells))
    balanced = all(abs(v) < 1e-6 for v in result["bs_check"] if v is not None)
    lines += ["", f"Balance sheet balances in every year: {'YES' if balanced else 'NO'}"]
    if result.interest_basis == "average":
        lines.append(f"Average-interest iterations per forecast year: {result.iterations}")
    return "\n".join(lines)


def to_json(result: ModelResult) -> str:
    """Serialise all lines (history plus forecast) as JSON."""
    payload = {
        "scenario": result.scenario,
        "interest_basis": result.interest_basis,
        "years": result.years,
        "n_hist": result.n_hist,
        "lines": result.lines,
    }
    return json.dumps(payload, indent=2)


def to_csv(result: ModelResult) -> str:
    """Serialise all lines as CSV, one row per line item and one column per year."""
    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerow(["line", *[f"FY{y}" for y in result.years]])
    for key, series in result.lines.items():
        writer.writerow([key, *["" if v is None else repr(v) for v in series]])
    return buf.getvalue()


def format_comparison(results: dict[str, ModelResult], inputs: Inputs) -> str:
    """Side-by-side summary of the final forecast year across scenarios."""
    final = inputs.years[-1]
    metrics = [
        (f"FY{final}E revenue ($m)", lambda r: r.last("revenue"), "num"),
        (f"FY{final}E EBITDA ($m)", lambda r: r.last("ebitda"), "num"),
        (f"FY{final}E EBITDA margin", lambda r: r.last("ebitda_margin"), "pct"),
        (f"FY{final}E net income ($m)", lambda r: r.last("net_income"), "num"),
        (f"FY{final}E free cash flow ($m)", lambda r: r.last("fcf"), "num"),
        (f"FY{final}E net debt / EBITDA", lambda r: r.last("net_debt_ebitda"), "mult"),
        ("Peak revolver balance ($m)", lambda r: max(r.forecast("revolver")), "num"),
        ("Cumulative FCF, 5 years ($m)", lambda r: sum(r.forecast("fcf")), "num"),
    ]
    names = list(results)
    lines = [f"{'':<34}" + "".join(f"{n:>12}" for n in names), "-" * (34 + 12 * len(names))]
    for label, fn, kind in metrics:
        lines.append(f"{label:<34}" + "".join(f"{_fmt(fn(results[n]), kind):>12}" for n in names))
    return "\n".join(lines)


def format_sensitivity(grid: SensitivityGrid) -> str:
    """Render both sensitivity grids as text."""
    out: list[str] = []
    for title, data in (
        (f"FY{grid.year}E EBITDA ($m)", grid.ebitda),
        (f"FY{grid.year}E free cash flow ($m)", grid.fcf),
    ):
        out.append(f"{title}: scenario {grid.scenario}")
        head = "margin shift \\ growth shift"
        out.append(f"{head:<28}" + "".join(f"{g * 100:>+10.1f}pp" for g in grid.growth_shifts))
        for m, row in zip(grid.margin_shifts, data, strict=True):
            out.append(f"{m * 100:<+26.1f}pp" + "".join(f"{v:>12,.1f}" for v in row))
        out.append("")
    return "\n".join(out).rstrip()
