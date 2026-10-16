"""Banker-formatted Excel workbook with live formulas.

The workbook is generated with XlsxWriter. All statements, schedules, checks and
ratios are real Excel formulas driven by the Assumptions sheet; the only static
numbers are blue inputs and the clearly labelled Python-generated sensitivity
tables. Cached formula results are written from the Python engine so that the
file displays correctly even in viewers that do not recalculate.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from . import __version__
from .engine import ModelResult, run_model
from .inputs import SCENARIOS, Inputs, canonical_scenario
from .layout import (
    FIRST_COL,
    FIRST_DATA_ROW,
    HEADER_ROW,
    NAME_CHECKS_OK,
    NAME_SCENARIO,
    NAME_TOL,
    SCALAR_NAMES,
    SHEET_NAMES,
    Col,
    Ctx,
    Refs,
    Row,
    assumptions_rows,
    balance_sheet_rows,
    cash_flow_rows,
    checks_rows,
    income_statement_rows,
    quote,
    schedules_rows,
    summary_rows,
)
from .northwind import northwind_inputs
from .sensitivity import SensitivityGrid, sensitivity_grid

# Banker colour code
BLUE = "#0000FF"  # hard-coded inputs
BLACK = "#000000"  # formulas on the same sheet
GREEN = "#008000"  # formulas that pull from other sheets
NAVY = "#1F3A5F"
STEEL = "#4A7FB5"
PALE = "#E8EEF7"
GREY = "#595959"
STATIC_FILL = "#F2F2F2"
STATIC_FONT = "#404040"
FONT = "Arial"

NUMFMT = {
    # conditional sections show "-" for anything that would otherwise display as 0.0 or (0.0)
    "num": r'[>=0.05]#,##0.0_);[<=-0.05]\(#,##0.0\);"-"_)',
    "pct": "0.0%_);(0.0%);0.0%_)",
    "pct2": "0.00%_);(0.00%);0.00%_)",
    "days": "#,##0.0_);(#,##0.0)",
    "mult": '0.00"x"_);(0.00"x")',
    "usd": "$#,##0.00_);($#,##0.00)",
    "int": "#,##0_)",
    "tol": "0.000_)",
    "chk": r'[>=0.0005]#,##0.000_);[<=-0.0005]\(#,##0.000\);"-"_)',
    "text": "@",
}
YEAR_HIST = '"FY"0"A"'
YEAR_FC = '"FY"0"E"'

TIMELINE_SHEETS = ("summary", "assump", "is", "bs", "cf", "sched", "checks")
SHEET_TITLES = {
    "summary": ("Summary dashboard", "Key financials, credit statistics and returns"),
    "assump": ("Assumptions", "All inputs live here. Blue cells are hard-coded inputs"),
    "is": ("Income statement", "Revenue build, margins and earnings"),
    "bs": ("Balance sheet", "Year-end balances"),
    "cf": ("Cash flow statement", "Indirect method"),
    "sched": ("Supporting schedules", "PP&E, working capital, debt, interest, equity and shares"),
    "checks": ("Model integrity checks", "Every row must be zero; status column must read TRUE"),
}
SHEET_PURPOSE = {
    "summary": "Dashboard: key lines, ratios (margins, ROE, ROIC, leverage, coverage, FCF), 3 charts",
    "assump": "Scenario dropdown and every driver (Base, Upside, Downside); history shown for context",
    "is": "Segment revenue build through net income and per-share data",
    "bs": "Assets, liabilities and equity; links to schedules; balance check",
    "cf": "Operating, investing and financing cash flows; ending cash feeds the balance sheet",
    "sched": "PP&E roll-forward, working capital days, term loan, revolver, interest, equity, shares",
    "checks": "Balance, cash tie-out, roll-forwards and debt-limit tests with TRUE/FALSE status",
    "sens": "Python-generated two-way sensitivity of FY2030E EBITDA and FCF (static values)",
}
NOTE_COL = 10  # column K


@dataclass
class BuildInfo:
    """What :func:`build_workbook` produced, for tests and tooling."""

    path: Path
    scenario: str
    result: ModelResult
    rows: dict[str, dict[str, int]]
    recon: list[tuple[str, str, int, str]] = field(default_factory=list)
    names: dict[str, str] = field(default_factory=dict)

    def cell(self, sheet: str, key: str, col: int) -> str:
        """A1 reference of ``key`` in timeline column ``col`` (0-based)."""
        return f"{xl_col_to_name(FIRST_COL + col)}{self.rows[sheet][key]}"


class Styles:
    """Cached XlsxWriter formats."""

    def __init__(self, wb: xlsxwriter.Workbook) -> None:
        self.wb = wb
        self._cache: dict[tuple[tuple[str, Any], ...], Any] = {}

    def get(self, **props: Any) -> Any:
        base: dict[str, Any] = {"font_name": FONT, "font_size": 10, "valign": "vcenter"}
        base.update(props)
        key = tuple(sorted(base.items(), key=lambda kv: kv[0]))
        if key not in self._cache:
            self._cache[key] = self.wb.add_format(base)
        return self._cache[key]


def is_external(formula: str, sheet_key: str) -> bool:
    """Green-font rule: the formula reads from another sheet (directly or via a named input)."""
    if "!" in formula:
        return True
    if sheet_key == "assump":
        return False
    return any(name in formula for name in (*SCALAR_NAMES, NAME_SCENARIO, NAME_TOL))


def _resolve(row: Row, c: Col) -> Any:
    src = row.both if row.both is not None else (row.hist if c.hist else row.fc)
    if src is None:
        return None
    if callable(src):
        return src(c)
    return src[c.i if c.hist else c.t]


def _cached(row: Row, c: Col, ctx: Ctx) -> Any:
    key = row.engine_key
    if key is not None and key in ctx.res.lines:
        value = ctx.res.lines[key][c.i]
        if value is not None:
            return value
    if row.cache is not None:
        return row.cache(c)
    return 0


def _row_props(row: Row) -> dict[str, Any]:
    props: dict[str, Any] = {}
    if row.kind in ("total", "grand"):
        props.update(bold=True, top=1)
    if row.kind == "grand":
        props.update(bottom=6)
    if row.italic:
        props["italic"] = True
    return props


def _write_cell(
    ws: Any, st: Styles, r0: int, c0: int, content: Any, row: Row, c: Col,
    ctx: Ctx, sheet_key: str,
) -> None:  # fmt: skip
    if content is None:
        return
    props = _row_props(row)
    numfmt = NUMFMT[row.fmt]
    if isinstance(content, str) and content.startswith("="):
        color = GREEN if is_external(content, sheet_key) else BLACK
        fmt = st.get(num_format=numfmt, font_color=color, align="right", **props)
        ws.write_formula(r0, c0, content, fmt, _cached(row, c, ctx))
    elif isinstance(content, str):
        ws.write_string(r0, c0, content, st.get(font_color=BLUE, align="right", **props))
    else:
        ws.write_number(r0, c0, content, st.get(num_format=numfmt, font_color=BLUE, align="right", **props))


def _label_fmt(st: Styles, row: Row) -> Any:
    props = _row_props(row)
    props["indent"] = row.indent
    return st.get(**props)


def _write_headers(
    ws: Any, st: Styles, ctx: Ctx, sheet_key: str, last_col: int, R: Refs
) -> None:  # fmt: skip
    title, subtitle = SHEET_TITLES[sheet_key]
    n_hist, n_fc = ctx.inp.n_hist, ctx.inp.n_forecast
    ws.write_string(0, 0, f"{ctx.inp.company} (fictional) | {title}", st.get(bold=True, font_size=14, font_color=NAVY))
    ws.write_string(
        1,
        0,
        f"{subtitle}. $ in millions except per share data; fiscal year ends December 31.",
        st.get(italic=True, font_color=GREY, font_size=9),
    )
    ws.write_formula(
        2, 0, f'="Scenario: "&{NAME_SCENARIO}', st.get(bold=True, font_color=GREEN), f"Scenario: {ctx.scenario}"
    )
    band = st.get(bold=True, font_color="#FFFFFF", bg_color=STEEL, align="center")
    ws.merge_range(3, FIRST_COL, 3, FIRST_COL + n_hist - 1, "Actual", band)
    ws.merge_range(3, FIRST_COL + n_hist, 3, FIRST_COL + n_hist + n_fc - 1, "Forecast", band)
    hdr = {"bold": True, "font_color": "#FFFFFF", "bg_color": NAVY, "align": "right"}
    ws.write_string(4, 0, "($ in millions)", st.get(**{**hdr, "align": "left"}))
    ws.write_string(4, 1, "Unit", st.get(**{**hdr, "align": "center"}))
    for i in range(n_hist + n_fc):
        c = Col(i, n_hist)
        year = ctx.res.years[i]
        numfmt = YEAR_HIST if c.hist else YEAR_FC
        fmt = st.get(num_format=numfmt, **hdr)
        suffix = "A" if c.hist else "E"
        if sheet_key == "assump":
            if c.first:
                ws.write_number(4, FIRST_COL, year, st.get(num_format=numfmt, **{**hdr, "font_color": "#9DC3FF"}))
            else:
                ws.write_formula(4, FIRST_COL + i, f"={c.prev}{HEADER_ROW}+1", fmt, year)
        elif sheet_key == "summary":
            ws.write_formula(
                4,
                FIRST_COL + i,
                f'="FY"&Assumptions!{c.letter}${HEADER_ROW}&"{suffix}"',
                st.get(**hdr),
                f"FY{year}{suffix}",
            )
        else:
            ws.write_formula(4, FIRST_COL + i, f"=Assumptions!{c.letter}${HEADER_ROW}", fmt, year)
    for col in range(FIRST_COL + n_hist + n_fc, last_col + 1):
        ws.write_blank(4, col, None, st.get(**hdr))
    if sheet_key == "assump":
        ws.write_string(4, NOTE_COL, "Notes", st.get(**{**hdr, "align": "left", "indent": 1}))


def _write_rows(
    ws: Any, st: Styles, ctx: Ctx, R: Refs, sheet_key: str, rows: list[Row], last_col: int
) -> None:  # fmt: skip
    n_hist, n_fc = ctx.inp.n_hist, ctx.inp.n_forecast
    for row in rows:
        r0 = row.xl - 1
        if row.kind == "blank":
            continue
        if row.kind == "section":
            fmt = st.get(bold=True, font_color=NAVY, bg_color=PALE, bottom=1, bottom_color=NAVY)
            ws.write_string(r0, 0, row.label, fmt)
            for col in range(1, last_col + 1):
                ws.write_blank(r0, col, None, fmt)
            continue
        if row.kind == "text":
            ws.write_string(r0, 0, row.label, st.get(italic=True, font_color=GREY, font_size=9))
            continue
        ws.write_string(r0, 0, row.label, _label_fmt(st, row))
        unit_props = {**_row_props(row), "font_color": GREY, "align": "center", "font_size": 9}
        unit_props.pop("italic", None)
        ws.write_string(r0, 1, row.unit, st.get(**unit_props))
        if row.kind == "scalar":
            _write_scalar(ws, st, ctx, R, sheet_key, row)
        else:
            for i in range(n_hist + n_fc):
                c = Col(i, n_hist)
                _write_cell(ws, st, r0, FIRST_COL + i, _resolve(row, c), row, c, ctx, sheet_key)
        if row.note and sheet_key == "assump":
            ws.write_string(r0, NOTE_COL, row.note, st.get(italic=True, font_color=GREY, font_size=9, indent=1))


def _write_scalar(ws: Any, st: Styles, ctx: Ctx, R: Refs, sheet_key: str, row: Row) -> None:
    r0 = row.xl - 1
    value = row.scalar(R) if callable(row.scalar) else row.scalar
    numfmt = NUMFMT[row.fmt]
    if isinstance(value, str) and value.startswith("="):
        fmt = st.get(num_format=numfmt, font_color=BLACK, align="right")
        cached = row.cache(Col(0, ctx.inp.n_hist)) if row.cache else 0
        ws.write_formula(r0, FIRST_COL, value, fmt, cached)
    elif isinstance(value, str):
        extra = {}
        if row.key == "scenario":
            extra = {"bg_color": "#FFF2CC", "border": 1, "bold": True}
        ws.write_string(r0, FIRST_COL, value, st.get(font_color=BLUE, align="right", **extra))
    else:
        ws.write_number(r0, FIRST_COL, value, st.get(num_format=numfmt, font_color=BLUE, align="right"))


def _page_setup(ws: Any, last_row: int, last_col: int, landscape: bool = True, tall: int = 0) -> None:
    ws.hide_gridlines(2)
    if landscape:
        ws.set_landscape()
    else:
        ws.set_portrait()
    ws.set_paper(1)
    ws.fit_to_pages(1, tall)
    ws.print_area(0, 0, last_row, last_col)
    ws.center_horizontally()
    ws.set_margins(left=0.5, right=0.5, top=0.6, bottom=0.6)
    ws.set_footer("&L&8Northwind Components Inc. (fictional company)&C&8&A&R&8Page &P of &N")


def _write_checks_extras(
    ws: Any, st: Styles, ctx: Ctx, R: Refs, rows: list[Row], last_col: int
) -> int:  # fmt: skip
    hdr = {"bold": True, "font_color": "#FFFFFF", "bg_color": NAVY, "align": "center"}
    ws.write_string(HEADER_ROW - 1, last_col, "Pass?", st.get(**hdr))
    check_rows = [r for r in rows if r.key and r.key.startswith("chk_")]
    for row in check_rows:
        rng = f"C{row.xl}:J{row.xl}"
        ws.write_formula(
            row.xl - 1,
            last_col,
            f"=SUMPRODUCT(ABS({rng}))<{NAME_TOL}",
            st.get(font_color=GREEN, align="center", bold=True),
            True,
        )
    first, last = check_rows[0].xl, check_rows[-1].xl
    r_all = last + 2
    ws.write_string(r_all - 1, 0, "ALL CHECKS PASS", st.get(bold=True, top=1, bottom=6))
    ws.write_formula(
        r_all - 1,
        FIRST_COL,
        f"=AND({xl_col_to_name(last_col)}{first}:{xl_col_to_name(last_col)}{last})",
        st.get(bold=True, font_color=BLACK, align="center", top=1, bottom=6),
        True,
    )
    ws.write_string(
        r_all,
        0,
        "Tolerance is set on the Assumptions sheet (named range Tolerance).",
        st.get(italic=True, font_color=GREY, font_size=9),
    )
    return r_all


# --------------------------------------------------------------------------- charts


def _add_charts(wb: Any, ws: Any, R: Refs, ctx: Ctx, anchor_row: int) -> int:
    palette = ["#1F3A5F", "#4A7FB5", "#9DC3E6"]
    accent, green = "#E07B39", "#2E8B57"
    cats = f"=Summary!$C${HEADER_ROW}:$J${HEADER_ROW}"
    font = {"name": FONT, "size": 9, "color": "#404040"}
    grid = {"visible": True, "line": {"color": "#D9D9D9", "width": 0.5}}

    def style(chart: Any, title: str, w: int, h: int) -> None:
        chart.set_title({"name": title, "name_font": {"name": FONT, "size": 11, "bold": True, "color": NAVY}})
        chart.set_legend({"position": "bottom", "font": font})
        chart.set_chartarea({"border": {"color": "#BFBFBF", "width": 0.75}})
        chart.set_size({"width": w, "height": h})
        chart.set_x_axis({"num_font": font, "line": {"color": "#808080"}})
        chart.set_y_axis({"num_font": font, "num_format": "#,##0", "major_gridlines": grid, "line": {"none": True}})

    c1 = wb.add_chart({"type": "column", "subtype": "stacked"})
    for k, _ in enumerate(ctx.inp.segments, 1):
        r = R.row("is", f"rev_s{k}")
        c1.add_series(
            {
                "name": f"='Income Statement'!$A${r}",
                "categories": cats,
                "values": f"='Income Statement'!$C${r}:$J${r}",
                "fill": {"color": palette[(k - 1) % 3]},
                "gap": 60,
            }
        )
    style(c1, "Revenue by segment ($m)", 518, 300)

    c2 = wb.add_chart({"type": "column"})
    r = R.row("summary", "ebitda")
    c2.add_series(
        {
            "name": "EBITDA ($m)",
            "categories": cats,
            "values": f"=Summary!$C${r}:$J${r}",
            "fill": {"color": NAVY},
            "gap": 60,
        }
    )
    line = wb.add_chart({"type": "line"})
    r = R.row("summary", "ebitda_margin")
    line.add_series(
        {
            "name": "EBITDA margin (right axis)",
            "categories": cats,
            "values": f"=Summary!$C${r}:$J${r}",
            "y2_axis": True,
            "line": {"color": accent, "width": 2.25},
            "marker": {"type": "circle", "size": 6, "fill": {"color": accent}, "border": {"color": accent}},
        }
    )
    line.set_y2_axis({"num_font": font, "num_format": "0%", "line": {"none": True}, "min": 0.10, "max": 0.25})
    c2.combine(line)
    style(c2, "EBITDA ($m) and EBITDA margin", 480, 300)

    c3 = wb.add_chart({"type": "column", "subtype": "stacked"})
    r = R.row("sched", "tl_close")
    c3.add_series(
        {
            "name": "Term loan",
            "categories": cats,
            "values": f"=Schedules!$C${r}:$J${r}",
            "fill": {"color": NAVY},
            "gap": 80,
        }
    )
    r = R.row("sched", "rev_close")
    c3.add_series(
        {"name": "Revolver", "categories": cats, "values": f"=Schedules!$C${r}:$J${r}", "fill": {"color": accent}}
    )
    line3 = wb.add_chart({"type": "line"})
    r = R.row("bs", "cash")
    line3.add_series(
        {
            "name": "Cash",
            "categories": cats,
            "values": f"='Balance Sheet'!$C${r}:$J${r}",
            "line": {"color": green, "width": 2.25},
            "marker": {"type": "diamond", "size": 7, "fill": {"color": green}, "border": {"color": green}},
        }
    )
    c3.combine(line3)
    style(c3, "Debt and cash ($m): the revolver draws to defend minimum cash", 1012, 290)

    ws.insert_chart(anchor_row, 0, c1, {"x_offset": 2, "y_offset": 4})
    ws.insert_chart(anchor_row, 4, c2, {"x_offset": 6, "y_offset": 4})
    ws.insert_chart(anchor_row + 15, 0, c3, {"x_offset": 2, "y_offset": 10})
    return anchor_row + 15 + 15


# --------------------------------------------------------------------------- other sheets


def _write_sensitivity(ws: Any, st: Styles, ctx: Ctx, grid: SensitivityGrid) -> int:
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 6, 13)
    ws.write_string(
        0, 0, f"{ctx.inp.company} (fictional) | Sensitivity analysis", st.get(bold=True, font_size=14, font_color=NAVY)
    )
    ws.write_string(
        1,
        0,
        f"FY{grid.year}E EBITDA and free cash flow. $ in millions. Scenario at build time: {grid.scenario}.",
        st.get(italic=True, font_color=GREY, font_size=9),
    )
    banner = st.get(
        bold=True,
        font_color="#7F6000",
        bg_color="#FFF2CC",
        text_wrap=True,
        border=1,
        border_color="#BF9000",
        valign="vcenter",
    )
    ws.merge_range(
        3,
        0,
        3,
        6,
        "PYTHON-GENERATED TABLE: static values, not Excel formulas. Each "
        "cell is a full re-run of the three-statement model (three_statement.sensitivity) "
        "with every segment's growth and the gross margin shifted in every forecast "
        f"year. It does NOT respond to the scenario dropdown; it reflects the {grid.scenario} "
        "scenario selected when the file was built. Rebuild with "
        "'three-statement build --scenario <name>' to refresh.",
        banner,
    )
    ws.set_row(3, 62)
    head = st.get(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True)
    shift = '+0.0" pp";-0.0" pp";0.0" pp"'
    value = st.get(
        num_format=NUMFMT["num"],
        font_color=STATIC_FONT,
        bg_color=STATIC_FILL,
        align="right",
        border=1,
        border_color="#D9D9D9",
    )
    centre = st.get(
        num_format=NUMFMT["num"], font_color=STATIC_FONT, bg_color="#FFE699", align="right", bold=True, border=2
    )
    shift_fmt = st.get(num_format=shift, bold=True, font_color=STATIC_FONT, bg_color=PALE, align="center")
    row = 5
    for title, data in (
        (f"FY{grid.year}E EBITDA ($m)", grid.ebitda),
        (f"FY{grid.year}E free cash flow ($m)", grid.fcf),
    ):
        ws.write_string(row, 0, title, st.get(bold=True, font_size=11, font_color=NAVY))
        ws.merge_range(row + 1, 2, row + 1, 6, "Revenue growth shift (pp per year, every segment)", head)
        ws.write_string(row + 2, 1, "Margin shift", head)
        for j, g in enumerate(grid.growth_shifts):
            ws.write_number(row + 2, 2 + j, g * 100, shift_fmt)
        for i, m in enumerate(grid.margin_shifts):
            if i == len(grid.margin_shifts) // 2:
                ws.write_string(
                    row + 3 + i,
                    0,
                    "Gross margin shift (pp, every forecast year)",
                    st.get(italic=True, font_color=GREY, align="right"),
                )
            ws.write_number(row + 3 + i, 1, m * 100, shift_fmt)
            for j, g in enumerate(grid.growth_shifts):
                is_centre = abs(g) < 1e-12 and abs(m) < 1e-12
                ws.write_number(row + 3 + i, 2 + j, data[i][j], centre if is_centre else value)
        row += 3 + len(grid.margin_shifts) + 2
    ws.write_string(
        row - 1,
        0,
        "Highlighted cell: unshifted model for the scenario above (matches "
        "the Summary sheet when that scenario is selected).",
        st.get(italic=True, font_color=GREY, font_size=9),
    )
    ws.freeze_panes(3, 0)
    _page_setup(ws, row, 6)
    return row


def _write_cover(wb: Any, ws: Any, st: Styles, ctx: Ctx, R: Refs, check_ok: bool) -> None:
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 96)
    ws.hide_gridlines(2)
    ws.set_tab_color(NAVY)
    scen_row = R.row("assump", "scenario")
    ws.write_string(1, 1, ctx.inp.company, st.get(bold=True, font_size=22, font_color=NAVY))
    ws.write_string(
        2,
        1,
        "Integrated three-statement financial model: income statement, balance sheet, cash flow",
        st.get(font_size=12, font_color=GREY),
    )
    ws.write_string(
        3,
        1,
        "FICTIONAL COMPANY. All figures are illustrative; this company and its numbers do not exist.",
        st.get(bold=True, font_color="#9C0006"),
    )
    ws.write_string(
        4,
        1,
        f"Built by Adham Karim | github.com/adhamkarim4815/three-statement-model | v{__version__}",
        st.get(italic=True, font_color=GREY, font_size=9),
    )

    def section(r: int, title: str) -> None:
        fmt = st.get(bold=True, font_color=NAVY, bg_color=PALE, bottom=1, bottom_color=NAVY)
        ws.write_string(r, 1, title, fmt)
        ws.write_blank(r, 2, None, fmt)

    def line(r: int, a: str, b: str, bold: bool = False) -> None:
        ws.write_string(r, 1, a, st.get(bold=True, valign="top"))
        ws.write_string(r, 2, b, st.get(text_wrap=True, valign="top", bold=bold))

    r = 6
    section(r, "Model status")
    ws.write_string(r + 1, 1, "Active scenario", st.get(bold=True))
    ws.write_formula(r + 1, 2, f"={NAME_SCENARIO}", st.get(font_color=GREEN, bold=True, align="left"), ctx.scenario)
    ws.write_string(r + 2, 1, "Model checks", st.get(bold=True))
    ws.write_formula(
        r + 2,
        2,
        f'=IF({NAME_CHECKS_OK},"ALL CHECKS PASS","CHECK ERRORS: see Checks sheet")',
        st.get(font_color=GREEN, bold=True, align="left"),
        "ALL CHECKS PASS" if check_ok else "CHECK ERRORS: see Checks sheet",
    )
    line(r + 3, "Units", "$ in millions except per share data. Fiscal year ends December 31.")
    line(
        r + 4,
        "Timeline",
        f"{ctx.inp.n_hist} actual years (FY{ctx.res.years[0]}A to "
        f"FY{ctx.res.years[ctx.inp.n_hist - 1]}A) and {ctx.inp.n_forecast} forecast years "
        f"(FY{ctx.res.years[ctx.inp.n_hist]}E to FY{ctx.res.years[-1]}E). "
        "Same column = same year on every sheet.",
    )

    r = 12
    section(r, "How to use")
    steps = [
        (
            "1. Pick a scenario",
            f"Use the dropdown in Assumptions!C{scen_row} (named range Scenario): "
            "Base, Upside or Downside. Everything recalculates, including the charts.",
        ),
        (
            "2. Change drivers",
            "Edit blue cells on the Assumptions sheet. Each scenario driver has a live "
            "row (CHOOSE on the scenario index) above Base, Upside and Downside input rows.",
        ),
        (
            "3. Read the output",
            "Summary shows ratios and charts. Statements and Schedules show the full "
            "build. Checks must read ALL CHECKS PASS.",
        ),
        (
            "4. Sensitivity",
            "The Sensitivity sheet is a static, Python-generated table. It does not update with the dropdown.",
        ),
    ]
    for k, (a, b) in enumerate(steps):
        line(r + 1 + k, a, b)

    r = 18
    section(r, "Sheet guide (click to navigate)")
    for k, key in enumerate(("summary", "assump", "is", "bs", "cf", "sched", "checks", "sens")):
        name = SHEET_NAMES[key]
        ws.write_url(r + 1 + k, 1, f"internal:'{name}'!A1", st.get(font_color="#0563C1", underline=1), string=name)
        ws.write_string(r + 1 + k, 2, SHEET_PURPOSE[key], st.get())

    r = 28
    section(r, "Colour code")
    ws.write_number(r + 1, 1, 1234.5, st.get(num_format=NUMFMT["num"], font_color=BLUE, align="left"))
    ws.write_string(r + 1, 2, "Blue: hard-coded input or historical data point. Change these.", st.get())
    ws.write_number(r + 2, 1, 1234.5, st.get(num_format=NUMFMT["num"], font_color=BLACK, align="left"))
    ws.write_string(r + 2, 2, "Black: formula that calculates from cells on the same sheet.", st.get())
    ws.write_number(r + 3, 1, 1234.5, st.get(num_format=NUMFMT["num"], font_color=GREEN, align="left"))
    ws.write_string(r + 3, 2, "Green: formula that pulls from another sheet or from a named input.", st.get())
    ws.write_number(
        r + 4, 1, 1234.5, st.get(num_format=NUMFMT["num"], font_color=STATIC_FONT, bg_color=STATIC_FILL, align="left")
    )
    ws.write_string(r + 4, 2, "Grey fill: static values generated by Python (Sensitivity sheet only).", st.get())
    ws.write_string(r + 5, 1, "Yellow cell", st.get(bg_color="#FFF2CC", border=1, bold=True))
    ws.write_string(r + 5, 2, "The scenario selector.", st.get())

    r = 35
    section(r, "Conventions")
    conv = [
        ("Signs", "Income statement expenses are positive and subtracted. Cash flow outflows are negative."),
        (
            "Interest",
            "Opening-balance basis: no circular references, so Excel never needs iterative "
            "calculation. The Python engine can also run an average-balance basis (iterative).",
        ),
        (
            "Revolver",
            "Draws to keep cash at or above the minimum cash input, repays from any excess, "
            "never goes negative and is capped at the commitment.",
        ),
        ("Named ranges", "Scenario, Scenario_Idx, Days_In_Year, Min_Cash, Revolver_Limit, Tolerance, Model_Checks_OK."),
        (
            "Validation",
            "The test suite recalculates this workbook in LibreOffice and reconciles every "
            "line to the Python engine in all three scenarios.",
        ),
    ]
    for k, (a, b) in enumerate(conv):
        line(r + 1 + k, a, b)
    ws.set_row(37, 30)
    ws.set_row(38, 30)
    ws.set_row(40, 30)
    r = 42
    section(r, "Limitations")
    ws.write_string(r + 1, 1, "Simplified model", st.get(bold=True, valign="top"))
    ws.write_string(
        r + 1,
        2,
        "Fictional company; simplified cash taxes (no deferred taxes or loss "
        "carry-forwards); no leases; goodwill and other long-term liabilities held flat; no "
        "FX; one term loan and one revolver; historical cash flows not presented.",
        st.get(text_wrap=True, valign="top"),
    )
    ws.set_row(r + 1, 45)
    ws.set_landscape()
    ws.set_paper(1)
    ws.fit_to_pages(1, 1)
    ws.print_area(0, 0, r + 2, 2)
    ws.center_horizontally()


# --------------------------------------------------------------------------- entry point


def build_workbook(
    path: str | Path, inputs: Inputs | None = None, scenario: str = "Base", sensitivity: SensitivityGrid | None = None
) -> BuildInfo:
    """Write the workbook to ``path`` and return a :class:`BuildInfo`.

    ``scenario`` sets the dropdown value and the cached results. The sensitivity
    sheet is computed from the engine for the same scenario unless supplied.
    """
    inputs = inputs or northwind_inputs()
    scenario = canonical_scenario(scenario)
    result = run_model(inputs, scenario, "opening")
    ctx = Ctx(inp=inputs, res=result, scenario=scenario)
    R = Refs()
    specs: dict[str, list[Row]] = {
        "assump": assumptions_rows(R, ctx),
        "is": income_statement_rows(R, ctx),
        "bs": balance_sheet_rows(R, ctx),
        "cf": cash_flow_rows(R, ctx),
        "sched": schedules_rows(R, ctx),
        "summary": summary_rows(R, ctx),
        "checks": checks_rows(R, ctx),
    }
    for sheet_key, rows in specs.items():
        R.rows[sheet_key] = {}
        for k, row in enumerate(rows):
            row.xl = FIRST_DATA_ROW + k
            if row.key:
                R.rows[sheet_key][row.key] = row.xl

    path = Path(path)
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties(
        {
            "title": f"{inputs.company} three-statement model (fictional)",
            "subject": "Integrated income statement, balance sheet and cash flow forecast",
            "author": "Adham Karim",
            "comments": "Fictional company for demonstration. Generated by three-statement-model.",
        }
    )
    st = Styles(wb)
    sheets = {key: wb.add_worksheet(name) for key, name in SHEET_NAMES.items()}
    info = BuildInfo(path=path, scenario=scenario, result=result, rows=R.rows)
    check_ok = all(abs(v) < inputs.tolerance for v in result["bs_check"] if v is not None)

    last_year_col = FIRST_COL + inputs.n_hist + inputs.n_forecast - 1
    chart_end = 0
    for sheet_key in TIMELINE_SHEETS:
        ws = sheets[sheet_key]
        R.current = sheet_key
        rows = specs[sheet_key]
        last_col = (
            NOTE_COL if sheet_key == "assump" else (last_year_col + 1 if sheet_key == "checks" else last_year_col)
        )
        ws.set_column(0, 0, {"assump": 50, "checks": 64}.get(sheet_key, 44))
        ws.set_column(1, 1, 7)
        ws.set_column(FIRST_COL, last_year_col, 11)
        if sheet_key == "assump":
            ws.set_column(NOTE_COL, NOTE_COL, 62)
        if sheet_key == "checks":
            ws.set_column(last_col, last_col, 9)
        _write_headers(ws, st, ctx, sheet_key, last_col, R)
        _write_rows(ws, st, ctx, R, sheet_key, rows, last_col)
        ws.freeze_panes(HEADER_ROW, FIRST_COL)
        ws.set_zoom(90)
        ws.repeat_rows(3, 4)
        end_row = rows[-1].xl
        if sheet_key == "checks":
            end_row = _write_checks_extras(ws, st, ctx, R, rows, last_col) + 1
        if sheet_key == "summary":
            chart_end = _add_charts(wb, ws, R, ctx, end_row + 1)
            _page_setup(ws, chart_end, last_col, landscape=False, tall=1)
        else:
            _page_setup(ws, end_row, last_col)
        ws.set_tab_color({"assump": "#0000FF", "checks": "#C00000"}.get(sheet_key, STEEL))
        for row in rows:
            if row.kind == "scalar" and row.name:
                wb.define_name(row.name, f"=Assumptions!$C${row.xl}")
                info.names[row.name] = f"Assumptions!$C${row.xl}"
            if row.kind not in ("scalar", "section", "blank", "text") and row.engine_key in result.lines:
                info.recon.append((sheet_key, SHEET_NAMES[sheet_key], row.xl, row.engine_key))

    r_all = R.rows["checks"]["chk_negdebt"] + 2
    wb.define_name(NAME_CHECKS_OK, f"=Checks!$C${r_all}")
    info.names[NAME_CHECKS_OK] = f"Checks!$C${r_all}"

    # Scenario dropdown
    scen_row = R.row("assump", "scenario")
    sc1, sc3 = R.row("assump", "sc1"), R.row("assump", "sc3")
    sheets["assump"].data_validation(
        scen_row - 1,
        FIRST_COL,
        scen_row - 1,
        FIRST_COL,
        {
            "validate": "list",
            "source": f"=$C${sc1}:$C${sc3}",
            "input_title": "Scenario",
            "input_message": "Choose Base, Upside or Downside.",
            "error_title": "Invalid scenario",
            "error_message": "Pick a scenario from the list.",
        },
    )

    grid = sensitivity or sensitivity_grid(inputs, scenario)
    _write_sensitivity(sheets["sens"], st, ctx, grid)
    sheets["sens"].set_tab_color("#7F7F7F")
    _write_cover(wb, sheets["cover"], st, ctx, R, check_ok)
    sheets["cover"].activate()
    del chart_end
    wb.close()
    return info


def scenario_names() -> tuple[str, ...]:
    return SCENARIOS


__all__ = ["BuildInfo", "build_workbook", "is_external", "quote", "scenario_names"]
