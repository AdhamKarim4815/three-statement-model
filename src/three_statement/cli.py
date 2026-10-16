"""Command line interface: ``three-statement build | run | compare | sensitivity``."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Sequence
from pathlib import Path

from . import __version__
from .engine import INTEREST_BASES, run_model
from .excel import build_workbook
from .inputs import SCENARIOS, Inputs, canonical_scenario
from .northwind import northwind_inputs
from .report import format_comparison, format_sensitivity, format_statements, to_csv, to_json
from .sensitivity import sensitivity_grid


def _scenario_arg(value: str) -> str:
    try:
        return canonical_scenario(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="three-statement",
        description="Integrated three-statement model for the fictional Northwind Components Inc.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    def add_scenario(p: argparse.ArgumentParser) -> None:
        p.add_argument(
            "--scenario",
            type=_scenario_arg,
            default="Base",
            metavar="NAME",
            help=f"one of {', '.join(s.lower() for s in SCENARIOS)} (default: base)",
        )

    p_build = sub.add_parser("build", help="write the Excel workbook with live formulas")
    add_scenario(p_build)
    p_build.add_argument(
        "--out",
        type=Path,
        default=Path("northwind_model.xlsx"),
        help="output .xlsx path (default: northwind_model.xlsx)",
    )

    p_run = sub.add_parser("run", help="run the Python engine and print the statements")
    add_scenario(p_run)
    p_run.add_argument(
        "--interest", choices=INTEREST_BASES, default="opening", help="interest basis (default: opening)"
    )
    p_run.add_argument("--format", choices=("table", "json", "csv"), default="table")
    p_run.add_argument("--forecast-only", action="store_true", help="hide historical columns")

    sub.add_parser("compare", help="compare scenarios side by side (final forecast year)")

    p_val = sub.add_parser("validate", help="recalculate the workbook in LibreOffice and reconcile to Python")
    p_val.add_argument("--scenario", default="all", help="base, upside, downside or all (default: all)")

    p_sens = sub.add_parser("sensitivity", help="growth versus gross margin sensitivity grids")
    add_scenario(p_sens)
    return parser


def _validate(inputs: Inputs, scenario: str) -> int:
    from .validate import find_soffice, validate_all

    if find_soffice() is None:
        print("LibreOffice (soffice) not found: cannot recalculate the workbook.", file=sys.stderr)
        return 3
    names = SCENARIOS if scenario.lower() == "all" else (canonical_scenario(scenario),)
    exit_code = 0
    for report in validate_all(inputs, names):
        status = "OK" if report.ok else "FAILED"
        print(
            f"{report.scenario:<9} {status}: {report.compared} cells reconciled, "
            f"max abs diff {report.max_abs_diff:.2e}, {report.check_cells} check cells, "
            f"checks pass: {report.checks_pass}"
        )
        for line in report.mismatches[:10]:
            print(f"  {line}")
        exit_code |= 0 if report.ok else 1
    return exit_code


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    inputs = northwind_inputs()
    if args.command == "build":
        info = build_workbook(args.out, inputs, args.scenario)
        print(
            f"Wrote {args.out} (scenario dropdown set to {args.scenario}; "
            f"{len(info.recon)} formula rows linked to the Assumptions sheet)"
        )
        return 0
    if args.command == "run":
        result = run_model(inputs, args.scenario, args.interest)
        if args.format == "json":
            print(to_json(result))
        elif args.format == "csv":
            sys.stdout.write(to_csv(result))
        else:
            print(format_statements(result, forecast_only=args.forecast_only))
        return 0
    if args.command == "compare":
        results = {s: run_model(inputs, s) for s in SCENARIOS}
        print(format_comparison(results, inputs))
        return 0
    if args.command == "validate":
        return _validate(inputs, args.scenario)
    if args.command == "sensitivity":
        print(format_sensitivity(sensitivity_grid(inputs, args.scenario)))
        return 0
    return 2  # pragma: no cover


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
