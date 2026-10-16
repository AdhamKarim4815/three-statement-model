from __future__ import annotations

import csv
import io
import json

import pytest

from three_statement import __version__
from three_statement.cli import build_parser, main
from three_statement.report import format_comparison, format_statements, to_csv, to_json


def test_statement_table_has_all_three_statements(results):
    text = format_statements(results["Base"])
    for heading in ("INCOME STATEMENT", "BALANCE SHEET", "CASH FLOW STATEMENT", "KEY RATIOS"):
        assert heading in text
    assert "FY2023A" in text and "FY2030E" in text
    assert "Balance sheet balances in every year: YES" in text


def test_forecast_only_hides_history(results):
    text = format_statements(results["Base"], forecast_only=True)
    assert "FY2023A" not in text and "FY2026E" in text


def test_average_basis_report_mentions_iterations(avg_results):
    assert "iterations" in format_statements(avg_results["Base"])


def test_json_round_trip(results):
    payload = json.loads(to_json(results["Downside"]))
    assert payload["scenario"] == "Downside"
    assert payload["years"][0] == 2023
    assert len(payload["lines"]["revenue"]) == 8


def test_csv_round_trip(results):
    rows = list(csv.reader(io.StringIO(to_csv(results["Base"]))))
    assert rows[0][0] == "line" and rows[0][-1] == "FY2030"
    revenue = next(r for r in rows if r[0] == "revenue")
    assert float(revenue[-1]) == pytest.approx(results["Base"].last("revenue"))


def test_comparison_table(inputs, results):
    text = format_comparison(results, inputs)
    assert all(name in text for name in ("Base", "Upside", "Downside"))
    assert "Peak revolver" in text


def test_cli_run_prints_statements(capsys):
    assert main(["run", "--scenario", "downside"]) == 0
    out = capsys.readouterr().out
    assert "scenario Downside" in out and "Revolver" in out


def test_cli_run_average_basis(capsys):
    assert main(["run", "--scenario", "base", "--interest", "average", "--forecast-only"]) == 0
    assert "average balances" in capsys.readouterr().out


def test_cli_run_json_and_csv(capsys):
    main(["run", "--format", "json"])
    assert json.loads(capsys.readouterr().out)["scenario"] == "Base"
    main(["run", "--format", "csv", "--scenario", "upside"])
    assert capsys.readouterr().out.startswith("line,FY2023")


def test_cli_compare_and_sensitivity(capsys):
    assert main(["compare"]) == 0
    assert "FY2030E EBITDA" in capsys.readouterr().out
    assert main(["sensitivity", "--scenario", "base"]) == 0
    assert "margin shift" in capsys.readouterr().out


def test_cli_build_writes_workbook(tmp_path, capsys):
    out = tmp_path / "model.xlsx"
    assert main(["build", "--scenario", "upside", "--out", str(out)]) == 0
    assert out.exists() and out.stat().st_size > 10_000
    assert "dropdown set to Upside" in capsys.readouterr().out


def test_cli_rejects_unknown_scenario(capsys):
    with pytest.raises(SystemExit) as exc:
        main(["run", "--scenario", "moonshot"])
    assert exc.value.code == 2
    assert "Unknown scenario" in capsys.readouterr().err


def test_cli_requires_a_command():
    with pytest.raises(SystemExit):
        build_parser().parse_args([])


def test_cli_version(capsys):
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0
    assert __version__ in capsys.readouterr().out
