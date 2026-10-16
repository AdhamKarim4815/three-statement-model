"""Repository-wide hygiene: typography rules and no placeholder or personal data."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
SKIP_DIRS = {".git", ".venv", "venv", "__pycache__", ".pytest_cache", ".ruff_cache", "build", "dist"}
TEXT_SUFFIXES = {".py", ".md", ".toml", ".yml", ".yaml", ".txt", ".cfg", ".ini", ""}
EM, EN = chr(0x2014), chr(0x2013)
PLACEHOLDER = "TO" + "DO"


def _files(suffixes):
    for path in ROOT.rglob("*"):
        if not path.is_file() or any(p in SKIP_DIRS or p.endswith(".egg-info") for p in path.parts):
            continue
        if path.suffix in suffixes:
            yield path


def test_no_em_or_en_dashes_in_text_files():
    offenders = [
        str(p.relative_to(ROOT)) for p in _files(TEXT_SUFFIXES) if {EM, EN} & set(p.read_text("utf-8", "ignore"))
    ]
    assert offenders == []


def test_no_em_or_en_dashes_inside_committed_workbooks():
    books = list(_files({".xlsx"}))
    assert books, "expected the example workbook to be committed"
    for book in books:
        with zipfile.ZipFile(book) as z:
            for name in z.namelist():
                if name.endswith((".xml", ".rels")):
                    text = z.read(name).decode("utf-8", "ignore")
                    assert EM not in text and EN not in text, f"{book.name}:{name}"


def test_no_placeholders_or_contact_details():
    email = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")
    phone = re.compile(r"\(?\b\d{3}\)?[ -]\d{3}-\d{4}\b")
    for path in _files(TEXT_SUFFIXES):
        if path.name == "test_repo_hygiene.py":
            continue
        text = path.read_text("utf-8", "ignore")
        assert PLACEHOLDER not in text, path
        assert not email.search(text.replace("noreply@", "")), path
        assert not phone.search(text), path


def test_license_and_gitignore_present():
    assert "Copyright (c) 2026 Adham Karim" in (ROOT / "LICENSE").read_text()
    assert "MIT License" in (ROOT / "LICENSE").read_text()
    assert "__pycache__" in (ROOT / ".gitignore").read_text()


@pytest.mark.parametrize("name", ["README.md", "pyproject.toml", ".github/workflows/ci.yml"])
def test_required_files_exist(name):
    assert (ROOT / name).exists()


def test_example_workbook_and_screenshots_committed():
    assert (ROOT / "examples" / "northwind_model.xlsx").stat().st_size > 10_000
    assert list((ROOT / "docs").glob("*.png"))
