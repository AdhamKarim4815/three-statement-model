"""Regenerate examples/northwind_model.xlsx and the README screenshots.

Usage (from the repository root, with LibreOffice and poppler-utils installed)::

    python scripts/render_docs.py

The workbook is converted to PDF with ``soffice --headless``; the relevant pages
are rasterised with ``pdftoppm`` and trimmed with Pillow when it is available.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from three_statement.excel import build_workbook
from three_statement.validate import find_soffice

ROOT = Path(__file__).resolve().parent.parent
PAGES = {
    "summary_sheet.png": "Summary dashboard",
    "assumptions_sheet.png": "| Assumptions",
    "income_statement.png": "| Income statement",
}


def _page_text(pdf: Path, page: int) -> str:
    out = subprocess.run(["pdftotext", "-f", str(page), "-l", str(page), str(pdf), "-"], capture_output=True, text=True)
    return out.stdout


def _trim(png: Path, margin: int = 24) -> None:
    try:
        from PIL import Image, ImageChops
    except ImportError:  # pragma: no cover
        return
    img = Image.open(png).convert("RGB")
    box = ImageChops.difference(img, Image.new("RGB", img.size, (255, 255, 255))).getbbox()
    if box:
        left, top, right, bottom = box
        img = img.crop(
            (
                max(0, left - margin),
                max(0, top - margin),
                min(img.width, right + margin),
                min(img.height, bottom + margin),
            )
        )
        img.save(png, optimize=True)


def main() -> int:
    soffice = find_soffice()
    if soffice is None or shutil.which("pdftoppm") is None:
        print("Needs soffice (LibreOffice) and pdftoppm (poppler-utils).", file=sys.stderr)
        return 1
    example = ROOT / "examples" / "northwind_model.xlsx"
    example.parent.mkdir(exist_ok=True)
    build_workbook(example, scenario="Base")
    print(f"wrote {example.relative_to(ROOT)}")
    docs = ROOT / "docs"
    docs.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        subprocess.run(
            [
                soffice,
                f"-env:UserInstallation={(tmp_dir / 'profile').as_uri()}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_dir),
                str(example),
            ],
            check=True,
            capture_output=True,
            timeout=240,
        )
        pdf = tmp_dir / "northwind_model.pdf"
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
        n_pages = int(next(line.split()[1] for line in info.splitlines() if line.startswith("Pages:")))
        for name, needle in PAGES.items():
            page = next(p for p in range(1, n_pages + 1) if needle in _page_text(pdf, p))
            stem = tmp_dir / name.removesuffix(".png")
            subprocess.run(
                ["pdftoppm", "-r", "130", "-f", str(page), "-l", str(page), "-png", "-singlefile", str(pdf), str(stem)],
                check=True,
            )
            target = docs / name
            shutil.move(str(stem) + ".png", target)
            _trim(target)
            print(f"wrote {target.relative_to(ROOT)} (PDF page {page})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
