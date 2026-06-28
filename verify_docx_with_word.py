#!/usr/bin/env python3
"""Render DOCX via Microsoft Word COM, then rasterize the PDF to PNG pages."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageStat


def run_word_export(docx_path: Path, pdf_path: Path) -> None:
    env = os.environ.copy()
    env["DOCX_QA_INPUT"] = str(docx_path)
    env["DOCX_QA_PDF"] = str(pdf_path)
    script = r"""
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $docx = [Environment]::GetEnvironmentVariable('DOCX_QA_INPUT', 'Process')
    $pdf = [Environment]::GetEnvironmentVariable('DOCX_QA_PDF', 'Process')
    $document = $word.Documents.Open($docx, $false, $true)
    try {
        $document.ExportAsFixedFormat($pdf, 17)
    } finally {
        $document.Close($false)
    }
} finally {
    $word.Quit()
}
"""
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        check=True,
        env=env,
        text=True,
    )


def render_pdf_with_pdftoppm(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    runtime_root = Path(r"C:\Users\wang_\.cache\codex-runtimes\codex-primary-runtime\dependencies")
    bundled_exe = runtime_root / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
    pdftoppm = str(bundled_exe) if bundled_exe.exists() else (shutil.which("pdftoppm") or shutil.which("pdftoppm.cmd"))
    if not pdftoppm:
        raise FileNotFoundError("pdftoppm was not found on PATH.")
    prefix = output_dir / "page"
    env = os.environ.copy()
    if bundled_exe.exists():
        env["PATH"] = str(bundled_exe.parent) + os.pathsep + env.get("PATH", "")
    subprocess.run([pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(prefix)], check=True, env=env)
    return sorted(output_dir.glob("page-*.png"))


def inspect_pages(pages: list[Path]) -> dict[str, object]:
    stats = []
    for page in pages:
        with Image.open(page) as image:
            rgb = image.convert("RGB")
            stat = ImageStat.Stat(rgb)
            mean = sum(stat.mean) / len(stat.mean)
            extrema = rgb.getextrema()
            non_blank = any(low < 250 for low, _ in extrema)
            stats.append(
                {
                    "file": str(page),
                    "width": image.width,
                    "height": image.height,
                    "mean_rgb": round(mean, 2),
                    "non_blank": bool(non_blank),
                }
            )
    return {
        "page_count": len(pages),
        "blank_like_pages": [item["file"] for item in stats if not item["non_blank"]],
        "pages": stats,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render DOCX using Microsoft Word and Poppler.")
    parser.add_argument("docx")
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--keep-pdf", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    docx_path = Path(args.docx).resolve()
    if not docx_path.exists():
        print(f"DOCX not found: {docx_path}", file=sys.stderr)
        return 2
    output_dir = Path(args.output_dir or docx_path.with_suffix("").name + "_render").resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = output_dir / "source.pdf"

    run_word_export(docx_path, pdf_path)
    pages = render_pdf_with_pdftoppm(pdf_path, output_dir, args.dpi)
    report = {
        "docx": str(docx_path),
        "pdf": str(pdf_path),
        "output_dir": str(output_dir),
        "render": inspect_pages(pages),
    }
    report_path = output_dir / "render_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if not args.keep_pdf:
        pdf_path.unlink(missing_ok=True)
        report["pdf"] = ""
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
