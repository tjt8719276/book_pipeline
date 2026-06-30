#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Merge cleaned volume chapters and build MD/DOCX/PDF outputs.

Usage:
  python merge_and_build.py <source_dir> <volume_title> <output_name> [output_dir]
"""

import glob
import os
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUILD_DOCX = SCRIPT_DIR / "build_docx.py"


def resolve_clean_dir(source_dir):
    candidates = [
        source_dir / "clean_chapters",
        source_dir / "output" / "clean_chapters",
    ]
    for candidate in candidates:
        if candidate.is_dir():
            return candidate
    raise FileNotFoundError(f"clean_chapters directory not found under {source_dir}")


def merge_chapters(source_dir, volume_title):
    clean_dir = resolve_clean_dir(source_dir)
    merged_dir = source_dir / "output"
    merged_dir.mkdir(parents=True, exist_ok=True)
    merged_path = merged_dir / "merged_chapter.md"

    md_files = sorted(clean_dir.glob("[01][0-9]_*.md"))
    if not md_files:
        raise FileNotFoundError(f"No chapter files found in {clean_dir}")

    print(f"[1/5] Merging {len(md_files)} chapter files...")
    with merged_path.open("w", encoding="utf-8") as out:
        out.write(f"# {volume_title}\n\n")
        for file_path in md_files:
            print(f"       + {file_path.name}")
            out.write(file_path.read_text(encoding="utf-8").strip())
            out.write("\n\n")

    size_kb = merged_path.stat().st_size / 1024
    print(f"       -> merged_chapter.md ({size_kb:.1f} KB)")
    return merged_path


def build_docx(merged_path, output_name, output_dir, volume_title):
    docx_path = output_dir / f"{output_name}.docx"
    print("\n[2/5] Building DOCX...")
    result = subprocess.run(
        [sys.executable, str(BUILD_DOCX), str(merged_path), str(docx_path), volume_title],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"build_docx.py failed:\n{result.stderr}")
    size_kb = docx_path.stat().st_size / 1024
    print(f"       -> DOCX ({size_kb:.1f} KB)")
    return docx_path


def convert_pdf(docx_path):
    pdf_path = docx_path.with_suffix(".pdf")
    print("\n[3/5] Converting to PDF...")
    try:
        import win32com.client  # type: ignore

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(str(docx_path))
            doc.Fields.Update()
            doc.TablesOfContents(1).Update() if doc.TablesOfContents.Count else None
            doc.Save()
            doc.SaveAs(str(pdf_path), FileFormat=17)
            doc.Close()
        finally:
            word.Quit()
        size_kb = pdf_path.stat().st_size / 1024
        print(f"       -> PDF ({size_kb:.1f} KB)")
        return pdf_path
    except ImportError:
        print("       [WARN] pywin32 not installed, skipping PDF conversion")
    except Exception as exc:
        print(f"       [WARN] PDF conversion failed: {exc}")
    return None


def copy_md(merged_path, output_name, output_dir):
    md_path = output_dir / f"{output_name}.md"
    print("\n[4/5] Copying markdown...")
    shutil.copyfile(merged_path, md_path)
    print(f"       -> {md_path}")
    return md_path


def main():
    if len(sys.argv) < 4:
        print("Usage: python merge_and_build.py <source_dir> <volume_title> <output_name> [output_dir]")
        sys.exit(1)

    source_dir = Path(sys.argv[1]).resolve()
    volume_title = sys.argv[2]
    output_name = sys.argv[3]
    output_dir = Path(sys.argv[4]).resolve() if len(sys.argv) >= 5 else Path(r"d:\desk\new")

    output_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 72)
    print("  Book Volume Pipeline")
    print(f"  Source:  {source_dir}")
    print(f"  Title:   {volume_title}")
    print(f"  Output:  {output_dir}\\{output_name}.{{md,docx,pdf}}")
    print("=" * 72)

    merged_path = merge_chapters(source_dir, volume_title)
    docx_path = build_docx(merged_path, output_name, output_dir, volume_title)
    pdf_path = convert_pdf(docx_path)
    md_path = copy_md(merged_path, output_name, output_dir)

    print("\n[5/5] Done!")
    print(f"  MD:   {md_path}")
    print(f"  DOCX: {docx_path}")
    if pdf_path:
        print(f"  PDF:  {pdf_path}")
    print("=" * 72)


if __name__ == "__main__":
    main()
