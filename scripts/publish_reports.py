#!/usr/bin/env python3
"""Publish Shift Flight Deck outputs (PDFs and workbook snapshot)."""
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
import shutil

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_WORKBOOK = REPO_ROOT / "excel" / "Shift_Flight_Deck.xlsm"
EXPORT_DIR = REPO_ROOT / "exports"
LOG_PATH = REPO_ROOT / "data" / "logs" / "publish.log"


def export_pdf_via_com(workbook_path: Path, sheets: list[str]):
    try:
        import win32com.client  # type: ignore
    except Exception:
        return [f"win32com unavailable: skipped PDF export for {','.join(sheets)}"]

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    wb = xl.Workbooks.Open(str(workbook_path))
    outputs = []
    try:
        for sheet in sheets:
            out = EXPORT_DIR / f"{sheet}_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            wb.Worksheets(sheet).ExportAsFixedFormat(0, str(out))
            outputs.append(str(out))
    finally:
        wb.Close(SaveChanges=False)
        xl.Quit()
    return outputs


def write_shift_summary(summary_path: Path):
    summary_path.write_text(
        "Shift Summary\n"
        "- Top issues by line: see Analysis_Report\n"
        "- Key downtime events: see Downtime_Log\n"
        "- Forecast vs plan: see Dash_Shift\n"
        "- Top 3 prompts: see Analysis_Report rules section\n",
        encoding="utf-8",
    )


def publish(workbook_path: Path):
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot = EXPORT_DIR / f"Shift_Flight_Deck_{ts}.xlsm"
    shutil.copy2(workbook_path, snapshot)
    pdfs = export_pdf_via_com(workbook_path, ["Dash_Shift", "Dash_Trends"])
    summary = EXPORT_DIR / f"Shift_Summary_{ts}.txt"
    write_shift_summary(summary)

    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    LOG_PATH.write_text(f"{dt.datetime.now().isoformat()} published snapshot={snapshot} pdfs={pdfs} summary={summary}\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK))
    args = parser.parse_args()
    publish(Path(args.workbook))
    print("Publish complete")


if __name__ == "__main__":
    main()
