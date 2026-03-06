#!/usr/bin/env python3
"""
Run the Excel-to-Python pipeline for a single sheet with cross-sheet support.

Usage:
    python run_this_onesheet.py <excel_file> <sheet_name>

Examples:
    python run_this_onesheet.py ExcelFiles/Indigo.xlsx "P&L"
    python run_this_onesheet.py ExcelFiles/ACC-Ltd.xlsx "Balance Sheet"
    python run_this_onesheet.py ExcelFiles/Bharti-Airtel\\(2\\).xlsx "Income Statement"

Why this is faster than run_this.py
-------------------------------------
All pipeline stages operate on a workbook that contains one sheet with
formulas (the target) plus frozen value-snapshots of every other sheet.
That eliminates formula analysis, dependency-graph traversal, and code
generation for every sheet except the one you care about.

How cross-sheet references are handled
----------------------------------------
Instead of stripping other sheets (which breaks cross-sheet formulas), this
script uses excel_pipeline.onesheet.pipeline which:
  1. Replaces every formula in non-target sheets with its last-cached value,
     creating intermediate/frozen_workbook.xlsx on disk for inspection.
  2. Runs Layer 1 on the frozen workbook so non-target sheet cells appear as
     "Input" in the mapping report.
  3. Layer 2a includes those frozen values in unstructured_inputs.xlsx.
  4. The generated Python script resolves cross-sheet references from the
     cell store c, which was pre-populated with the frozen values.

Artifacts produced in  output/<filename>_<sheet>/
    intermediate/
        frozen_workbook.xlsx   intermediate: other sheets hardcoded to cached values
    mapping_report.xlsx        cell metadata & classifications
    unstructured_inputs.xlsx   input values + frozen cross-sheet values
    unstructured_calculate.py  generated Python calculation script
    output.xlsx                fully populated Excel output for the target sheet
    pipeline.log               detailed run log
"""

import sys
import os
import time
from pathlib import Path

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_pipeline.onesheet.pipeline import run as run_onesheet


def _safe_name(name: str) -> str:
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in name)


def main() -> None:
    if len(sys.argv) != 3:
        print("Usage: python run_this_onesheet.py <excel_file> <sheet_name>")
        print()
        print("Examples:")
        print('  python run_this_onesheet.py ExcelFiles/Indigo.xlsx "P&L"')
        print('  python run_this_onesheet.py ExcelFiles/ACC-Ltd.xlsx "Balance Sheet"')
        sys.exit(1)

    input_path = sys.argv[1]
    sheet_name = sys.argv[2]

    input_file = Path(input_path)
    if not input_file.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    # Check extension — openpyxl only supports .xlsx / .xlsm
    if input_file.suffix.lower() == ".xls":
        print(f"Error: .xls format is not supported (openpyxl requires .xlsx).")
        print("Open the file in Excel and save as .xlsx first.")
        sys.exit(1)

    # Validate sheet name before doing any heavy work
    print(f"Checking available sheets...")
    probe = openpyxl.load_workbook(str(input_file), read_only=True, data_only=True)
    available = probe.sheetnames
    probe.close()

    if sheet_name not in available:
        print(f"Error: Sheet '{sheet_name}' not found in {input_path}")
        print(f"Available sheets ({len(available)}):")
        for s in available:
            print(f"  • {s}")
        sys.exit(1)

    stem = input_file.stem
    safe_sheet = _safe_name(sheet_name)
    output_dir = Path("output") / f"{stem}_{safe_sheet}"

    start = time.time()
    print("=" * 70)
    print("Excel-to-Python Pipeline  (single sheet, cross-sheet refs preserved)")
    print(f"Input:  {input_path}")
    print(f"Sheet:  {sheet_name}")
    print(f"Output: {output_dir}/")
    print("=" * 70)

    try:
        run_onesheet(
            input_path=str(input_file),
            sheet_name=sheet_name,
            output_dir=output_dir,
            log_level="INFO",
        )
    except (ValueError, FileNotFoundError) as exc:
        print(f"\nError: {exc}")
        sys.exit(1)
    except RuntimeError as exc:
        print(f"\nPipeline error: {exc}")
        sys.exit(1)

    total = time.time() - start
    print()
    print("=" * 70)
    print(f"Done in {total:.1f}s  |  Artifacts in {output_dir}/")
    print()
    print("  intermediate/frozen_workbook.xlsx  other sheets hardcoded to cached values")
    print("  mapping_report.xlsx                cell metadata & classifications")
    print("  unstructured_inputs.xlsx           inputs + frozen cross-sheet values")
    print("  unstructured_calculate.py          generated Python calculation script")
    print("  output.xlsx                        fully populated Excel output")
    print("=" * 70)


if __name__ == "__main__":
    main()
