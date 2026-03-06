#!/usr/bin/env python3
"""
Run the full Excel-to-Python pipeline for a given file.

Usage:
    python run_this.py <excel_file>

Example:
    python run_this.py ExcelFiles/Indigo.xlsx

Artifacts produced in output/<filename>/:
    mapping_report.xlsx        - Cell metadata and classifications (single source of truth)
    unstructured_inputs.xlsx   - Input cells in original layout (editable template)
    unstructured_calculate.py  - Generated Python calculation script
    output.xlsx                - Fully populated Excel output
"""

import sys
import os
import subprocess
import time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_pipeline.utils.config import config
from excel_pipeline.utils.logging_setup import setup_logging
from excel_pipeline.layer1.parser import generate_mapping_report
from excel_pipeline.layer2.unstructured_generator import generate_unstructured_inputs
from excel_pipeline.layer4a.code_generator import generate_unstructured_code


def run_pipeline(input_path: str) -> None:
    input_file = Path(input_path)
    if not input_file.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    stem = input_file.stem
    output_dir = Path("output") / stem
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        config.load("config.yaml")
    except FileNotFoundError:
        pass

    setup_logging(level="INFO", log_file=str(output_dir / "pipeline.log"))

    mapping_path  = str(output_dir / "mapping_report.xlsx")
    inputs_path   = str(output_dir / "unstructured_inputs.xlsx")
    script_path   = str(output_dir / "unstructured_calculate.py")
    output_path   = str(output_dir / "output.xlsx")

    start = time.time()
    print("=" * 70)
    print("Excel-to-Python Pipeline")
    print(f"Input:  {input_path}")
    print(f"Output: {output_dir}/")
    print("=" * 70)

    # Layer 1 — mapping report
    print("\n[1/4] Layer 1: Mapping report...")
    t0 = time.time()
    generate_mapping_report(str(input_file), mapping_path)
    print(f"      {time.time()-t0:.1f}s  →  {mapping_path}")

    # Layer 2a — unstructured inputs
    print("\n[2/4] Layer 2a: Unstructured inputs...")
    t0 = time.time()
    generate_unstructured_inputs(mapping_path, inputs_path)
    print(f"      {time.time()-t0:.1f}s  →  {inputs_path}")

    # Layer 4a — generate Python script
    print("\n[3/4] Layer 4a: Generating calculation script...")
    t0 = time.time()
    generate_unstructured_code(mapping_path, inputs_path, script_path)
    print(f"      {time.time()-t0:.1f}s  →  {script_path}")

    # Execute generated script → output.xlsx
    # The generated script uses relative paths ("unstructured_inputs.xlsx", "output.xlsx"),
    # so we run it from output_dir where both files live.
    print("\n[4/4] Running generated script → output.xlsx...")
    t0 = time.time()
    result = subprocess.run(
        [sys.executable, "unstructured_calculate.py"],
        cwd=str(output_dir),
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print("      ERROR running generated script:")
        print(result.stderr)
        sys.exit(1)
    print(f"      {time.time()-t0:.1f}s  →  {output_path}")

    total = time.time() - start
    print("\n" + "=" * 70)
    print(f"Done in {total:.1f}s  |  Artifacts in {output_dir}/")
    print("  mapping_report.xlsx        cell metadata & classifications")
    print("  unstructured_inputs.xlsx   input values template")
    print("  unstructured_calculate.py  generated Python calculation script")
    print("  output.xlsx                fully populated Excel output")
    print("=" * 70)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python run_this.py <excel_file>")
        print()
        print("Examples:")
        print("  python run_this.py ExcelFiles/Indigo.xlsx")
        print("  python run_this.py ExcelFiles/Bharti-Airtel\\(2\\).xlsx")
        sys.exit(1)

    run_pipeline(sys.argv[1])
