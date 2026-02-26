#!/usr/bin/env python
"""
Excel-to-Python Vectorised Converter – CLI entry point.

Usage:
    python -m excel_to_python_vectorized.main <excel_file> [--config config.yaml] [--output-dir output]

Or, from the repository root:
    python excel_to_python_vectorized/main.py <excel_file> [--output-dir output]
"""

import argparse
import os
import sys

# Ensure the repository root is on the path so both the package and
# the original ``formula_converter`` / ``excel_to_python`` modules
# can be imported.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from excel_to_python_vectorized.converter import convert_excel_to_python_vectorized


def main():
    parser = argparse.ArgumentParser(
        description="Convert an Excel workbook to a vectorised Python script"
    )
    parser.add_argument(
        "excel_file",
        help="Path to the input Excel file (.xlsx)",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="Path to config YAML file (default: config.yaml in repo root)",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Output directory (default: ./output next to the Excel file)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.excel_file):
        print(f"Error: File '{args.excel_file}' not found.")
        sys.exit(1)

    convert_excel_to_python_vectorized(
        args.excel_file,
        config_path=args.config,
        output_dir=args.output_dir,
    )


if __name__ == "__main__":
    main()
