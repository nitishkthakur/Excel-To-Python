#!/usr/bin/env python
"""
Excel-to-Mapping Report – CLI entry point.

Usage:
    python -m excel_to_mapping.main <excel_file> [--sheets Sheet1 Sheet2] [--output report.xlsx]
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from excel_to_mapping.mapper import generate_mapping_report


def main():
    parser = argparse.ArgumentParser(
        description="Generate an Inputs / Calculations / Outputs mapping report from an Excel workbook"
    )
    parser.add_argument(
        "excel_file",
        help="Path to the input Excel file (.xlsx)",
    )
    parser.add_argument(
        "--sheets",
        nargs="*",
        default=None,
        help="Sheet names to include (default: all sheets)",
    )
    parser.add_argument(
        "--config",
        default=None,
        help="Path to config YAML file",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output report path (default: ./output/mapping_report.xlsx)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.excel_file):
        print(f"Error: File '{args.excel_file}' not found.")
        sys.exit(1)

    generate_mapping_report(
        args.excel_file,
        sheet_names=args.sheets,
        config_path=args.config,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
