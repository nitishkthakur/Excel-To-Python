#!/usr/bin/env python
"""
Excel-to-Mapping Report – CLI entry point.

Usage:
    # Generate mapping report from an Excel workbook
    python -m excel_to_mapping.main map <excel_file> [--sheets Sheet1 Sheet2] [--output report.xlsx]

    # Generate input template from a mapping report
    python -m excel_to_mapping.main template <mapping_file> [--output input_template.xlsx]

    # Regenerate Excel workbook from a mapping report
    python -m excel_to_mapping.main regenerate <mapping_file> [--output regenerated.xlsx] [--inputs input_template.xlsx]
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from excel_to_mapping.mapper import generate_mapping_report
from excel_to_mapping.regenerator import regenerate_workbook, generate_input_template


def main():
    parser = argparse.ArgumentParser(
        description="Excel ↔ Mapping Report: generate, review, and regenerate"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # ---- map ----
    p_map = sub.add_parser(
        "map",
        help="Generate a tabular mapping report from an Excel workbook",
    )
    p_map.add_argument("excel_file", help="Path to the input Excel file (.xlsx)")
    p_map.add_argument(
        "--sheets", nargs="*", default=None,
        help="Sheet names to include (default: all sheets)",
    )
    p_map.add_argument("--config", default=None, help="Path to config YAML file")
    p_map.add_argument(
        "--output", default=None,
        help="Output report path (default: ./output/mapping_report.xlsx)",
    )

    # ---- template ----
    p_tpl = sub.add_parser(
        "template",
        help="Generate an input-values template from a mapping report",
    )
    p_tpl.add_argument("mapping_file", help="Path to the mapping report (.xlsx)")
    p_tpl.add_argument(
        "--output", default=None,
        help="Output template path (default: ./output/input_template.xlsx)",
    )

    # ---- regenerate ----
    p_reg = sub.add_parser(
        "regenerate",
        help="Regenerate an Excel workbook from a mapping report",
    )
    p_reg.add_argument("mapping_file", help="Path to the mapping report (.xlsx)")
    p_reg.add_argument(
        "--output", default=None,
        help="Output workbook path (default: ./output/regenerated.xlsx)",
    )
    p_reg.add_argument(
        "--inputs", default=None,
        help="Path to an input template with overridden values",
    )

    args = parser.parse_args()

    if args.command == "map":
        if not os.path.exists(args.excel_file):
            print(f"Error: File '{args.excel_file}' not found.")
            sys.exit(1)
        generate_mapping_report(
            args.excel_file,
            sheet_names=args.sheets,
            config_path=args.config,
            output_path=args.output,
        )

    elif args.command == "template":
        if not os.path.exists(args.mapping_file):
            print(f"Error: File '{args.mapping_file}' not found.")
            sys.exit(1)
        out = args.output or os.path.join("output", "input_template.xlsx")
        generate_input_template(args.mapping_file, out)

    elif args.command == "regenerate":
        if not os.path.exists(args.mapping_file):
            print(f"Error: File '{args.mapping_file}' not found.")
            sys.exit(1)
        out = args.output or os.path.join("output", "regenerated.xlsx")
        regenerate_workbook(args.mapping_file, out, input_values_path=args.inputs)


if __name__ == "__main__":
    main()
