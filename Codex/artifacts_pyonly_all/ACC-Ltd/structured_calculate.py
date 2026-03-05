#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.runtime import run_structured_calculation

DEFAULT_MAPPING_REPORT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_pyonly_all/ACC-Ltd/mapping_report.xlsx")
DEFAULT_INPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_pyonly_all/ACC-Ltd/structured_input.xlsx")
DEFAULT_OUTPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts_pyonly_all/ACC-Ltd/output_structured.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser(description="Rebuild output workbook from structured inputs")
    parser.add_argument("--inputs", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    run_structured_calculation(
        mapping_report_path=args.mapping,
        structured_input_path=args.inputs,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
