#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.runtime import run_unstructured_calculation

DEFAULT_MAPPING_REPORT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts/TF7b509ab0-9b81-4f16-8e2f-b8fd20cd9fd56229c2f9_wac-bb7d2591b1d5/mapping_report.xlsx")
DEFAULT_INPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts/TF7b509ab0-9b81-4f16-8e2f-b8fd20cd9fd56229c2f9_wac-bb7d2591b1d5/unstructured_inputs.xlsx")
DEFAULT_OUTPUT = Path(r"/home/nitish/Documents/github/Excel-To-Python/Codex/artifacts/TF7b509ab0-9b81-4f16-8e2f-b8fd20cd9fd56229c2f9_wac-bb7d2591b1d5/output_unstructured.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser(description="Rebuild output workbook from unstructured inputs")
    parser.add_argument("--inputs", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--mapping", type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    run_unstructured_calculation(
        mapping_report_path=args.mapping,
        unstructured_input_path=args.inputs,
        output_path=args.output,
    )


if __name__ == "__main__":
    main()
