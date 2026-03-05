#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from excel_pipeline.codegen import generate_structured_calculate, generate_unstructured_calculate
from excel_pipeline.layer2_structured import create_structured_input
from excel_pipeline.layer2_unstructured import create_unstructured_inputs
from excel_pipeline.mapping import create_mapping_report
from excel_pipeline.runtime import run_structured_calculation, run_unstructured_calculation


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    cleaned = cleaned.strip("._")
    return cleaned or "workbook"


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Run selected pipeline sections for a single workbook: "
            "1, 1-2a, 1-2b, 1-2a-3a, 1-2b-3b"
        )
    )
    parser.add_argument("--file", type=Path, required=True, help="Path to source .xls/.xlsx")
    parser.add_argument(
        "--mode",
        required=True,
        choices=["1", "1-2a", "1-2b", "1-2a-3a", "1-2b-3b"],
        help="Pipeline path to execute",
    )
    parser.add_argument("--output-root", type=Path, default=Path("artifacts_sections"))
    parser.add_argument("--cache-dir", type=Path, default=Path(".cache/normalized"))
    args = parser.parse_args()

    output_root = args.output_root.resolve()
    cache_dir = args.cache_dir.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    model_dir = output_root / _safe_name(args.file.stem)
    model_dir.mkdir(parents=True, exist_ok=True)

    mapping_report = model_dir / "mapping_report.xlsx"
    unstructured_inputs = model_dir / "unstructured_inputs.xlsx"
    structured_input = model_dir / "structured_input.xlsx"
    unstructured_calculate = model_dir / "unstructured_calculate.py"
    structured_calculate = model_dir / "structured_calculate.py"
    output_unstructured = model_dir / "output_unstructured.xlsx"
    output_structured = model_dir / "output_structured.xlsx"

    create_mapping_report(args.file, mapping_report, cache_dir)

    if args.mode in {"1-2a", "1-2a-3a"}:
        create_unstructured_inputs(mapping_report, unstructured_inputs)

    if args.mode in {"1-2b", "1-2b-3b"}:
        create_structured_input(mapping_report, structured_input)

    if args.mode == "1-2a-3a":
        generate_unstructured_calculate(
            mapping_report_path=mapping_report,
            output_script_path=unstructured_calculate,
            default_input_path=unstructured_inputs,
            default_output_path=output_unstructured,
        )
        run_unstructured_calculation(mapping_report, unstructured_inputs, output_unstructured)

    if args.mode == "1-2b-3b":
        generate_structured_calculate(
            mapping_report_path=mapping_report,
            output_script_path=structured_calculate,
            default_input_path=structured_input,
            default_output_path=output_structured,
        )
        run_structured_calculation(mapping_report, structured_input, output_structured)

    artifacts = {
        "mode": args.mode,
        "source": str(args.file.resolve()),
        "mapping_report": str(mapping_report),
        "unstructured_inputs": str(unstructured_inputs if unstructured_inputs.exists() else ""),
        "structured_input": str(structured_input if structured_input.exists() else ""),
        "unstructured_calculate": str(
            unstructured_calculate if unstructured_calculate.exists() else ""
        ),
        "structured_calculate": str(
            structured_calculate if structured_calculate.exists() else ""
        ),
        "output_unstructured": str(output_unstructured if output_unstructured.exists() else ""),
        "output_structured": str(output_structured if output_structured.exists() else ""),
    }
    print(json.dumps(artifacts, indent=2, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
