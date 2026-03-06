from __future__ import annotations

import argparse
import importlib.util
import json
import re
from pathlib import Path
from typing import Any

from .compare import compare_workbooks
from .layer2_unstructured import create_unstructured_inputs
from .mapping import create_mapping_report
from .mapping_io import read_mapping_report
from .python_codegen import generate_unstructured_python_engine


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    cleaned = cleaned.strip("._")
    return cleaned or "workbook"


def run_unstructured_python_pipeline_for_workbook(
    source_workbook: Path,
    output_root: Path,
    cache_dir: Path,
    python_executable: Path | None = None,
) -> dict[str, Any]:
    output_root.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    model_dir = output_root / _safe_name(source_workbook.stem)
    model_dir.mkdir(parents=True, exist_ok=True)

    mapping_report = model_dir / "mapping_report.xlsx"
    unstructured_inputs = model_dir / "unstructured_inputs.xlsx"
    python_engine_script = model_dir / "unstructured_python_engine.py"
    output_python = model_dir / "unstructured_output_python.xlsx"

    create_mapping_report(
        source_workbook=source_workbook,
        mapping_report_path=mapping_report,
        cache_dir=cache_dir,
    )

    create_unstructured_inputs(
        mapping_report_path=mapping_report,
        output_path=unstructured_inputs,
    )

    generate_unstructured_python_engine(
        mapping_report_path=mapping_report,
        output_script_path=python_engine_script,
        default_input_path=unstructured_inputs,
        default_output_path=output_python,
    )

    if python_executable is not None:
        _ = python_executable  # preserved for API compatibility

    module_name = f"_generated_engine_{_safe_name(source_workbook.stem)}"
    spec = importlib.util.spec_from_file_location(module_name, python_engine_script.resolve())
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load generated engine: {python_engine_script}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.run(
        mapping_report_path=mapping_report,
        inputs_path=unstructured_inputs,
        output_path=output_python,
    )

    model = read_mapping_report(mapping_report)
    baseline = model.normalized_workbook
    mismatches = compare_workbooks(baseline, output_python)

    report = {
        "source_workbook": str(source_workbook),
        "artifacts": {
            "mapping_report": str(mapping_report),
            "unstructured_inputs": str(unstructured_inputs),
            "unstructured_python_engine": str(python_engine_script),
            "output_unstructured_python": str(output_python),
        },
        "mismatches": [m.__dict__ for m in mismatches],
        "success": not mismatches,
    }

    report_path = model_dir / "python_engine_validation_report.json"
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=True), encoding="utf-8")

    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Run unstructured Excel-to-Python pipeline and execute generated Python engine "
            "to produce unstructured_output_python.xlsx"
        )
    )
    parser.add_argument("--file", type=Path, required=True, help="Path to source .xls/.xlsx")
    parser.add_argument("--output-root", type=Path, default=Path("artifacts_python_engine"))
    parser.add_argument("--cache-dir", type=Path, default=Path(".cache/normalized"))
    args = parser.parse_args(argv)

    result = run_unstructured_python_pipeline_for_workbook(
        source_workbook=args.file,
        output_root=args.output_root,
        cache_dir=args.cache_dir,
    )
    print(json.dumps(result, indent=2, ensure_ascii=True))
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
