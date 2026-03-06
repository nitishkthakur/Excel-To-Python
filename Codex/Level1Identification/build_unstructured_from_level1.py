#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import stat
import sys
from pathlib import Path
from typing import Any


UNSTRUCTURED_CALCULATE_TEMPLATE = """#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

CODEX_ROOT = Path(r"{codex_root}")
if str(CODEX_ROOT) not in sys.path:
    sys.path.insert(0, str(CODEX_ROOT))

from excel_pipeline.runtime import run_unstructured_calculation

DEFAULT_MAPPING_REPORT = Path(r"{mapping_report}")
DEFAULT_INPUT = Path(r"{default_input}")
DEFAULT_OUTPUT = Path(r"{default_output}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rebuild output workbook from unstructured inputs"
    )
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
"""


def _safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    cleaned = cleaned.strip("._")
    return cleaned or "workbook"


def _import_pipeline_modules(codex_root: Path) -> tuple[Any, Any, Any, Any]:
    if str(codex_root) not in sys.path:
        sys.path.insert(0, str(codex_root))

    try:
        from excel_pipeline.compare import compare_workbooks
        from excel_pipeline.layer2_unstructured import create_unstructured_inputs
        from excel_pipeline.mapping import create_mapping_report
        from excel_pipeline.runtime import run_unstructured_calculation
    except Exception as exc:
        raise RuntimeError(
            "Failed to import excel_pipeline modules. "
            "Run with the project virtualenv, for example: ../.venv/bin/python "
            "build_unstructured_from_level1.py --file <level1_output.xlsx>"
        ) from exc

    return (
        create_mapping_report,
        create_unstructured_inputs,
        run_unstructured_calculation,
        compare_workbooks,
    )


def _write_unstructured_calculate_script(
    script_path: Path,
    codex_root: Path,
    mapping_report: Path,
    default_input: Path,
    default_output: Path,
) -> Path:
    script_path.parent.mkdir(parents=True, exist_ok=True)
    script_text = UNSTRUCTURED_CALCULATE_TEMPLATE.format(
        codex_root=str(codex_root.resolve()),
        mapping_report=str(mapping_report.resolve()),
        default_input=str(default_input.resolve()),
        default_output=str(default_output.resolve()),
    )
    script_path.write_text(script_text, encoding="utf-8")
    mode = script_path.stat().st_mode
    script_path.chmod(mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    return script_path


def build_unstructured_artifacts(
    source_workbook: Path,
    output_root: Path,
    cache_dir: Path,
    run_verify: bool,
) -> dict[str, Any]:
    source_workbook = source_workbook.resolve()
    output_root = output_root.resolve()
    cache_dir = cache_dir.resolve()

    if not source_workbook.exists():
        raise FileNotFoundError(f"Source workbook not found: {source_workbook}")
    if source_workbook.suffix.lower() not in {".xlsx", ".xls"}:
        raise ValueError("Supported source extensions are .xlsx and .xls")

    codex_root = Path(__file__).resolve().parent.parent
    (
        create_mapping_report,
        create_unstructured_inputs,
        run_unstructured_calculation,
        compare_workbooks,
    ) = _import_pipeline_modules(codex_root)

    model_dir = output_root / _safe_name(source_workbook.stem)
    model_dir.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    mapping_report = model_dir / "mapping_report.xlsx"
    unstructured_inputs = model_dir / "unstructured_inputs.xlsx"
    unstructured_calculate = model_dir / "unstructured_calculate.py"
    output_unstructured = model_dir / "output_unstructured.xlsx"

    create_mapping_report(
        source_workbook=source_workbook,
        mapping_report_path=mapping_report,
        cache_dir=cache_dir,
    )
    create_unstructured_inputs(mapping_report, unstructured_inputs)
    _write_unstructured_calculate_script(
        script_path=unstructured_calculate,
        codex_root=codex_root,
        mapping_report=mapping_report,
        default_input=unstructured_inputs,
        default_output=output_unstructured,
    )

    mismatches: list[dict[str, Any]] = []
    verification_error: str | None = None
    if run_verify:
        try:
            run_unstructured_calculation(mapping_report, unstructured_inputs, output_unstructured)
            raw_mismatches = compare_workbooks(source_workbook, output_unstructured)
            mismatches = [m.__dict__ for m in raw_mismatches]
        except Exception as exc:
            verification_error = str(exc)

    return {
        "source_workbook": str(source_workbook),
        "artifacts": {
            "mapping_report": str(mapping_report),
            "unstructured_inputs": str(unstructured_inputs),
            "unstructured_calculate": str(unstructured_calculate),
            "output_unstructured": str(output_unstructured if output_unstructured.exists() else ""),
        },
        "verification": {
            "enabled": run_verify,
            "success": (len(mismatches) == 0 and verification_error is None) if run_verify else None,
            "mismatch_count": len(mismatches) if run_verify and verification_error is None else None,
            "mismatches_preview": mismatches[:20] if run_verify else [],
            "error": verification_error,
        },
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate unstructured input template and unstructured_calculate.py "
            "for a Level1-hardcoded workbook."
        )
    )
    parser.add_argument(
        "--file",
        type=Path,
        required=True,
        help="Path to the Level1 output workbook (.xlsx/.xls)",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path(__file__).resolve().parent / "generated_unstructured",
        help="Directory where artifacts are created",
    )
    parser.add_argument(
        "--cache-dir",
        type=Path,
        default=Path(__file__).resolve().parent / ".cache" / "normalized",
        help="Cache directory for normalization and intermediates",
    )
    parser.add_argument(
        "--skip-verify",
        action="store_true",
        help="Skip runtime reproduction check against the source workbook",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        report = build_unstructured_artifacts(
            source_workbook=args.file,
            output_root=args.output_root,
            cache_dir=args.cache_dir,
            run_verify=not args.skip_verify,
        )
    except Exception as exc:
        print(f"Failed: {exc}", file=sys.stderr)
        return 1

    print(json.dumps(report, indent=2, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
