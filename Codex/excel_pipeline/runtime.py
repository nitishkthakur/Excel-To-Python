from __future__ import annotations

from pathlib import Path

from .vectorized_runtime import (
    run_structured_calculation_vectorized,
    run_unstructured_calculation_vectorized,
)


def run_unstructured_calculation(
    mapping_report_path: Path,
    unstructured_input_path: Path,
    output_path: Path,
) -> Path:
    return run_unstructured_calculation_vectorized(
        mapping_report_path=mapping_report_path,
        unstructured_input_path=unstructured_input_path,
        output_path=output_path,
    )


def run_structured_calculation(
    mapping_report_path: Path,
    structured_input_path: Path,
    output_path: Path,
) -> Path:
    return run_structured_calculation_vectorized(
        mapping_report_path=mapping_report_path,
        structured_input_path=structured_input_path,
        output_path=output_path,
    )
