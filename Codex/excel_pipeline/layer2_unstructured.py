from __future__ import annotations

from pathlib import Path

from .mapping_io import read_mapping_report
from .reconstruct import save_workbook_from_mapping


def create_unstructured_inputs(mapping_report_path: Path, output_path: Path) -> Path:
    model = read_mapping_report(mapping_report_path)
    return save_workbook_from_mapping(
        model=model,
        output_path=output_path,
        include_formulas=False,
        unstructured_input_mode=True,
    )
