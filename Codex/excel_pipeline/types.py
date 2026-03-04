from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class CellRecord:
    sheet: str
    row: int
    column: int
    cell: str
    cell_type: str
    formula: str | None
    value: Any
    include_flag: bool
    number_format: str | None
    font_bold: bool | None
    font_italic: bool | None
    font_size: float | None
    font_color: str | None
    fill_color: str | None
    horizontal_alignment: str | None
    vertical_alignment: str | None
    wrap_text: bool | None
    style_json: str | None
    value_json: str
    group_id: str | None = None
    group_direction: str | None = None
    group_size: int | None = None
    pattern_formula: str | None = None


@dataclass
class SheetLayout:
    title: str
    index: int
    merged_ranges: list[str] = field(default_factory=list)
    freeze_panes: str | None = None
    tab_color: str | None = None
    row_dimensions: list[dict[str, Any]] = field(default_factory=list)
    column_dimensions: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class MappingModel:
    source_workbook: Path
    normalized_workbook: Path
    sheet_order: list[str]
    layouts: dict[str, SheetLayout]
    cells_by_sheet: dict[str, list[CellRecord]]
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class StructuredMappingRow:
    source_sheet: str
    source_cell: str
    input_sheet: str
    input_cell: str
    table_name: str
    entry_type: str
    is_transposed: bool
    patch_bounds: str
    metric_label: str | None
    period_label: str | None


@dataclass
class PipelineArtifacts:
    mapping_report: Path
    unstructured_inputs: Path
    structured_input: Path
    unstructured_calculate: Path
    structured_calculate: Path
    output_unstructured: Path
    output_structured: Path
