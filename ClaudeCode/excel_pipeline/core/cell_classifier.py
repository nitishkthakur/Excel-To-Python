"""Cell classification module for categorizing cells as Input/Calculation/Output."""

from typing import Optional
from openpyxl.cell import Cell
from excel_pipeline.core.dependency_graph import DependencyGraph
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


def classify_cell(cell: Cell, dep_graph: DependencyGraph, sheet_name: Optional[str] = None) -> str:
    """
    Classify a cell as Input, Calculation, or Output.

    Classification rules:
    - Input: Cell with no formula (hardcoded value or empty)
    - Calculation: Cell with formula that is referenced by other formulas
    - Output: Cell with formula that is NOT referenced by other formulas (terminal node)

    Args:
        cell: openpyxl Cell object
        dep_graph: Dependency graph containing precedent/dependent relationships
        sheet_name: Sheet name (optional, uses cell.parent.title if not provided)

    Returns:
        Cell type: "Input", "Calculation", or "Output"

    Examples:
        >>> classify_cell(sheet['A1'], dep_graph)
        'Input'
        >>> classify_cell(sheet['B1'], dep_graph)  # =A1*2, used in C1
        'Calculation'
        >>> classify_cell(sheet['C1'], dep_graph)  # =B1+10, not used anywhere
        'Output'
    """
    if sheet_name is None:
        sheet_name = cell.parent.title

    # Use dependency graph's classification logic
    return dep_graph.classify_cell(sheet_name, cell.coordinate)


def get_cell_type_counts(workbook, dep_graph: DependencyGraph) -> dict:
    """
    Count cells by type across entire workbook.

    Args:
        workbook: openpyxl Workbook object
        dep_graph: Dependency graph

    Returns:
        Dictionary with counts: {'Input': n, 'Calculation': n, 'Output': n}

    Example:
        >>> counts = get_cell_type_counts(wb, dep_graph)
        >>> print(f"Inputs: {counts['Input']}, Calculations: {counts['Calculation']}")
    """
    counts = {'Input': 0, 'Calculation': 0, 'Output': 0}

    for sheet in workbook.worksheets:
        sheet_name = sheet.title

        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:  # Skip empty cells
                    cell_type = classify_cell(cell, dep_graph, sheet_name)
                    counts[cell_type] += 1

    return counts


def is_input_cell(cell: Cell, dep_graph: DependencyGraph, sheet_name: Optional[str] = None) -> bool:
    """Check if cell is an Input cell."""
    return classify_cell(cell, dep_graph, sheet_name) == "Input"


def is_calculation_cell(cell: Cell, dep_graph: DependencyGraph, sheet_name: Optional[str] = None) -> bool:
    """Check if cell is a Calculation cell."""
    return classify_cell(cell, dep_graph, sheet_name) == "Calculation"


def is_output_cell(cell: Cell, dep_graph: DependencyGraph, sheet_name: Optional[str] = None) -> bool:
    """Check if cell is an Output cell."""
    return classify_cell(cell, dep_graph, sheet_name) == "Output"
