"""
Dependency Graph Module
=======================
Builds a dependency graph of all cells and performs topological sort
to determine the correct evaluation order.
"""

import logging
from collections import defaultdict, deque
from .excel_parser import WorkbookInfo
from .formula_translator import extract_cell_references

logger = logging.getLogger(__name__)


def build_dependency_graph(workbook_info: WorkbookInfo):
    """
    Build a dependency graph from the parsed workbook.

    Returns:
        adjacency: dict mapping cell_key -> set of cell_keys it depends on
        reverse_adj: dict mapping cell_key -> set of cell_keys that depend on it
    """
    adjacency = defaultdict(set)  # cell -> cells it depends on
    reverse_adj = defaultdict(set)  # cell -> cells that depend on it

    for cell_key, cell_info in workbook_info.all_cells.items():
        if cell_info.is_formula and cell_info.formula:
            refs = extract_cell_references(cell_info.formula, cell_info.sheet)
            for ref in refs:
                adjacency[cell_key].add(ref)
                reverse_adj[ref].add(cell_key)

    return adjacency, reverse_adj


def topological_sort(workbook_info: WorkbookInfo, adjacency: dict) -> list:
    """
    Perform topological sort on the dependency graph.
    Returns cells in evaluation order (dependencies first).

    Uses Kahn's algorithm to handle the graph properly.
    Handles cycles by breaking them and logging warnings.
    """
    # Gather all cells that are formulas (need to be computed)
    formula_cells = set()
    for cell_key, cell_info in workbook_info.all_cells.items():
        if cell_info.is_formula:
            formula_cells.add(cell_key)

    # In-degree count for formula cells
    in_degree = defaultdict(int)
    for cell in formula_cells:
        for dep in adjacency.get(cell, set()):
            if dep in formula_cells:
                in_degree[cell] += 1

    # Start with formula cells that have no formula dependencies
    queue = deque()
    for cell in formula_cells:
        if in_degree[cell] == 0:
            queue.append(cell)

    sorted_cells = []
    visited = set()

    while queue:
        cell = queue.popleft()
        if cell in visited:
            continue
        visited.add(cell)
        sorted_cells.append(cell)

        # Find cells that depend on this cell
        for dependent in formula_cells:
            if cell in adjacency.get(dependent, set()):
                in_degree[dependent] -= 1
                if in_degree[dependent] <= 0 and dependent not in visited:
                    queue.append(dependent)

    # Handle any remaining cells (circular dependencies)
    remaining = formula_cells - visited
    if remaining:
        logger.warning(f"Found {len(remaining)} cells in circular dependencies. "
                       f"Breaking cycles by adding them in order.")
        # Add remaining in sheet order
        remaining_sorted = sorted(remaining, key=lambda x: (x[0], x[1], x[2]))
        sorted_cells.extend(remaining_sorted)

    return sorted_cells


def find_referenced_hardcoded(workbook_info: WorkbookInfo, adjacency: dict, reverse_adj: dict) -> set:
    """
    Find all hardcoded values that are referenced by at least one formula.
    Returns a set of (sheet, row, col) keys.
    """
    referenced = set()
    for cell_key, cell_info in workbook_info.all_cells.items():
        if cell_info.is_hardcoded_number or cell_info.is_label:
            if cell_key in reverse_adj and len(reverse_adj[cell_key]) > 0:
                referenced.add(cell_key)
    return referenced


def find_unreferenced_hardcoded(workbook_info: WorkbookInfo, reverse_adj: dict) -> set:
    """
    Find all hardcoded numeric values that are NOT referenced by any formula.
    """
    unreferenced = set()
    for cell_key, cell_info in workbook_info.all_cells.items():
        if cell_info.is_hardcoded_number:
            if cell_key not in reverse_adj or len(reverse_adj[cell_key]) == 0:
                unreferenced.add(cell_key)
    return unreferenced
