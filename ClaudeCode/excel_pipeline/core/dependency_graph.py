"""Dependency graph builder for cell classification and calculation ordering."""

import re
from typing import Dict, Set, List, Optional, Tuple
from collections import defaultdict, deque
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import range_boundaries
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


class DependencyGraph:
    """
    Build and manage dependency graph for Excel workbook.

    The graph tracks:
    - Precedents: cells that a formula references
    - Dependents: cells that reference a given cell
    - Formula cells: all cells containing formulas
    """

    def __init__(self, workbook: Workbook):
        """
        Initialize dependency graph.

        Args:
            workbook: Excel workbook to analyze
        """
        self.workbook = workbook
        self.precedents: Dict[str, Set[str]] = defaultdict(set)  # cell -> cells it references
        self.dependents: Dict[str, Set[str]] = defaultdict(set)  # cell -> cells referencing it
        self.formula_cells: Set[str] = set()
        self.circular_refs: List[List[str]] = []

    def build(self) -> None:
        """
        Build the dependency graph from workbook.

        Parses all formulas and extracts cell references to build
        precedent and dependent relationships.
        """
        logger.info("Building dependency graph...")

        total_formulas = 0

        for sheet in self.workbook.worksheets:
            sheet_name = sheet.title

            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f':  # Formula cell
                        cell_ref = self._make_cell_ref(sheet_name, cell.coordinate)
                        self.formula_cells.add(cell_ref)

                        formula = str(cell.value)
                        refs = self._extract_cell_references(formula, sheet_name)

                        for ref in refs:
                            self.precedents[cell_ref].add(ref)
                            self.dependents[ref].add(cell_ref)

                        total_formulas += 1

        logger.info(f"Dependency graph built: {total_formulas} formulas, "
                   f"{len(self.precedents)} cells with precedents")

        # Detect circular references
        self._detect_circular_references()

    def _make_cell_ref(self, sheet: str, coord: str) -> str:
        """Create full cell reference with sheet name."""
        return f"{sheet}!{coord}"

    def _extract_cell_references(self, formula: str, current_sheet: str) -> Set[str]:
        """
        Extract all cell references from a formula.

        Args:
            formula: Excel formula string
            current_sheet: Current sheet name (for relative references)

        Returns:
            Set of cell references (Sheet!Cell format)

        Examples:
            >>> refs = self._extract_cell_references("=A1+B2", "Sheet1")
            >>> # Returns: {"Sheet1!A1", "Sheet1!B2"}
        """
        refs = set()

        # Remove leading '='
        if formula.startswith('='):
            formula = formula[1:]

        # Pattern for cell references (with optional sheet name)
        # Matches: A1, $A$1, Sheet1!A1, 'Sheet 1'!A1
        pattern = r"(?:'([^']+)'|([^\s!]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)"

        for match in re.finditer(pattern, formula):
            if match.group(1):  # Sheet with quotes: 'Sheet 1'!A1
                sheet = match.group(1)
                cell = match.group(3)
            elif match.group(2):  # Sheet without quotes: Sheet1!A1
                sheet = match.group(2)
                cell = match.group(3)
            elif match.group(4):  # No sheet: A1
                sheet = current_sheet
                cell = match.group(4)
            else:
                continue

            # Remove $ signs (absolute reference markers)
            cell = cell.replace('$', '')

            refs.add(self._make_cell_ref(sheet, cell))

        # Handle range references (A1:B10)
        range_pattern = r"(?:'([^']+)'|([^\s!]+))!(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)"

        for match in re.finditer(range_pattern, formula):
            if match.group(1):  # Sheet with quotes
                sheet = match.group(1)
                range_ref = match.group(3)
            elif match.group(2):  # Sheet without quotes
                sheet = match.group(2)
                range_ref = match.group(3)
            elif match.group(4):  # No sheet
                sheet = current_sheet
                range_ref = match.group(4)
            else:
                continue

            # Expand range to individual cells
            range_refs = self._expand_range(range_ref, sheet)
            refs.update(range_refs)

        return refs

    def _expand_range(self, range_ref: str, sheet: str) -> Set[str]:
        """
        Expand a range reference to individual cell references.

        Args:
            range_ref: Range like "A1:B10"
            sheet: Sheet name

        Returns:
            Set of individual cell references
        """
        refs = set()

        # Remove $ signs
        range_ref = range_ref.replace('$', '')

        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_ref)

            # Limit expansion to avoid memory issues with very large ranges
            MAX_CELLS = 10000
            total_cells = (max_col - min_col + 1) * (max_row - min_row + 1)

            if total_cells > MAX_CELLS:
                logger.warning(f"Range {range_ref} contains {total_cells} cells, "
                             f"limiting to first {MAX_CELLS} for dependency tracking")
                # Sample the range instead of expanding all
                return {f"{sheet}!{range_ref}"}  # Keep as range reference

            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    col_letter = openpyxl.utils.get_column_letter(col)
                    refs.add(self._make_cell_ref(sheet, f"{col_letter}{row}"))

        except Exception as e:
            logger.warning(f"Failed to expand range {range_ref}: {e}")

        return refs

    def _detect_circular_references(self) -> None:
        """Detect circular references in the dependency graph using DFS."""
        visited = set()
        rec_stack = set()

        def dfs(cell: str, path: List[str]) -> bool:
            """Depth-first search to detect cycles."""
            visited.add(cell)
            rec_stack.add(cell)
            path.append(cell)

            for dependent in self.precedents.get(cell, []):
                if dependent not in visited:
                    if dfs(dependent, path):
                        return True
                elif dependent in rec_stack:
                    # Found circular reference
                    cycle_start = path.index(dependent)
                    cycle = path[cycle_start:] + [dependent]
                    self.circular_refs.append(cycle)
                    logger.warning(f"Circular reference detected: {' -> '.join(cycle)}")
                    return True

            path.pop()
            rec_stack.remove(cell)
            return False

        for cell in self.formula_cells:
            if cell not in visited:
                dfs(cell, [])

        if self.circular_refs:
            logger.warning(f"Found {len(self.circular_refs)} circular reference(s)")

    def classify_cell(self, sheet_name: str, coordinate: str) -> str:
        """
        Classify a cell as Input, Calculation, or Output.

        Classification rules:
        - Input: No formula
        - Output: Has formula but no dependents
        - Calculation: Has formula and has dependents

        Args:
            sheet_name: Sheet name
            coordinate: Cell coordinate (e.g., "A1")

        Returns:
            Cell type: "Input", "Calculation", or "Output"
        """
        cell_ref = self._make_cell_ref(sheet_name, coordinate)

        # Check if cell has formula
        if cell_ref not in self.formula_cells:
            return "Input"

        # Has formula - check if it has dependents
        if self.dependents.get(cell_ref):
            return "Calculation"
        else:
            return "Output"

    def get_calculation_order(self) -> List[str]:
        """
        Get topological sort of formula cells for calculation order.

        Uses Kahn's algorithm for topological sorting.

        Returns:
            Ordered list of cell references to calculate

        Raises:
            ValueError: If circular dependencies prevent ordering
        """
        logger.info("Computing calculation order...")

        # Build in-degree map (number of dependencies)
        in_degree = defaultdict(int)
        for cell in self.formula_cells:
            in_degree[cell] = len([p for p in self.precedents[cell] if p in self.formula_cells])

        # Start with cells that have no dependencies
        queue = deque([cell for cell in self.formula_cells if in_degree[cell] == 0])
        calc_order = []

        while queue:
            cell = queue.popleft()
            calc_order.append(cell)

            # Reduce in-degree for dependents
            for dependent in self.dependents.get(cell, []):
                if dependent in self.formula_cells:
                    in_degree[dependent] -= 1
                    if in_degree[dependent] == 0:
                        queue.append(dependent)

        # Check if all formula cells were processed
        if len(calc_order) != len(self.formula_cells):
            remaining = len(self.formula_cells) - len(calc_order)
            logger.error(f"Circular dependency prevents complete ordering. "
                        f"{remaining} cells not ordered.")
            # Return partial order
            logger.warning("Returning partial calculation order (circular refs excluded)")

        logger.info(f"Calculation order computed: {len(calc_order)} cells")
        return calc_order

    def get_dependents(self, sheet_name: str, coordinate: str) -> Set[str]:
        """Get all cells that depend on the given cell."""
        cell_ref = self._make_cell_ref(sheet_name, coordinate)
        return self.dependents.get(cell_ref, set())

    def get_precedents(self, sheet_name: str, coordinate: str) -> Set[str]:
        """Get all cells that the given cell depends on."""
        cell_ref = self._make_cell_ref(sheet_name, coordinate)
        return self.precedents.get(cell_ref, set())

    def has_circular_refs(self) -> bool:
        """Check if workbook has circular references."""
        return len(self.circular_refs) > 0

    def get_stats(self) -> Dict[str, int]:
        """Get dependency graph statistics."""
        return {
            'total_formula_cells': len(self.formula_cells),
            'cells_with_precedents': len(self.precedents),
            'cells_with_dependents': len(self.dependents),
            'circular_references': len(self.circular_refs),
        }
