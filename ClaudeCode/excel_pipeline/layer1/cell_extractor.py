"""Extract all cell data and metadata from Excel workbook."""

from typing import List, Dict, Any
from dataclasses import dataclass, asdict
from openpyxl import Workbook
from excel_pipeline.core.excel_io import get_cell_info
from excel_pipeline.core.cell_classifier import classify_cell
from excel_pipeline.core.dependency_graph import DependencyGraph
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


@dataclass
class CellMetadata:
    """
    Comprehensive metadata for a single cell.

    This structure captures everything needed to reconstruct the cell
    and its position in the original workbook.
    """
    # Position
    sheet_name: str
    row_num: int
    col_num: int
    col_letter: str
    cell_coordinate: str

    # Classification
    cell_type: str  # "Input", "Calculation", or "Output"

    # Content
    formula: str  # Raw formula string or empty
    value: Any  # Calculated value

    # Formatting
    number_format: str
    font_bold: bool
    font_italic: bool
    font_size: int
    font_color: str
    fill_color: str
    alignment: str
    wrap_text: bool

    # Grouping (for dragged formulas and vectorization)
    group_id: int  # 0 if not in a group
    group_direction: str  # "horizontal", "vertical", or empty
    group_size: int  # 0 if not in a group
    pattern_formula: str  # Template formula or empty
    is_vectorizable: bool  # True if group will be vectorized (size >= threshold)

    # Control
    include_flag: bool  # User can set to False to exclude from processing

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)


class CellExtractor:
    """Extract all cell data from workbook with metadata."""

    def __init__(self, workbook: Workbook, dep_graph: DependencyGraph):
        """
        Initialize cell extractor.

        Args:
            workbook: Excel workbook to extract from
            dep_graph: Dependency graph for classification
        """
        self.workbook = workbook
        self.dep_graph = dep_graph

    def extract_all(self) -> List[CellMetadata]:
        """
        Extract all cells from workbook with complete metadata.

        Returns:
            List of CellMetadata objects

        Example:
            >>> extractor = CellExtractor(wb, dep_graph)
            >>> cells = extractor.extract_all()
            >>> print(f"Extracted {len(cells)} cells")
        """
        logger.info("Extracting cell data from workbook...")

        all_cells = []

        for sheet in self.workbook.worksheets:
            sheet_name = sheet.title
            logger.info(f"Extracting cells from sheet: {sheet_name}")

            sheet_cells = self._extract_sheet(sheet, sheet_name)
            all_cells.extend(sheet_cells)

        logger.info(f"Extracted {len(all_cells)} cells total")
        return all_cells

    def _extract_sheet(self, sheet, sheet_name: str) -> List[CellMetadata]:
        """Extract all cells from a single sheet."""
        cells = []

        for row in sheet.iter_rows():
            for cell in row:
                # Only capture cells with actual content or meaningful formatting
                # Skip empty cells with only default styling
                if not self._is_meaningful_cell(cell):
                    continue

                metadata = self._extract_cell(cell, sheet_name)
                cells.append(metadata)

        logger.debug(f"Sheet {sheet_name}: extracted {len(cells)} cells")
        return cells

    def _extract_cell(self, cell, sheet_name: str) -> CellMetadata:
        """Extract metadata from a single cell."""
        # Get basic cell info
        info = get_cell_info(cell)

        # Classify cell
        cell_type = classify_cell(cell, self.dep_graph, sheet_name)

        # Create metadata object
        metadata = CellMetadata(
            # Position
            sheet_name=sheet_name,
            row_num=info['row'],
            col_num=info['column'],
            col_letter=info['column_letter'],
            cell_coordinate=info['coordinate'],

            # Classification
            cell_type=cell_type,

            # Content
            formula=info['formula'] or "",
            value=info['value'],

            # Formatting
            number_format=info['number_format'],
            font_bold=info['font_bold'] or False,
            font_italic=info['font_italic'] or False,
            font_size=info['font_size'] or 11,
            font_color=info['font_color'] or "",
            fill_color=info['fill_color'] or "",
            alignment=info['alignment_horizontal'] or "general",
            wrap_text=info['wrap_text'] or False,

            # Grouping (will be filled by pattern detector)
            group_id=0,
            group_direction="",
            group_size=0,
            pattern_formula="",
            is_vectorizable=False,

            # Control
            include_flag=True  # Default to include all cells
        )

        return metadata

    def _is_meaningful_cell(self, cell) -> bool:
        """
        Check if cell has meaningful content or formatting.

        A cell is meaningful if it has:
        1. A value (number, text, formula, etc.)
        2. OR significant non-default formatting (bold, italic, non-default colors)

        Cells with only default Excel styling are considered empty.
        """
        # Has content
        if cell.value is not None:
            return True

        # Check for meaningful formatting (not just default styles)
        if hasattr(cell, 'font') and cell.font:
            # Bold or italic text
            if cell.font.bold or cell.font.italic:
                return True
            # Non-default font size (Excel default is 11)
            if cell.font.size and cell.font.size != 11 and cell.font.size != 10:
                return True
            # Explicit color (not theme-based default)
            if cell.font.color and hasattr(cell.font.color, 'rgb') and cell.font.color.rgb:
                return True

        # Check for meaningful fill color (not default/theme)
        if hasattr(cell, 'fill') and cell.fill:
            if cell.fill.patternType == 'solid':
                # Has explicit solid fill color
                if cell.fill.fgColor and hasattr(cell.fill.fgColor, 'rgb') and cell.fill.fgColor.rgb:
                    # Not black (default)
                    if str(cell.fill.fgColor.rgb) != '00000000':
                        return True

        # Default: cell is not meaningful
        return False

    def annotate_with_groups(self, cells: List[CellMetadata],
                            groups: List) -> List[CellMetadata]:
        """
        Annotate cells with group information for vectorization.

        Args:
            cells: List of CellMetadata objects
            groups: List of FormulaGroup objects from FormulaAnalyzer

        Returns:
            Updated list of CellMetadata with group information

        This is CRITICAL for performance - marks which cells can be vectorized.
        """
        logger.info(f"Annotating {len(cells)} cells with {len(groups)} group(s)")

        # Build lookup: (sheet_name, coordinate) -> CellMetadata
        cell_lookup = {}
        for cell in cells:
            key = (cell.sheet_name, cell.cell_coordinate)
            cell_lookup[key] = cell

        # Annotate cells with group information
        for group in groups:
            for coord in group.cells:
                key = (group.sheet_name, coord)
                if key in cell_lookup:
                    cell = cell_lookup[key]
                    cell.group_id = group.group_id
                    cell.group_direction = group.direction
                    cell.group_size = group.size
                    cell.pattern_formula = group.pattern
                    cell.is_vectorizable = group.is_vectorizable

        grouped_count = sum(1 for c in cells if c.group_id > 0)
        vectorizable_count = sum(1 for c in cells if c.is_vectorizable)
        logger.info(f"Marked {grouped_count} cells in dragged formula groups "
                   f"({vectorizable_count} vectorizable)")

        return cells
