"""Main parser for Layer 1: Generate mapping report from Excel workbook."""

from pathlib import Path
from excel_pipeline.core.excel_io import load_workbook
from excel_pipeline.core.dependency_graph import DependencyGraph
from excel_pipeline.core.formula_analyzer import FormulaAnalyzer
from excel_pipeline.layer1.cell_extractor import CellExtractor
from excel_pipeline.layer1.mapping_writer import MappingWriter
from excel_pipeline.utils.logging_setup import get_logger
from excel_pipeline.utils.config import config

logger = get_logger(__name__)


def generate_mapping_report(input_path: str, output_path: str) -> None:
    """
    Generate mapping report from Excel workbook.

    This is Layer 1 of the pipeline - the single source of truth.
    The mapping report contains all cell metadata and is used by all downstream layers.

    Process:
    1. Load workbook (preserving formulas)
    2. Build dependency graph (precedents/dependents)
    3. Analyze formulas for patterns (vectorization groups)
    4. Extract all cell metadata
    5. Annotate cells with group information
    6. Write mapping_report.xlsx

    Args:
        input_path: Path to original Excel file
        output_path: Path to save mapping_report.xlsx

    Example:
        >>> generate_mapping_report("model.xlsx", "mapping_report.xlsx")

    Note:
        The mapping_report.xlsx is the CONTRACT between parser and all downstream stages.
        ALL decisions about what to include/group/regenerate are encoded in that file.
    """
    logger.info("=" * 80)
    logger.info("LAYER 1: Generating Mapping Report")
    logger.info("=" * 80)
    logger.info(f"Input: {input_path}")
    logger.info(f"Output: {output_path}")

    # Validate input
    input_file = Path(input_path)
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Step 1: Load workbook (formulas preserved)
    logger.info("\n[Step 1/6] Loading workbook...")
    wb = load_workbook(input_path, data_only=False, read_only=False)
    logger.info(f"Loaded {len(wb.sheetnames)} sheets: {', '.join(wb.sheetnames)}")

    # Step 2: Build dependency graph
    logger.info("\n[Step 2/6] Building dependency graph...")
    dep_graph = DependencyGraph(wb)
    dep_graph.build()

    dep_stats = dep_graph.get_stats()
    logger.info(f"Dependency graph stats: {dep_stats}")

    if dep_graph.has_circular_refs():
        logger.warning(f"WARNING: {dep_stats['circular_references']} circular reference(s) detected!")
        for cycle in dep_graph.circular_refs:
            logger.warning(f"  Circular: {' -> '.join(cycle[:5])}...")

    # Step 3: Analyze formulas for patterns (CRITICAL for vectorization)
    logger.info("\n[Step 3/6] Analyzing formulas for vectorization patterns...")
    analyzer = FormulaAnalyzer(vectorization_threshold=config.vectorization_threshold)

    all_groups = []
    for sheet in wb.worksheets:
        groups = analyzer.analyze_sheet(sheet)
        all_groups.extend(groups)

    analyzer_stats = analyzer.get_stats()
    logger.info(f"Formula analysis stats: {analyzer_stats}")
    logger.info(f"DRAGGED FORMULAS: {analyzer_stats['total_cells_in_groups']} cells "
               f"in {analyzer_stats['total_groups']} groups")
    logger.info(f"VECTORIZABLE: {analyzer_stats['vectorizable_cells']} cells "
               f"in {analyzer_stats['vectorizable_groups']} groups (will use numpy/pandas)")

    # Step 4: Extract all cell metadata
    logger.info("\n[Step 4/6] Extracting cell metadata...")
    extractor = CellExtractor(wb, dep_graph)
    cells = extractor.extract_all()

    # Step 5: Annotate cells with group information
    logger.info("\n[Step 5/6] Annotating cells with vectorization groups...")
    cells = extractor.annotate_with_groups(cells, all_groups)

    # Step 6: Write mapping report
    logger.info("\n[Step 6/6] Writing mapping report...")
    writer = MappingWriter(input_file.name)
    writer.write(cells, output_path, dep_stats, analyzer_stats)

    # Summary
    logger.info("\n" + "=" * 80)
    logger.info("LAYER 1: Complete!")
    logger.info("=" * 80)
    logger.info(f"Mapping report: {output_path}")
    logger.info(f"Total cells: {len(cells)}")

    # Count by type
    type_counts = {'Input': 0, 'Calculation': 0, 'Output': 0}
    for cell in cells:
        type_counts[cell.cell_type] += 1

    logger.info(f"  Input: {type_counts['Input']}")
    logger.info(f"  Calculation: {type_counts['Calculation']}")
    logger.info(f"  Output: {type_counts['Output']}")
    logger.info(f"Dragged formulas: {sum(1 for c in cells if c.group_id > 0)} cells "
               f"in {analyzer_stats['total_groups']} groups")
    logger.info(f"Vectorizable: {sum(1 for c in cells if c.is_vectorizable)} cells "
               f"in {analyzer_stats['vectorizable_groups']} groups")
    logger.info("=" * 80)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python -m excel_pipeline.layer1.parser <input.xlsx> <mapping_report.xlsx>")
        sys.exit(1)

    generate_mapping_report(sys.argv[1], sys.argv[2])
