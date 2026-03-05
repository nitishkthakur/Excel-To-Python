"""Main entry point for Layer 4a: Python Code Generator."""

from typing import Dict, List, Union, Tuple
from excel_pipeline.layer4a.mapping_reader import MappingReader, CellMetadata, GroupMetadata
from excel_pipeline.layer4a.formula_translator import FormulaTranslator
from excel_pipeline.layer4a.dependency_graph import DependencyGraph
from excel_pipeline.layer4a.vectorization_engine import VectorizationEngine
from excel_pipeline.layer4a.code_emitter import CodeEmitter
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


def generate_unstructured_code(
    mapping_report_path: str,
    unstructured_inputs_path: str,
    output_script_path: str = "unstructured_calculate.py"
) -> None:
    """
    Generate Python code for unstructured calculation.

    This is the main entry point for Layer 4a.

    Args:
        mapping_report_path: Path to mapping_report.xlsx
        unstructured_inputs_path: Path to unstructured_inputs.xlsx (for reference)
        output_script_path: Path to write generated Python script

    Process:
        1. Read mapping report and identify vectorizable groups
        2. Build dependency graph for calculation order
        3. Translate formulas to Python expressions
        4. Generate vectorized code for eligible groups
        5. Assemble final Python script
    """
    logger.info("=" * 80)
    logger.info("LAYER 4a: Python Code Generator (Unstructured Path)")
    logger.info("=" * 80)
    logger.info(f"Mapping report: {mapping_report_path}")
    logger.info(f"Input reference: {unstructured_inputs_path}")
    logger.info(f"Output script: {output_script_path}")

    # Step 1: Read mapping report
    logger.info("\n[Step 1/5] Reading mapping report...")
    reader = MappingReader(mapping_report_path)
    cells_by_sheet = reader.read_mapping_report()
    vectorizable_groups = reader.identify_vectorizable_groups()

    # Separate vectorizable groups from individual cells
    # Don't expand vectorizable groups - keep them consolidated for pandas operations
    vectorizable_group_ids = {g.group_id for g in vectorizable_groups}

    # Collect cells that are NOT part of vectorizable groups
    non_vectorized_cells = []
    for sheet_name, cells in cells_by_sheet.items():
        for cell in cells:
            # Skip cells that belong to vectorizable groups
            if cell.group_id and cell.group_id in vectorizable_group_ids:
                continue

            # Expand non-vectorizable cells/groups
            if ':' in cell.cell:
                # This is a range - expand it
                expanded = reader._expand_range(cell)
                non_vectorized_cells.extend(expanded)
            else:
                # Individual cell
                non_vectorized_cells.append(cell)

    logger.info(f"  Vectorizable groups: {len(vectorizable_groups)}")
    logger.info(f"  Individual cells (non-vectorized): {len(non_vectorized_cells)}")

    # Step 2: Build dependency graph
    logger.info("\n[Step 2/5] Building dependency graph...")
    dep_graph = DependencyGraph()
    translator = FormulaTranslator()

    # Add vectorizable groups to graph as single nodes
    for group in vectorizable_groups:
        # Skip groups with no cells
        if not group.cells:
            logger.warning(f"Group {group.group_id} has no cells, skipping")
            continue

        # Use first cell as representative for dependencies
        dep_graph.add_cell(group.sheet, group.cells[0], group.group_id)

        if group.pattern_formula:
            # Extract dependencies from pattern formula
            dependencies = translator.extract_dependencies(group.pattern_formula, group.sheet)
            for dep_sheet, dep_col, dep_row in dependencies:
                # Add dependency from group to the precedent cell
                dep_graph.add_dependency(
                    group.sheet, group.cells[0],
                    dep_sheet, f"{dep_col}{dep_row}"
                )

    # Add individual cells to graph
    for cell in non_vectorized_cells:
        dep_graph.add_cell(cell.sheet, cell.cell, cell.group_id)

        # Add dependencies if cell has formula
        if cell.formula:
            dependencies = translator.extract_dependencies(cell.formula, cell.sheet)
            for dep_sheet, dep_col, dep_row in dependencies:
                dep_graph.add_dependency(
                    cell.sheet, cell.cell,
                    dep_sheet, f"{dep_col}{dep_row}"
                )

    # Get calculation order
    calc_order = dep_graph.topological_sort()

    stats = dep_graph.get_statistics()
    logger.info(f"  Calculation order: {len(calc_order)} nodes")
    logger.info(f"  Max dependency level: {stats['max_calculation_level']}")

    # Step 3: Initialize code generation components
    logger.info("\n[Step 3/5] Initializing code generators...")
    vec_engine = VectorizationEngine(translator)
    emitter = CodeEmitter()

    # Add standard sections
    emitter.add_imports()
    emitter.add_helper_functions()
    emitter.add_input_loading()

    # Step 4: Generate calculation code
    logger.info("\n[Step 4/5] Generating calculation code...")

    processed_groups = set()
    individual_cells = []

    for node in calc_order:
        if isinstance(node, int):
            # This is a group
            group_id = node
            if group_id in processed_groups:
                continue

            processed_groups.add(group_id)

            # Find the group metadata
            group = reader.groups.get(group_id)
            if not group:
                logger.warning(f"Group {group_id} not found in metadata")
                continue

            # Generate vectorized code for this group
            logger.debug(f"  Generating code for group {group_id} ({len(group.cells)} cells)")
            code_lines = vec_engine.generate_vectorized_code(group)
            emitter.add_calculation_section(code_lines)

        else:
            # Individual cell
            sheet, cell_coord = node
            individual_cells.append((sheet, cell_coord))

    # Generate code for individual cells
    logger.info(f"  Generated code for {len(processed_groups)} groups and {len(individual_cells)} individual cells")

    for sheet, cell_coord in individual_cells:
        # Find cell metadata
        cell_meta = next(
            (c for c in non_vectorized_cells if c.sheet == sheet and c.cell == cell_coord),
            None
        )

        if not cell_meta or not cell_meta.formula:
            continue

        # Generate individual calculation
        python_expr = translator.translate_formula(
            cell_meta.formula,
            sheet,
            is_vectorized=False
        )

        # Extract column and row from coordinate
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell_coord)
        if match:
            col, row = match.groups()
            code_line = f"c[('{sheet}', '{col}', {row})] = {python_expr}"
            emitter.add_calculation_section([code_line, ""])

    # Step 5: Assemble and write final script
    logger.info("\n[Step 5/5] Assembling final script...")

    # Reconstruct complete cell list for output writing
    # Include both vectorized groups (expanded) and individual cells
    all_output_cells = []
    for group in vectorizable_groups:
        # Expand vectorizable groups for output writing only
        first_cell_meta = next(
            (c for sheet_cells in cells_by_sheet.values() for c in sheet_cells
             if c.group_id == group.group_id),
            None
        )
        if first_cell_meta and ':' in first_cell_meta.cell:
            all_output_cells.extend(reader._expand_range(first_cell_meta))
        else:
            # Group cells are already individual
            all_output_cells.extend([
                next((c for sheet_cells in cells_by_sheet.values() for c in sheet_cells
                      if c.sheet == group.sheet and c.cell == cell_coord), None)
                for cell_coord in group.cells
            ])
    all_output_cells.extend(non_vectorized_cells)

    emitter.add_output_writing(all_output_cells)
    emitter.write_to_file(output_script_path)

    code_stats = emitter.get_code_stats()
    logger.info(f"  Total lines: {code_stats['total_lines']}")
    logger.info(f"  Calculation lines: {code_stats['calculation_lines']}")

    # Summary
    logger.info("\n" + "=" * 80)
    logger.info("LAYER 4a: Complete!")
    logger.info("=" * 80)
    logger.info(f"Generated: {output_script_path}")
    logger.info(f"  Lines of code: {code_stats['total_lines']}")
    logger.info(f"  Vectorized groups: {len(vectorizable_groups)}")
    logger.info(f"  Individual cells: {len(individual_cells)}")
    logger.info("=" * 80)


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        print("Usage: python -m excel_pipeline.layer4a.code_generator <mapping_report.xlsx> <unstructured_inputs.xlsx> [output_script.py]")
        sys.exit(1)

    mapping_path = sys.argv[1]
    inputs_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else "unstructured_calculate.py"

    generate_unstructured_code(mapping_path, inputs_path, output_path)
