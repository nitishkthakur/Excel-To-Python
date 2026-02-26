"""
Core mapping logic: classify cells into Inputs / Calculations / Outputs
and generate an Excel mapping report.

Reuses the workbook parser and cell classifier from ``excel_to_python``
and the vectorised grouping logic from ``excel_to_python_vectorized``.
"""

import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from excel_to_python import (
    load_config,
    parse_workbook,
    classify_cells,
    find_all_references,
    col_letter_to_index,
    index_to_col_letter,
)

from excel_to_python_vectorized.vectorizer import (
    group_formulas,
    extract_references,
)


# ------------------------------------------------------------------
# Determine which formula cells are "outputs" (not referenced by others)
# ------------------------------------------------------------------

def _build_all_referenced_cells(formula_cells):
    """Return the set of (sheet, col, row) cells referenced by any formula."""
    referenced = set()
    for sheet, col, row, formula, _ci in formula_cells:
        refs = extract_references(formula, sheet)
        for ref in refs:
            if ref.kind == "cell":
                referenced.add((ref.sheet, ref.col, ref.row))
            elif ref.kind == "range":
                ci1 = col_letter_to_index(ref.col)
                ci2 = col_letter_to_index(ref.end_col)
                for r in range(ref.row, ref.end_row + 1):
                    for c in range(ci1, ci2 + 1):
                        referenced.add((ref.sheet, index_to_col_letter(c), r))
    return referenced


def _classify_formula_cells(formula_cells):
    """Split formula cells into *calculations* and *outputs*.

    A formula cell is an **output** if no other formula references it.
    Otherwise it is a **calculation** (intermediate result).
    """
    referenced = _build_all_referenced_cells(formula_cells)

    calculations = []
    outputs = []
    for entry in formula_cells:
        sheet, col, row, formula, cell_info = entry
        if (sheet, col, row) in referenced:
            calculations.append(entry)
        else:
            outputs.append(entry)
    return calculations, outputs


# ------------------------------------------------------------------
# Group description helpers (compact representation like vectorizer)
# ------------------------------------------------------------------

def _describe_group(group):
    """Return a human-readable summary of a vectorised group."""
    cells = group["cells"]
    direction = group["direction"]
    sheet, col0, row0, formula0, _ = cells[0]
    _, col1, row1, _, _ = cells[-1]
    count = len(cells)

    if direction == "vertical":
        rng = f"{col0}{row0}:{col0}{row1}"
    else:
        rng = f"{col0}{row0}:{col1}{row0}"

    return (
        f"{sheet}!{rng}  ({count} cells, {direction})  "
        f"Pattern: {formula0}"
    )


def _describe_single(entry):
    """Return a description for a non-grouped formula cell."""
    sheet, col, row, formula, _ = entry
    return f"{sheet}!{col}{row}  Formula: {formula}"


# ------------------------------------------------------------------
# Build per-sheet mapping data
# ------------------------------------------------------------------

def _build_sheet_mapping(sheet_name, hardcoded_cells, formula_cells):
    """Return (inputs, calc_descriptions, output_descriptions) for *sheet_name*.

    * inputs: list of (col, row, value)
    * calc_descriptions: list of str (one per group or single calc cell)
    * output_descriptions: list of str (one per group or single output cell)
    """
    # --- inputs ---
    inputs = [
        (col, row, val)
        for s, col, row, val, _ci in hardcoded_cells
        if s == sheet_name
    ]
    inputs.sort(key=lambda x: (x[1], col_letter_to_index(x[0])))

    # --- all formula cells on this sheet ---
    sheet_formulas = [
        entry for entry in formula_cells if entry[0] == sheet_name
    ]

    # classify into calculations vs outputs (using *all* formulas for
    # reference analysis so cross-sheet deps are accounted for)
    calculations_all, outputs_all = _classify_formula_cells(formula_cells)
    calc_set = {(s, c, r) for s, c, r, *_ in calculations_all}
    output_set = {(s, c, r) for s, c, r, *_ in outputs_all}

    sheet_calc_formulas = [
        e for e in sheet_formulas if (e[0], e[1], e[2]) in calc_set
    ]
    sheet_output_formulas = [
        e for e in sheet_formulas if (e[0], e[1], e[2]) in output_set
    ]

    # --- vectorise each subset independently ---
    calc_descriptions = _grouped_descriptions(sheet_calc_formulas)
    output_descriptions = _grouped_descriptions(sheet_output_formulas)

    return inputs, calc_descriptions, output_descriptions


def _grouped_descriptions(formula_subset):
    """Run the vectoriser grouping on *formula_subset* and return descriptions."""
    if not formula_subset:
        return []
    groups, singles = group_formulas(formula_subset)

    descriptions = []

    # Sort groups by the first cell position for deterministic output
    groups.sort(key=lambda g: (
        g["cells"][0][0],
        g["cells"][0][2],
        col_letter_to_index(g["cells"][0][1]),
    ))
    for g in groups:
        descriptions.append(_describe_group(g))

    # Sort singles similarly
    singles.sort(key=lambda s: (s[0], s[2], col_letter_to_index(s[1])))
    for s in singles:
        descriptions.append(_describe_single(s))

    return descriptions


# ------------------------------------------------------------------
# Report generation
# ------------------------------------------------------------------

_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4",
                           fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_SECTION_FONT = Font(bold=True, size=12)


def _write_section_header(ws, row, title):
    """Write a section title (Inputs / Calculations / Outputs)."""
    ws.cell(row=row, column=1, value=title).font = _SECTION_FONT
    return row + 1


def _write_table_header(ws, row, headers):
    """Write a bold, coloured header row."""
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=ci, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
    return row + 1


def generate_mapping_report(excel_path, sheet_names=None,
                            config_path=None, output_path=None):
    """Generate the mapping report Excel file.

    Parameters
    ----------
    excel_path : str
        Path to the source Excel workbook.
    sheet_names : list[str] or None
        Sheets to include.  ``None`` means *all* sheets.
    config_path : str or None
        Optional path to config YAML.
    output_path : str or None
        Where to write the report.  Defaults to
        ``<excel_dir>/output/mapping_report.xlsx``.

    Returns
    -------
    str
        Path to the generated report file.
    """
    config = load_config(config_path)

    wb_src = load_workbook(excel_path)
    sheets, tables = parse_workbook(wb_src)
    formula_cells, hardcoded_cells = classify_cells(sheets, tables)

    if sheet_names is None:
        sheet_names = list(sheets.keys())
    else:
        # Keep only sheets that actually exist
        sheet_names = [s for s in sheet_names if s in sheets]

    if output_path is None:
        out_dir = os.path.join(os.path.dirname(excel_path) or ".", "output")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, "mapping_report.xlsx")
    else:
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    wb_report = Workbook()
    # Remove default sheet only after we've added at least one
    has_default = "Sheet" in wb_report.sheetnames

    for sn in sheet_names:
        inputs, calcs, outputs = _build_sheet_mapping(
            sn, hardcoded_cells, formula_cells)

        ws = wb_report.create_sheet(sn)
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 60

        row = 1

        # ---- Inputs section ----
        row = _write_section_header(ws, row, "Inputs")
        row = _write_table_header(ws, row, ["Cell", "Value", ""])
        for col, r, val in inputs:
            ws.cell(row=row, column=1, value=f"{col}{r}")
            ws.cell(row=row, column=2, value=val)
            row += 1
        if not inputs:
            ws.cell(row=row, column=1, value="(none)")
            row += 1
        row += 1  # blank separator

        # ---- Calculations section ----
        row = _write_section_header(ws, row, "Calculations")
        row = _write_table_header(ws, row, ["#", "Description", ""])
        for idx, desc in enumerate(calcs, 1):
            ws.cell(row=row, column=1, value=idx)
            ws.cell(row=row, column=2, value=desc)
            row += 1
        if not calcs:
            ws.cell(row=row, column=1, value="(none)")
            row += 1
        row += 1

        # ---- Outputs section ----
        row = _write_section_header(ws, row, "Outputs")
        row = _write_table_header(ws, row, ["#", "Description", ""])
        for idx, desc in enumerate(outputs, 1):
            ws.cell(row=row, column=1, value=idx)
            ws.cell(row=row, column=2, value=desc)
            row += 1
        if not outputs:
            ws.cell(row=row, column=1, value="(none)")
            row += 1

    # Remove the default "Sheet" if we added real sheets
    if has_default and "Sheet" in wb_report.sheetnames and len(wb_report.sheetnames) > 1:
        del wb_report["Sheet"]

    wb_report.save(output_path)
    wb_src.close()

    print(f"Generated mapping report: {output_path}")
    return output_path
