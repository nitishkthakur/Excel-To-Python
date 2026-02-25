"""
Full sampling strategy — loads all contents of each sheet within the
detected data bounds.

Parameters *nrows* and *ncols* default to ``None``, which means the
entire sheet is loaded.  When set to an integer the output is capped at
that many rows / columns starting from the top-left of each detected
region.
"""

from typing import Any

from excel_reader_smart_sampler import (
    open_workbook,
    open_workbook_values,
    detect_regions,
    _cell_info,
)


def extract_sheet_data(path: str, sheet_name: str,
                       nrows: int | None = None,
                       ncols: int | None = None) -> dict[str, Any]:
    """
    Extract all data from a single sheet.

    Args:
        path: Path to the .xlsx file.
        sheet_name: Name of the sheet to read.
        nrows: Maximum number of rows to load per region.  ``None`` loads all.
        ncols: Maximum number of columns to load per region.  ``None`` loads all.

    Returns a dict compatible with the formatter functions (to_markdown,
    to_json, to_xml).
    """
    wb_f = open_workbook(path)
    wb_v = open_workbook_values(path)
    ws_f = wb_f[sheet_name]
    ws_v = wb_v[sheet_name]

    regions = detect_regions(ws_f)
    result_regions: list[dict[str, Any]] = []
    total_rows = 0
    loaded_rows = 0

    for reg in regions:
        # Determine effective row / column bounds
        eff_max_row = reg.max_row
        if nrows is not None:
            eff_max_row = min(reg.max_row, reg.min_row + nrows - 1)

        eff_max_col = reg.max_col
        if ncols is not None:
            eff_max_col = min(reg.max_col, reg.min_col + ncols - 1)

        total_rows += reg.row_count()

        # Extract headers
        headers: list[str] = []
        if reg.header_row is not None and reg.header_row <= eff_max_row:
            for c in range(reg.min_col, eff_max_col + 1):
                v = ws_f.cell(row=reg.header_row, column=c).value
                headers.append(str(v) if v is not None else "")

        # Extract rows
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in range(reg.min_row, eff_max_row + 1):
            loaded_rows += 1
            if r == reg.header_row:
                continue
            cells: list[Any] = []
            for c in range(reg.min_col, eff_max_col + 1):
                info = _cell_info(ws_f, ws_v, r, c)
                cells.append(info["value"])
                if "formula" in info:
                    formulas.append({
                        "address": info["address"],
                        "formula": info["formula"],
                        "cached_value": info["value"],
                    })
            row_data.append({"row_number": r, "values": cells})

        result_regions.append({
            "region": str(reg),
            "min_row": reg.min_row,
            "max_row": eff_max_row,
            "min_col": reg.min_col,
            "max_col": eff_max_col,
            "headers": headers,
            "rows": row_data,
            "formulas": formulas,
        })

    wb_f.close()
    wb_v.close()

    return {
        "sheet_name": sheet_name,
        "regions": result_regions,
        "sampled": False,
        "total_rows": total_rows,
        "sampled_rows": loaded_rows,
    }
