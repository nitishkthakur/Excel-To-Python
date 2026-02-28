"""
Column-N sampling strategy — extracts a vertical strip from each sheet.

The strategy locates the first column that contains data (typically a
label column with finance line items such as "Revenue", "COGS", etc.)
and then loads that column plus the next *num_columns* data columns,
producing a narrow vertical slice of the sheet.
"""

from typing import Any

from excel_reader_smart_sampler import (
    open_workbook,
    open_workbook_values,
    detect_regions,
    _cell_info,
)


DEFAULT_NUM_COLUMNS = 10


def _find_label_column(ws, region) -> int:
    """Return the column index of the first non-empty column in *region*.

    This column typically holds row labels (e.g. "Revenue", "Expenses").
    """
    for c in range(region.min_col, region.max_col + 1):
        for r in range(region.min_row, region.max_row + 1):
            if ws.cell(row=r, column=c).value is not None:
                return c
    return region.min_col


def extract_sheet_data(path: str, sheet_name: str,
                       num_columns: int = DEFAULT_NUM_COLUMNS) -> dict[str, Any]:
    """
    Extract a vertical strip from a single sheet.

    For every detected data region the strip starts at the first column
    that contains data (the *label column*) and extends for
    ``num_columns`` additional columns to the right (or until the region
    boundary, whichever comes first).  All rows within the region are
    included.

    Args:
        path: Path to the .xlsx file.
        sheet_name: Name of the sheet to read.
        num_columns: How many columns to include *after* the label column.
            Default is 10.

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
        total_rows += reg.row_count()

        label_col = _find_label_column(ws_f, reg)
        # Strip spans: label_col .. label_col + num_columns  (clamped to region)
        strip_max_col = min(label_col + num_columns, reg.max_col)

        # Extract headers
        headers: list[str] = []
        if reg.header_row is not None:
            for c in range(label_col, strip_max_col + 1):
                v = ws_f.cell(row=reg.header_row, column=c).value
                headers.append(str(v) if v is not None else "")

        # Extract all rows in the vertical strip
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in range(reg.min_row, reg.max_row + 1):
            loaded_rows += 1
            if r == reg.header_row:
                continue
            cells: list[Any] = []
            for c in range(label_col, strip_max_col + 1):
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
            "max_row": reg.max_row,
            "min_col": label_col,
            "max_col": strip_max_col,
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
