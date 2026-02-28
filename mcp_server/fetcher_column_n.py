"""
Column-N sampling strategy — extracts a vertical strip from each sheet.

The strategy locates the first column that contains data (typically a
label column with finance line items such as "Revenue", "COGS", etc.)
and then loads that column plus the next *num_columns* data columns,
producing a narrow vertical slice of the sheet.

The implementation is fully vectorized using **pandas** and **numpy**
for speed.
"""

from typing import Any

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

from fetcher_smart_random import (
    load_sheet_frames,
    detect_regions,
)


DEFAULT_NUM_COLUMNS = 10


def extract_sheet_data(path: str, sheet_name: str,
                       num_columns: int = DEFAULT_NUM_COLUMNS) -> dict[str, Any]:
    """
    Extract a vertical strip from a single sheet (vectorized).

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
    df_v, df_f = load_sheet_frames(path, sheet_name)
    if df_f.empty:
        return {
            "sheet_name": sheet_name,
            "regions": [],
            "sampled": False,
            "total_rows": 0,
            "sampled_rows": 0,
        }

    regions = detect_regions(df_f)
    result_regions: list[dict[str, Any]] = []
    total_rows = 0
    loaded_rows = 0

    for reg in regions:
        rmin, rmax = reg["min_row"], reg["max_row"]
        cmin, cmax = reg["min_col"], reg["max_col"]
        total_rows += rmax - rmin + 1

        # Label column is the first column with data (= min_col, detected by detect_regions)
        label_col = cmin
        # Strip spans: label_col .. label_col + num_columns  (clamped to region)
        strip_max_col = min(label_col + num_columns, cmax)
        cols = list(range(label_col, strip_max_col + 1))
        _empty_row = pd.Series([None] * len(cols), index=cols)

        # Headers (vectorized slice)
        headers: list[str] = []
        if reg["header_row"] is not None:
            hdr_vals = df_f.loc[reg["header_row"], cols]
            headers = [str(v) if pd.notna(v) else "" for v in hdr_vals]

        # Extract all rows in the vertical strip
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in range(rmin, rmax + 1):
            loaded_rows += 1
            if r == reg["header_row"]:
                continue
            val_row = df_v.loc[r, cols] if r in df_v.index else _empty_row
            frm_row = df_f.loc[r, cols]

            cells: list[Any] = []
            for c in cols:
                v_val = val_row[c] if c in val_row.index else None
                f_val = frm_row[c] if c in frm_row.index else None

                if isinstance(v_val, (np.integer,)):
                    v_val = int(v_val)
                elif isinstance(v_val, (np.floating,)):
                    v_val = float(v_val)
                if not isinstance(v_val, str) and pd.isna(v_val):
                    v_val = None

                cells.append(v_val)

                if isinstance(f_val, str) and f_val.startswith("="):
                    formulas.append({
                        "address": f"{get_column_letter(c)}{r}",
                        "formula": f_val,
                        "cached_value": v_val,
                    })

            row_data.append({"row_number": r, "values": cells})

        result_regions.append({
            "region": (f"DataRegion(rows={rmin}-{rmax}, "
                       f"cols={cmin}-{strip_max_col}, header={reg['header_row']})"),
            "min_row": rmin,
            "max_row": rmax,
            "min_col": label_col,
            "max_col": strip_max_col,
            "headers": headers,
            "rows": row_data,
            "formulas": formulas,
        })

    return {
        "sheet_name": sheet_name,
        "regions": result_regions,
        "sampled": False,
        "total_rows": total_rows,
        "sampled_rows": loaded_rows,
    }
