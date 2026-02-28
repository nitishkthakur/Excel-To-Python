"""
Full sampling strategy — loads all contents of each sheet within the
detected data bounds.

Parameters *nrows* and *ncols* default to ``None``, which means the
entire sheet is loaded.  When set to an integer the output is capped at
that many rows / columns starting from the top-left of each detected
region.

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


def extract_sheet_data(path: str, sheet_name: str,
                       nrows: int | None = None,
                       ncols: int | None = None) -> dict[str, Any]:
    """
    Extract all data from a single sheet (vectorized).

    Args:
        path: Path to the .xlsx file.
        sheet_name: Name of the sheet to read.
        nrows: Maximum number of rows to load per region.  ``None`` loads all.
        ncols: Maximum number of columns to load per region.  ``None`` loads all.

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

        # Determine effective row / column bounds
        eff_max_row = rmax
        if nrows is not None:
            eff_max_row = min(rmax, rmin + nrows - 1)

        eff_max_col = cmax
        if ncols is not None:
            eff_max_col = min(cmax, cmin + ncols - 1)

        total_rows += rmax - rmin + 1

        cols = list(range(cmin, eff_max_col + 1))
        _empty_row = pd.Series([None] * len(cols), index=cols)

        # Headers (vectorized slice)
        headers: list[str] = []
        if reg["header_row"] is not None and reg["header_row"] <= eff_max_row:
            hdr_vals = df_f.loc[reg["header_row"], cols]
            headers = [str(v) if pd.notna(v) else "" for v in hdr_vals]

        # Row data + formulas
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in range(rmin, eff_max_row + 1):
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
                       f"cols={cmin}-{cmax}, header={reg['header_row']})"),
            "min_row": rmin,
            "max_row": eff_max_row,
            "min_col": cmin,
            "max_col": eff_max_col,
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
