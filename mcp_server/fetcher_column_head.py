"""
Column-Head Fetcher — vectorized sampling that captures all headers and
the first *n* data columns of each detected patch.

This mode is the **column-wise** counterpart of ``row_head``.  It is
designed for financial Excel sheets where:

* **Column headers** are dates (e.g. Q1 2023, Q2 2023, … spanning
  several years).
* **Row indices** are entities such as *Income*, *Profit*, *Revenue*,
  *COGS*, etc.

By reading the first *n* columns of each patch the LLM can see the
entity labels in the row index **and** the first few date-period
columns, making it easy to relate entities to their time-series values.

The budget (``max_cols``) is **per sheet** and is divided across patches
proportionally to each patch's total column count.

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


# ---------------------------------------------------------------------------
# Column selection — vectorized head(n) per patch
# ---------------------------------------------------------------------------

def _allocate_col_budget(regions: list[dict], total_budget: int) -> list[int]:
    """Divide *total_budget* columns across *regions* proportionally to width.

    Each region is guaranteed at least 2 columns (label + 1 data col)
    when the budget allows.  Returns a list of per-region budgets.
    """
    if not regions:
        return []

    widths = np.array([r["max_col"] - r["min_col"] + 1 for r in regions],
                      dtype=float)
    total_width = widths.sum()

    if total_width == 0:
        return [0] * len(regions)

    raw = (widths / total_width) * total_budget

    # Only enforce minimum of 2 per region if budget allows
    min_per_region = 2 if total_budget >= 2 * len(regions) else 1
    budgets = np.maximum(np.floor(raw).astype(int), min_per_region)

    # If sum exceeds total_budget, scale back proportionally
    current_sum = int(budgets.sum())
    if current_sum > total_budget:
        budgets = np.maximum(np.floor(raw).astype(int), 1)
        current_sum = int(budgets.sum())
        while current_sum > total_budget:
            order = np.argsort(widths)
            for idx in order:
                if current_sum <= total_budget:
                    break
                if budgets[idx] > 1:
                    budgets[idx] -= 1
                    current_sum -= 1

    remaining = total_budget - int(budgets.sum())
    if remaining > 0:
        order = np.argsort(-widths)
        for idx in order:
            if remaining <= 0:
                break
            budgets[idx] += 1
            remaining -= 1

    for i, reg in enumerate(regions):
        reg_cols = reg["max_col"] - reg["min_col"] + 1
        budgets[i] = min(int(budgets[i]), reg_cols)

    return budgets.tolist()


def column_head_indices(region: dict, budget: int) -> list[int]:
    """Return 1-based column indices: first *budget* columns of the region."""
    cmin = region["min_col"]
    cmax = region["max_col"]
    end = min(cmin + budget, cmax + 1)
    return list(range(cmin, end))


# ---------------------------------------------------------------------------
# High-level extraction
# ---------------------------------------------------------------------------

def extract_sheet_data(path: str, sheet_name: str,
                       max_cols: int = 20) -> dict[str, Any]:
    """Extract column-head sampled data from a single sheet (vectorized).

    Identifies all patches, allocates *max_cols* budget across patches
    proportionally, then reads **all rows** but only the first N
    columns of each patch.  This preserves the row-entity labels and the
    first few date/period columns.

    Parameters
    ----------
    path : str
        Path to the ``.xlsx`` file.
    sheet_name : str
        Name of the sheet to read.
    max_cols : int
        Total column budget **per sheet**.  Divided across patches by
        width.

    Returns
    -------
    dict
        Compatible with the formatter functions (to_markdown, to_json,
        to_xml).
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
    budgets = _allocate_col_budget(regions, max_cols)

    result_regions: list[dict[str, Any]] = []
    total_rows = 0
    sampled_rows = 0
    was_sampled = False

    for reg, col_budget in zip(regions, budgets):
        rmin, rmax = reg["min_row"], reg["max_row"]
        cmin, cmax = reg["min_col"], reg["max_col"]
        reg_total = rmax - rmin + 1
        total_rows += reg_total

        cols_to_read = column_head_indices(reg, col_budget)
        all_rows = list(range(rmin, rmax + 1))
        sampled_rows += len(all_rows)

        if len(cols_to_read) < (cmax - cmin + 1):
            was_sampled = True

        # Headers (vectorized slice)
        headers: list[str] = []
        if reg["header_row"] is not None:
            hdr_vals = df_f.loc[reg["header_row"], cols_to_read]
            headers = [str(v) if pd.notna(v) else "" for v in hdr_vals]

        # Row data + formulas
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in all_rows:
            if r == reg["header_row"]:
                continue
            val_row = df_v.loc[r, cols_to_read] if r in df_v.index else pd.Series(
                [None] * len(cols_to_read), index=cols_to_read)
            frm_row = df_f.loc[r, cols_to_read]

            cells: list[Any] = []
            for c in cols_to_read:
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

        eff_max_col = cols_to_read[-1] if cols_to_read else cmin
        result_regions.append({
            "region": (f"DataRegion(rows={rmin}-{rmax}, "
                       f"cols={cmin}-{eff_max_col}, header={reg['header_row']})"),
            "min_row": rmin,
            "max_row": rmax,
            "min_col": cmin,
            "max_col": eff_max_col,
            "headers": headers,
            "rows": row_data,
            "formulas": formulas,
        })

    return {
        "sheet_name": sheet_name,
        "regions": result_regions,
        "sampled": was_sampled,
        "total_rows": total_rows,
        "sampled_rows": sampled_rows,
    }
