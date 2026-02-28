"""
Row-Head Fetcher — vectorized sampling that captures all headers and
the first *n* data rows of each detected patch.

This mode is designed for **structured overview** of an Excel sheet:
it identifies every contiguous data patch (region), always captures
the header row, and reads the first *n* rows beneath it.  The budget
(``max_rows``) is **per sheet** and is divided across patches
proportionally to each patch's total row count.

The implementation is fully vectorized using **pandas** and **numpy**
for speed.

Typical use case
----------------
When the user mentions a particular sheet is important but you still
want all other sheets scanned efficiently, use ``row_head`` on the
non-priority sheets.  This guarantees every header is captured along
with a few representative data rows so the LLM can relate values to
their column headings.
"""

from typing import Any

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

from fetcher_smart_random import (
    load_sheet_frames,
    detect_regions,
    _non_null_mask,
)


# ---------------------------------------------------------------------------
# Row selection — vectorized head(n) per patch
# ---------------------------------------------------------------------------

def _allocate_budget(regions: list[dict], total_budget: int) -> list[int]:
    """Divide *total_budget* rows across *regions* proportionally to size.

    Each region is guaranteed at least 2 rows (header + 1 data row) when
    the budget allows.  Returns a list of per-region budgets.
    """
    if not regions:
        return []

    sizes = np.array([r["max_row"] - r["min_row"] + 1 for r in regions],
                     dtype=float)
    total_size = sizes.sum()

    if total_size == 0:
        return [0] * len(regions)

    # Proportional allocation (vectorized)
    raw = (sizes / total_size) * total_budget

    # Only enforce minimum of 2 per region if budget allows
    min_per_region = 2 if total_budget >= 2 * len(regions) else 1
    budgets = np.maximum(np.floor(raw).astype(int), min_per_region)

    # If sum exceeds total_budget, scale back proportionally
    current_sum = int(budgets.sum())
    if current_sum > total_budget:
        budgets = np.maximum(np.floor(raw).astype(int), 1)
        current_sum = int(budgets.sum())
        # Trim further if still over budget
        while current_sum > total_budget:
            order = np.argsort(sizes)  # trim smallest first
            for idx in order:
                if current_sum <= total_budget:
                    break
                if budgets[idx] > 1:
                    budgets[idx] -= 1
                    current_sum -= 1

    # Distribute any remaining budget to the largest regions
    remaining = total_budget - int(budgets.sum())
    if remaining > 0:
        order = np.argsort(-sizes)
        for idx in order:
            if remaining <= 0:
                break
            budgets[idx] += 1
            remaining -= 1

    # Cap each budget to the actual region size
    for i, reg in enumerate(regions):
        reg_rows = reg["max_row"] - reg["min_row"] + 1
        budgets[i] = min(int(budgets[i]), reg_rows)

    return budgets.tolist()


def row_head_indices(region: dict, budget: int) -> list[int]:
    """Return 1-based row indices: header row + the first *budget-1* data rows."""
    min_row = region["min_row"]
    max_row = region["max_row"]
    header = region["header_row"]

    selected: list[int] = []

    # Always include header
    if header is not None:
        selected.append(header)
        data_start = header + 1
    else:
        data_start = min_row

    # Remaining budget after header
    remaining = budget - len(selected)
    end = min(data_start + remaining, max_row + 1)
    selected.extend(range(data_start, end))

    return sorted(set(selected))


# ---------------------------------------------------------------------------
# High-level extraction
# ---------------------------------------------------------------------------

def extract_sheet_data(path: str, sheet_name: str,
                       max_rows: int = 100) -> dict[str, Any]:
    """Extract row-head sampled data from a single sheet (vectorized).

    Identifies all patches, allocates *max_rows* budget across patches
    proportionally, then reads the header + first N data rows of each
    patch.

    Parameters
    ----------
    path : str
        Path to the ``.xlsx`` file.
    sheet_name : str
        Name of the sheet to read.
    max_rows : int
        Total row budget **per sheet**.  Divided across patches by size.

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
    budgets = _allocate_budget(regions, max_rows)

    result_regions: list[dict[str, Any]] = []
    total_rows = 0
    sampled_rows = 0
    was_sampled = False

    for reg, budget in zip(regions, budgets):
        rmin, rmax = reg["min_row"], reg["max_row"]
        cmin, cmax = reg["min_col"], reg["max_col"]
        reg_total = rmax - rmin + 1
        total_rows += reg_total

        rows_to_read = row_head_indices(reg, budget)
        if len(rows_to_read) < reg_total:
            was_sampled = True
        sampled_rows += len(rows_to_read)

        # Headers (vectorized slice)
        headers: list[str] = []
        cols = list(range(cmin, cmax + 1))
        if reg["header_row"] is not None:
            hdr_vals = df_f.loc[reg["header_row"], cols]
            headers = [str(v) if pd.notna(v) else "" for v in hdr_vals]

        # Row data + formulas (vectorized reads)
        row_data: list[dict[str, Any]] = []
        formulas: list[dict[str, str]] = []

        for r in rows_to_read:
            if r == reg["header_row"]:
                continue
            val_row = df_v.loc[r, cols] if r in df_v.index else pd.Series(
                [None] * len(cols), index=cols)
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
            "max_row": rmax,
            "min_col": cmin,
            "max_col": cmax,
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
