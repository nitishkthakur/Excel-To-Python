"""
MCP Server — exposes Excel workbook analysis tools to LLM clients.

**Formula-first philosophy**: unless the user explicitly asks for raw
values, always prefer to fetch *formulas* first.  Formulas reveal the
calculation logic and business rules; values are just one snapshot of the
result.  Use ``get_sheet_formulas`` or the formula-aware sampling modes
before falling back to value-only inspection.

Tools provided:
  1. list_sheets          – list sheet names in a workbook
  2. get_sheet_data       – get sheet data as markdown / JSON / XML (auto‑sampled)
  3. get_sheet_formulas   – extract all formulas from a sheet  ← **start here**
  4. get_workbook_summary – lightweight structural overview of the workbook
  5. get_sheet_sample     – get a small representative sample of a sheet
"""

import os
import json
from typing import Annotated
from mcp.server.fastmcp import FastMCP
from pydantic import Field

from excel_reader_smart_sampler import (
    sheet_names,
    extract_sheet_data as extract_sheet_data_smart,
    extract_formulas,
    workbook_summary,
    DEFAULT_SAMPLE_ROWS,
)
from excel_sample_full import extract_sheet_data as extract_sheet_data_full
from column_n import extract_sheet_data as extract_sheet_data_column_n
from row_head_fetcher import extract_sheet_data as extract_sheet_data_row_head
from column_head_fetcher import extract_sheet_data as extract_sheet_data_column_head
from formatters import to_markdown, to_json, to_xml

VALID_MODES = ("smart_random", "full", "column_n", "row_head", "column_head")


mcp = FastMCP("Excel Analysis Server")


# ---------------------------------------------------------------------------
# Sampling dispatcher
# ---------------------------------------------------------------------------

def _dispatch_extract(mode: str, file_path: str, sheet_name: str,
                      max_sample_rows: int = DEFAULT_SAMPLE_ROWS,
                      nrows: int | None = None,
                      ncols: int | None = None,
                      num_columns: int = 10,
                      max_cols: int = 20) -> dict:
    """Route to the correct sampling strategy based on *mode*."""
    mode = mode.lower().strip()
    if mode not in VALID_MODES:
        raise ValueError(
            f"Invalid mode '{mode}'. Must be one of {VALID_MODES}."
        )

    if mode == "smart_random":
        return extract_sheet_data_smart(file_path, sheet_name,
                                        max_sample_rows=max_sample_rows)
    elif mode == "full":
        return extract_sheet_data_full(file_path, sheet_name,
                                       nrows=nrows, ncols=ncols)
    elif mode == "column_n":
        return extract_sheet_data_column_n(file_path, sheet_name,
                                           num_columns=num_columns)
    elif mode == "row_head":
        return extract_sheet_data_row_head(file_path, sheet_name,
                                           max_rows=max_sample_rows)
    elif mode == "column_head":
        return extract_sheet_data_column_head(file_path, sheet_name,
                                              max_cols=max_cols)


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

@mcp.tool()
def list_sheets(
    file_path: Annotated[str, Field(
        description=(
            "Absolute path to the .xlsx file on disk. "
            "Example: '/home/user/reports/q4_financials.xlsx'"
        ),
    )],
) -> str:
    """List all sheet names in an Excel workbook.

    Use this tool as the first step when you receive an Excel file and need to
    discover what sheets it contains before reading any data.  The result is a
    JSON array of sheet names that you can pass to other tools.

    Returns a JSON array of strings, e.g. ["Sales", "Costs", "Summary"].

    Example call:
        list_sheets(file_path="/home/user/reports/q4_financials.xlsx")
        → '["Sales", "Costs", "Summary"]'
    """
    names = sheet_names(file_path)
    return json.dumps(names)


@mcp.tool()
def get_sheet_data(
    file_path: Annotated[str, Field(
        description=(
            "Absolute path to the .xlsx file on disk. "
            "Example: '/data/report.xlsx', '/home/user/reports/q4_financials.xlsx'"
        ),
    )],
    sheet_name: Annotated[str, Field(
        description=(
            "Exact name of the sheet to read (case-sensitive, as returned by "
            "list_sheets). Example: 'Sales', 'P&L Summary', 'Sheet1'"
        ),
    )],
    format: Annotated[str, Field(
        default="markdown",
        description=(
            "Output format for the returned data. "
            "Set to 'markdown' for human-readable tables (default). "
            "Set to 'json' for structured programmatic output. "
            "Set to 'xml' for XML interchange. "
            "Example values: 'markdown', 'json', 'xml'"
        ),
    )] = "markdown",
    max_sample_rows: Annotated[int, Field(
        default=DEFAULT_SAMPLE_ROWS,
        description=(
            "Maximum number of rows to return **per sheet** across all detected "
            "data regions. Budget is divided across patches proportionally. "
            "Only used when mode is 'smart_random' or 'row_head'. "
            "Set to 50 for a smaller sample to conserve context tokens. "
            "Set to 200 for wider coverage. Default is 100. "
            "Example values: 30, 50, 100, 200"
        ),
    )] = DEFAULT_SAMPLE_ROWS,
    mode: Annotated[str, Field(
        default="smart_random",
        description=(
            "Sampling strategy to use. All modes return both formulas and values.\n"
            "Set to 'smart_random' (default) — ONLY use when exploring the Excel "
            "file for an overview or first-time inspection. Prioritises headers, "
            "formula rows, and a spread of head/mid/tail rows. Best general-purpose "
            "choice; keeps token usage low.\n"
            "Set to 'full' — use when comprehensively analysing a particular sheet "
            "for detailed analysis. Loads every row/column in the detected region. "
            "Use for small lookup tables or when you need an exhaustive dump.\n"
            "Set to 'column_n' — use for wide sheets where you only need a vertical "
            "slice. Loads the label column plus the next N columns. Ideal for "
            "financial models with many period columns.\n"
            "Set to 'row_head' — use to capture all patch headers and the first few "
            "data rows below each header. Budget is per sheet, divided across "
            "patches by size. Best for scanning non-priority sheets efficiently.\n"
            "Set to 'column_head' — use for financial sheets where columns are dates "
            "(e.g. Q1 2023, Q2 2023) and rows are entities (Income, Profit). "
            "Reads all rows but only the first N columns per patch, relating "
            "entities to their time-series values. Budget is per sheet.\n"
            "Example values: 'smart_random', 'full', 'column_n', 'row_head', "
            "'column_head'"
        ),
    )] = "smart_random",
    nrows: Annotated[int | None, Field(
        default=None,
        description=(
            "Maximum number of rows to load per region (only used in 'full' mode). "
            "Set to None to load all rows. "
            "Example values: None, 100, 500"
        ),
    )] = None,
    ncols: Annotated[int | None, Field(
        default=None,
        description=(
            "Maximum number of columns to load per region (only used in 'full' mode). "
            "Set to None to load all columns. "
            "Example values: None, 5, 10"
        ),
    )] = None,
    num_columns: Annotated[int, Field(
        default=10,
        description=(
            "Number of data columns to load after the label column "
            "(only used in 'column_n' mode). Default is 10. "
            "Example values: 3, 5, 10, 20"
        ),
    )] = 10,
    max_cols: Annotated[int, Field(
        default=20,
        description=(
            "Total column budget **per sheet** (only used in 'column_head' mode). "
            "Divided across patches proportionally by width. "
            "Set to 10 for a narrow view, 20 (default) for moderate coverage. "
            "Example values: 10, 15, 20, 30"
        ),
    )] = 20,
) -> str:
    """Read a sheet and return its data — including formulas — in the requested format.

    **Formulas are always included** in the output alongside values.  All five
    sampling modes extract both formulas and cached values for every cell they
    visit, so you never need a separate call just to get formulas from the
    sampled rows.  If you need *only* formulas (e.g. to trace cross-sheet
    dependencies), prefer ``get_sheet_formulas`` instead.

    Large sheets are automatically sampled to avoid context overflow: the
    sample always includes header rows, formula rows, and a spread of
    head / middle / tail data rows so the LLM can understand the sheet's
    structure and logic.

    The output contains detected data regions (the server handles unstructured
    sheets where headers may not be on the first row, or where blank rows and
    company banners precede the actual data).

    Returns a formatted string (Markdown table, JSON object, or XML document)
    containing the sheet's regions, row data, and formulas.

    Example calls:
        # Get a Markdown overview of the Sales sheet (sampled)
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")

        # Get the Costs sheet as JSON with a smaller sample
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Costs",
                       format="json", max_sample_rows=50)

        # Get every row using full mode for detailed analysis
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Lookups",
                       mode="full")

        # Get a vertical strip using column_n mode
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales",
                       mode="column_n", num_columns=5)

        # Capture all headers + first few rows per patch
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Revenue",
                       mode="row_head", max_sample_rows=50)

        # Financial sheet with date columns — capture entity rows + first columns
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="PnL",
                       mode="column_head", max_cols=15)
    """
    data = _dispatch_extract(mode, file_path, sheet_name,
                             max_sample_rows=max_sample_rows,
                             nrows=nrows, ncols=ncols,
                             num_columns=num_columns,
                             max_cols=max_cols)
    fmt = format.lower().strip()
    if fmt == "json":
        return to_json(data)
    if fmt == "xml":
        return to_xml(data)
    return to_markdown(data)


@mcp.tool()
def get_sheet_formulas(
    file_path: Annotated[str, Field(
        description=(
            "Absolute path to the .xlsx file on disk. "
            "Example: '/data/report.xlsx', '/home/user/reports/budget_model.xlsx'"
        ),
    )],
    sheet_name: Annotated[str, Field(
        description=(
            "Exact name of the sheet (case-sensitive). "
            "Example: 'Sales', 'Assumptions'"
        ),
    )],
) -> str:
    """Extract every formula from a sheet, returning each formula's cell address,
    the raw Excel formula string, and its last-cached value.

    **Use this tool first** when analysing a sheet — formulas reveal the
    calculation logic and business rules, which are more important than the
    raw values.  Only fall back to ``get_sheet_data`` for values when the user
    explicitly asks to inspect data values rather than formulas.

    This is also the best tool for tracing cross-sheet references
    (e.g. =Sales!D8-B4) and deeply nested derived quantities.

    Returns a JSON array of objects.  Each object has:
      - "address"      — cell reference, e.g. "D5"
      - "formula"      — raw Excel formula, e.g. "=B5*C5"
      - "cached_value" — last computed value stored in the file (may be null if
                         the file was never opened in Excel)

    Example calls:
        # Extract all formulas from the Sales sheet
        get_sheet_formulas(file_path="/data/report.xlsx", sheet_name="Sales")
        → '[{"address":"D5","formula":"=B5*C5","cached_value":999},...]'

        # Inspect cross-sheet formulas in the Summary sheet
        get_sheet_formulas(file_path="/data/report.xlsx",
                           sheet_name="Summary")
    """
    formulas = extract_formulas(file_path, sheet_name)
    return json.dumps(formulas, default=str)


@mcp.tool()
def get_workbook_summary(
    file_path: Annotated[str, Field(
        description=(
            "Absolute path to the .xlsx file on disk. "
            "Example: '/data/report.xlsx', '/home/user/reports/budget_model.xlsx'"
        ),
    )],
) -> str:
    """Return a lightweight structural summary of the entire workbook without
    reading all cell data.

    Use this tool FIRST when you receive a new Excel file.  It gives you an
    overview of every sheet — dimensions, how many data regions were detected,
    and how many formula cells exist — so you can decide which sheets to
    examine in detail with get_sheet_data or get_sheet_formulas.

    Returns a JSON object with:
      - "file"   — the file path
      - "sheets" — array of per-sheet summaries, each containing:
          - "name"          — sheet name
          - "max_row"       — total row count
          - "max_column"    — total column count
          - "regions"       — list of detected data-region descriptions
          - "region_count"  — number of data regions
          - "formula_count" — total formula cells in the sheet

    Example call:
        get_workbook_summary(file_path="/data/report.xlsx")
        → '{"file":"/data/report.xlsx","sheets":[{"name":"Sales",
            "max_row":500,"max_column":10,"region_count":2,
            "formula_count":48}, ...]}'
    """
    summary = workbook_summary(file_path)
    return json.dumps(summary, indent=2, default=str)


@mcp.tool()
def get_sheet_sample(
    file_path: Annotated[str, Field(
        description=(
            "Absolute path to the .xlsx file on disk. "
            "Example: '/data/report.xlsx', '/home/user/reports/inventory_2025.xlsx'"
        ),
    )],
    sheet_name: Annotated[str, Field(
        description=(
            "Exact name of the sheet (case-sensitive). "
            "Example: 'Sales', 'Raw Data'"
        ),
    )],
    max_rows: Annotated[int, Field(
        default=30,
        description=(
            "Maximum number of rows to include in the sample **per sheet**. "
            "Budget is divided across patches when using 'row_head' mode. "
            "Only used when mode is 'smart_random' or 'row_head'. "
            "Set to 10 for a tiny preview (just headers + a few rows). "
            "Set to 30 (default) for a moderate preview. "
            "Set to 50 for a slightly larger preview. "
            "Example values: 10, 20, 30, 50"
        ),
    )] = 30,
    format: Annotated[str, Field(
        default="markdown",
        description=(
            "Output format. "
            "Set to 'markdown' (default) for human-readable tables. "
            "Set to 'json' for structured output. "
            "Set to 'xml' for XML interchange. "
            "Example values: 'markdown', 'json', 'xml'"
        ),
    )] = "markdown",
    mode: Annotated[str, Field(
        default="smart_random",
        description=(
            "Sampling strategy to use. All modes return both formulas and values.\n"
            "Set to 'smart_random' (default) — ONLY use when exploring the Excel "
            "file for an overview or first-time inspection. Prioritises headers, "
            "formula rows, and a spread of head/mid/tail rows.\n"
            "Set to 'full' — use when comprehensively analysing a particular sheet "
            "for detailed analysis. Loads all data. Use only for small sheets.\n"
            "Set to 'column_n' — use for wide sheets when you only need a few "
            "columns. Loads the label column plus the next N columns.\n"
            "Set to 'row_head' — use to capture all patch headers and the first "
            "few data rows below each header. Budget is per sheet.\n"
            "Set to 'column_head' — use for financial sheets where columns are "
            "dates and rows are entities. Reads all rows, first N columns per "
            "patch. Budget is per sheet.\n"
            "Example values: 'smart_random', 'full', 'column_n', 'row_head', "
            "'column_head'"
        ),
    )] = "smart_random",
    num_columns: Annotated[int, Field(
        default=10,
        description=(
            "Number of data columns to load after the label column "
            "(only used in 'column_n' mode). Default is 10. "
            "Example values: 3, 5, 10"
        ),
    )] = 10,
    max_cols: Annotated[int, Field(
        default=20,
        description=(
            "Total column budget **per sheet** (only used in 'column_head' mode). "
            "Divided across patches proportionally by width. "
            "Example values: 10, 15, 20"
        ),
    )] = 20,
) -> str:
    """Get a very small, representative sample of a sheet — useful for a quick
    preview before deciding whether to request more data.

    Use this tool when you want a fast glance at a sheet's structure and content
    without pulling in a large amount of data.  It behaves like get_sheet_data
    but defaults to a much smaller sample (30 rows instead of 100).

    Returns the sampled data in the requested format (Markdown, JSON, or XML),
    including any formulas found in the sampled rows.  All five modes return
    both formulas and values.

    Example calls:
        # Quick Markdown preview of the Sales sheet
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales")

        # Tiny JSON preview (10 rows) to check column names
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Raw Data",
                         max_rows=10, format="json")

        # Preview using column_n mode
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales",
                         mode="column_n", num_columns=5)

        # Row head preview — all headers + first few data rows per patch
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Costs",
                         mode="row_head", max_rows=20)

        # Column head preview — entity labels + first date columns
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="PnL",
                         mode="column_head", max_cols=10)
    """
    data = _dispatch_extract(mode, file_path, sheet_name,
                             max_sample_rows=max_rows,
                             num_columns=num_columns,
                             max_cols=max_cols)
    fmt = format.lower().strip()
    if fmt == "json":
        return to_json(data)
    if fmt == "xml":
        return to_xml(data)
    return to_markdown(data)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
