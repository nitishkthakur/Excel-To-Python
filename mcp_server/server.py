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
from mcp.server.fastmcp import FastMCP

from excel_reader_smart_sampler import (
    sheet_names,
    extract_sheet_data as extract_sheet_data_smart,
    extract_formulas,
    workbook_summary,
    DEFAULT_SAMPLE_ROWS,
)
from excel_sample_full import extract_sheet_data as extract_sheet_data_full
from column_n import extract_sheet_data as extract_sheet_data_column_n
from formatters import to_markdown, to_json, to_xml

VALID_MODES = ("smart_random", "full", "column_n")


mcp = FastMCP("Excel Analysis Server")


# ---------------------------------------------------------------------------
# Sampling dispatcher
# ---------------------------------------------------------------------------

def _dispatch_extract(mode: str, file_path: str, sheet_name: str,
                      max_sample_rows: int = DEFAULT_SAMPLE_ROWS,
                      nrows: int | None = None,
                      ncols: int | None = None,
                      num_columns: int = 10) -> dict:
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


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

@mcp.tool()
def list_sheets(file_path: str) -> str:
    """List all sheet names in an Excel workbook.

    Use this tool as the first step when you receive an Excel file and need to
    discover what sheets it contains before reading any data.  The result is a
    JSON array of sheet names that you can pass to other tools.

    Returns a JSON array of strings, e.g. ["Sales", "Costs", "Summary"].

    Args:
        file_path: Absolute path to the .xlsx file on disk.
            - Example: "/home/user/reports/q4_financials.xlsx"
            - Example: "/data/uploads/inventory_2025.xlsx"

    Example call:
        list_sheets(file_path="/home/user/reports/q4_financials.xlsx")
        → '["Sales", "Costs", "Summary"]'
    """
    names = sheet_names(file_path)
    return json.dumps(names)


@mcp.tool()
def get_sheet_data(
    file_path: str,
    sheet_name: str,
    format: str = "markdown",
    max_sample_rows: int = DEFAULT_SAMPLE_ROWS,
    mode: str = "smart_random",
    nrows: int | None = None,
    ncols: int | None = None,
    num_columns: int = 10,
) -> str:
    """Read a sheet and return its data — including formulas — in the requested format.

    **Formulas are always included** in the output alongside values.  All three
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

    Args:
        file_path: Absolute path to the .xlsx file on disk.
            - Example: "/home/user/reports/q4_financials.xlsx"
            - Example: "/data/uploads/inventory_2025.xlsx"
        sheet_name: Exact name of the sheet to read (case-sensitive, as returned
            by list_sheets).
            - Example: "Sales"
            - Example: "P&L Summary"
            - Example: "Sheet1"
        format: Output format for the returned data. One of "markdown", "json",
            or "xml".  Default is "markdown".
            - Example: "markdown" — returns a human-readable Markdown table
            - Example: "json"     — returns a structured JSON object
            - Example: "xml"      — returns an XML document
        max_sample_rows: Maximum number of rows to return per detected data
            region.  Default is 100.  Increase for wider coverage; decrease to
            save context tokens.  Only used when mode is "smart_random".
            - Example: 50  — smaller sample to conserve context
            - Example: 200 — larger sample for more thorough analysis
        mode: Sampling strategy to use.  One of "smart_random", "full", or
            "column_n".  Default is "smart_random".  All three modes return
            both formulas and values.
            - "smart_random" — (default) prioritises headers, formula rows,
              and a spread of head/mid/tail rows.  Best general-purpose choice;
              keeps token usage low while capturing the calculation logic.
            - "full" — loads every row/column in the detected region.  Use for
              small lookup tables or when you need an exhaustive dump; avoid on
              sheets with thousands of rows.
            - "column_n" — loads the label column plus the next N columns.
              Ideal for wide sheets where you only need a vertical slice (e.g.
              a financial model with many period columns).
        nrows: Maximum number of rows to load (only used in "full" mode).
            None means load all rows.
        ncols: Maximum number of columns to load (only used in "full" mode).
            None means load all columns.
        num_columns: Number of data columns to load after the label column
            (only used in "column_n" mode).  Default is 10.

    Example calls:
        # Get a Markdown overview of the Sales sheet (sampled)
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")

        # Get the Costs sheet as JSON with a smaller sample
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Costs",
                       format="json", max_sample_rows=50)

        # Get every row using full mode
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Lookups",
                       mode="full")

        # Get a vertical strip using column_n mode
        get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales",
                       mode="column_n", num_columns=5)
    """
    data = _dispatch_extract(mode, file_path, sheet_name,
                             max_sample_rows=max_sample_rows,
                             nrows=nrows, ncols=ncols,
                             num_columns=num_columns)
    fmt = format.lower().strip()
    if fmt == "json":
        return to_json(data)
    if fmt == "xml":
        return to_xml(data)
    return to_markdown(data)


@mcp.tool()
def get_sheet_formulas(file_path: str, sheet_name: str) -> str:
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

    Args:
        file_path: Absolute path to the .xlsx file on disk.
            - Example: "/home/user/reports/q4_financials.xlsx"
            - Example: "/data/uploads/budget_model.xlsx"
        sheet_name: Exact name of the sheet (case-sensitive).
            - Example: "Sales"
            - Example: "Assumptions"

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
def get_workbook_summary(file_path: str) -> str:
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

    Args:
        file_path: Absolute path to the .xlsx file on disk.
            - Example: "/home/user/reports/q4_financials.xlsx"
            - Example: "/data/uploads/budget_model.xlsx"
            - Example: "C:/Users/alice/Documents/forecast.xlsx"

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
    file_path: str,
    sheet_name: str,
    max_rows: int = 30,
    format: str = "markdown",
    mode: str = "smart_random",
    num_columns: int = 10,
) -> str:
    """Get a very small, representative sample of a sheet — useful for a quick
    preview before deciding whether to request more data.

    Use this tool when you want a fast glance at a sheet's structure and content
    without pulling in a large amount of data.  It behaves like get_sheet_data
    but defaults to a much smaller sample (30 rows instead of 100).

    Returns the sampled data in the requested format (Markdown, JSON, or XML),
    including any formulas found in the sampled rows.  All three modes return
    both formulas and values.

    Args:
        file_path: Absolute path to the .xlsx file on disk.
            - Example: "/home/user/reports/q4_financials.xlsx"
            - Example: "/data/uploads/inventory_2025.xlsx"
        sheet_name: Exact name of the sheet (case-sensitive).
            - Example: "Sales"
            - Example: "Raw Data"
        max_rows: Maximum number of rows to include in the sample.  Default 30.
            Only used when mode is "smart_random".
            - Example: 10 — tiny preview, just headers + a few rows
            - Example: 30 — (default) a moderate preview
            - Example: 50 — slightly larger preview
        format: Output format. One of "markdown", "json", or "xml".
            Default is "markdown".
            - Example: "markdown"
            - Example: "json"
        mode: Sampling strategy to use.  One of "smart_random", "full", or
            "column_n".  Default is "smart_random".  All three modes return
            both formulas and values.
            - "smart_random" — (default) prioritises headers, formula rows,
              and a spread of head/mid/tail rows.  Best general-purpose choice.
            - "full" — loads all data.  Use only for small sheets.
            - "column_n" — loads the label column plus the next N columns.
              Great for wide sheets when you only need a few columns.
        num_columns: Number of data columns to load after the label column
            (only used in "column_n" mode).  Default is 10.

    Example calls:
        # Quick Markdown preview of the Sales sheet
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales")

        # Tiny JSON preview (10 rows) to check column names
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Raw Data",
                         max_rows=10, format="json")

        # Preview using column_n mode
        get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales",
                         mode="column_n", num_columns=5)
    """
    data = _dispatch_extract(mode, file_path, sheet_name,
                             max_sample_rows=max_rows,
                             num_columns=num_columns)
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
