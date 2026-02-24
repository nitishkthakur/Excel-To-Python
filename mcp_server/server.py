"""
MCP Server — exposes Excel workbook analysis tools to LLM clients.

Tools provided:
  1. list_sheets        – list sheet names in a workbook
  2. get_sheet_data     – get sheet data as markdown / JSON / XML (auto‑sampled)
  3. get_sheet_formulas – extract all formulas from a sheet
  4. get_workbook_summary – lightweight structural overview of the workbook
  5. get_sheet_sample   – get a small representative sample of a sheet
"""

import os
import json
from mcp.server.fastmcp import FastMCP

from excel_reader import (
    sheet_names,
    extract_sheet_data,
    extract_formulas,
    workbook_summary,
    DEFAULT_SAMPLE_ROWS,
)
from formatters import to_markdown, to_json, to_xml


mcp = FastMCP("Excel Analysis Server")


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

@mcp.tool()
def list_sheets(file_path: str) -> str:
    """
    List all sheet names in an Excel workbook.

    Args:
        file_path: Absolute path to the .xlsx file.

    Returns:
        JSON array of sheet names.
    """
    names = sheet_names(file_path)
    return json.dumps(names)


@mcp.tool()
def get_sheet_data(
    file_path: str,
    sheet_name: str,
    format: str = "markdown",
    max_sample_rows: int = DEFAULT_SAMPLE_ROWS,
    full: bool = False,
) -> str:
    """
    Read a sheet and return its data (with formulas) in the requested format.

    Large sheets are automatically sampled to avoid context overflow.
    The sample always includes headers, formula rows, and a spread of
    head / middle / tail data rows.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Name of the sheet to read.
        format: Output format — "markdown", "json", or "xml".
        max_sample_rows: Maximum rows to return per data region (default 100).
        full: Set True to disable sampling and return every row.

    Returns:
        The sheet data in the requested format.
    """
    data = extract_sheet_data(file_path, sheet_name,
                              max_sample_rows=max_sample_rows, full=full)
    fmt = format.lower().strip()
    if fmt == "json":
        return to_json(data)
    if fmt == "xml":
        return to_xml(data)
    return to_markdown(data)


@mcp.tool()
def get_sheet_formulas(file_path: str, sheet_name: str) -> str:
    """
    Extract every formula from a sheet.

    Returns a JSON array of objects with address, formula, and cached value.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Name of the sheet.
    """
    formulas = extract_formulas(file_path, sheet_name)
    return json.dumps(formulas, default=str)


@mcp.tool()
def get_workbook_summary(file_path: str) -> str:
    """
    Return a lightweight structural summary of the entire workbook.

    Includes sheet names, dimensions, detected data regions, and formula
    counts.  Use this first to understand the workbook before diving into
    individual sheets.

    Args:
        file_path: Absolute path to the .xlsx file.
    """
    summary = workbook_summary(file_path)
    return json.dumps(summary, indent=2, default=str)


@mcp.tool()
def get_sheet_sample(
    file_path: str,
    sheet_name: str,
    max_rows: int = 30,
    format: str = "markdown",
) -> str:
    """
    Get a very small sample of a sheet — useful for a quick preview before
    requesting more data.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Name of the sheet.
        max_rows: Maximum rows to include (default 30).
        format: Output format — "markdown", "json", or "xml".
    """
    data = extract_sheet_data(file_path, sheet_name,
                              max_sample_rows=max_rows)
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
