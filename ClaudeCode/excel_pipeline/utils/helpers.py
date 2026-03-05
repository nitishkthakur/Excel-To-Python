"""Common utility functions for the Excel pipeline."""

import re
from typing import Any, Optional, Tuple
from datetime import datetime
import openpyxl.utils as xl_utils


def is_financial_date(value: Any) -> bool:
    """
    Check if a value represents a financial date/period.

    Financial date patterns include:
    - Integers: 2020, 2021, 2022, etc. (4-digit years)
    - Strings with year: "2020E", "2021E", "FY2020", "FY 2020", etc.
    - Datetime objects
    - Quarter patterns: "Q1 2020", "2020Q1", "Q1-20", etc.
    - Month patterns: "Jan 2020", "2020-01", "Jan-20", etc.

    Args:
        value: Value to check

    Returns:
        True if value is a financial date pattern

    Examples:
        >>> is_financial_date(2020)
        True
        >>> is_financial_date("2020E")
        True
        >>> is_financial_date("Q1 2020")
        True
        >>> is_financial_date("Revenue")
        False
    """
    if isinstance(value, datetime):
        return True

    if isinstance(value, int):
        # Check if 4-digit year (1900-2199)
        return 1900 <= value <= 2199

    if not isinstance(value, str):
        return False

    value = value.strip()

    # Pattern 1: Year only (2020E, FY2020, FY 2020, etc.)
    if re.match(r'^(FY\s*)?\d{4}[A-Za-z]?$', value):
        return True

    # Pattern 2: Quarter patterns (Q1 2020, 2020Q1, Q1-20, etc.)
    if re.match(r'^(Q[1-4]\s*[-\s]?\d{2,4}|\d{2,4}\s*[-\s]?Q[1-4])$', value, re.IGNORECASE):
        return True

    # Pattern 3: Month patterns (Jan 2020, 2020-01, Jan-20, etc.)
    month_names = r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
    if re.match(rf'^({month_names}\s*[-\s]?\d{{2,4}}|\d{{2,4}}\s*[-\s]?{month_names})', value, re.IGNORECASE):
        return True

    # Pattern 4: Date-like strings (2020-01, 2020-Q1, etc.)
    if re.match(r'^\d{4}[-/]\d{1,2}$', value):
        return True

    return False


def column_letter_to_index(col_letter: str) -> int:
    """
    Convert Excel column letter to 0-based index.

    Args:
        col_letter: Column letter (e.g., "A", "B", "AA")

    Returns:
        0-based column index

    Examples:
        >>> column_letter_to_index("A")
        0
        >>> column_letter_to_index("Z")
        25
        >>> column_letter_to_index("AA")
        26
    """
    return xl_utils.column_index_from_string(col_letter) - 1


def column_index_to_letter(col_index: int) -> str:
    """
    Convert 0-based column index to Excel column letter.

    Args:
        col_index: 0-based column index

    Returns:
        Column letter

    Examples:
        >>> column_index_to_letter(0)
        'A'
        >>> column_index_to_letter(25)
        'Z'
        >>> column_index_to_letter(26)
        'AA'
    """
    return xl_utils.get_column_letter(col_index + 1)


def parse_cell_reference(cell_ref: str) -> Tuple[str, int, str]:
    """
    Parse a cell reference into sheet, row, column.

    Args:
        cell_ref: Cell reference (e.g., "Sheet1!A5" or "A5")

    Returns:
        Tuple of (sheet_name, row, column_letter)

    Examples:
        >>> parse_cell_reference("Sheet1!A5")
        ('Sheet1', 5, 'A')
        >>> parse_cell_reference("A5")
        ('', 5, 'A')
    """
    if '!' in cell_ref:
        sheet, cell = cell_ref.split('!', 1)
        sheet = sheet.strip("'\"")
    else:
        sheet = ""
        cell = cell_ref

    # Extract column letters and row number
    match = re.match(r'([A-Z]+)(\d+)', cell)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")

    col_letter = match.group(1)
    row_num = int(match.group(2))

    return sheet, row_num, col_letter


def make_cell_reference(sheet: str, row: int, col: str, absolute_row: bool = False, absolute_col: bool = False) -> str:
    """
    Create a cell reference string.

    Args:
        sheet: Sheet name (empty string for no sheet)
        row: Row number (1-based)
        col: Column letter
        absolute_row: If True, make row absolute ($A1)
        absolute_col: If True, make column absolute (A$1)

    Returns:
        Cell reference string

    Examples:
        >>> make_cell_reference("Sheet1", 5, "A")
        'Sheet1!A5'
        >>> make_cell_reference("", 5, "A", absolute_col=True)
        '$A5'
    """
    col_prefix = '$' if absolute_col else ''
    row_prefix = '$' if absolute_row else ''

    cell = f"{col_prefix}{col}{row_prefix}{row}"

    if sheet:
        # Escape sheet name if it contains spaces
        if ' ' in sheet:
            sheet = f"'{sheet}'"
        return f"{sheet}!{cell}"

    return cell


def rgb_to_hex(rgb: Optional[Any]) -> Optional[str]:
    """
    Convert RGB string or object to hex color code.

    Args:
        rgb: RGB string (e.g., "FF000000") or openpyxl RGB object

    Returns:
        Hex color code (e.g., "#000000") or None

    Examples:
        >>> rgb_to_hex("FF000000")
        '#000000'
        >>> rgb_to_hex("FFFFFFFF")
        '#FFFFFF'
        >>> rgb_to_hex(None)
        None
    """
    if not rgb:
        return None

    # Convert RGB object to string if needed
    rgb_str = str(rgb) if not isinstance(rgb, str) else rgb

    # Remove alpha channel if present (first 2 chars)
    if len(rgb_str) == 8:
        rgb_str = rgb_str[2:]

    return f"#{rgb_str}"


def hex_to_rgb(hex_color: Optional[str]) -> Optional[str]:
    """
    Convert hex color code to RGB string.

    Args:
        hex_color: Hex color code (e.g., "#000000")

    Returns:
        RGB string (e.g., "FF000000") or None

    Examples:
        >>> hex_to_rgb("#000000")
        'FF000000'
        >>> hex_to_rgb("#FFFFFF")
        'FFFFFFFF'
        >>> hex_to_rgb(None)
        None
    """
    if not hex_color:
        return None

    hex_color = hex_color.lstrip('#')
    # Add alpha channel (FF = fully opaque)
    return f"FF{hex_color}"


def safe_str(value: Any) -> str:
    """
    Safely convert value to string.

    Args:
        value: Any value

    Returns:
        String representation or empty string if None

    Examples:
        >>> safe_str(123)
        '123'
        >>> safe_str(None)
        ''
        >>> safe_str("test")
        'test'
    """
    return str(value) if value is not None else ""


def generate_labels(count: int, prefix: str = "Line") -> list:
    """
    Generate sequential labels.

    Args:
        count: Number of labels to generate
        prefix: Label prefix (default: "Line")

    Returns:
        List of labels

    Examples:
        >>> generate_labels(3)
        ['Line1', 'Line2', 'Line3']
        >>> generate_labels(2, "Col")
        ['Col1', 'Col2']
    """
    return [f"{prefix}{i+1}" for i in range(count)]
