"""
formula_parser.py — Parse and transform Excel formulas.

Key capabilities:
  1. Extract cell references from formulas
  2. Convert A1-style ↔ R1C1-style references (for group detection)
  3. Normalize formulas to R1C1 patterns
  4. Reconstruct A1 formulas from R1C1 patterns

Design: all functions are pure and stateless; heavy lifting is
regex-based with careful masking of string literals and sheet names.
"""

from __future__ import annotations

import re
from typing import Any, Optional

from openpyxl.utils import get_column_letter, column_index_from_string

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MAX_COL = 16_384          # XFD
MAX_ROW = 1_048_576

# ---------------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------------

def col_to_num(col_letter: str) -> int:
    """Convert column letter(s) to 1-based number.  'A' → 1, 'Z' → 26."""
    return column_index_from_string(col_letter)


def num_to_col(num: int) -> str:
    """Convert 1-based column number to letters.  1 → 'A'."""
    return get_column_letter(num)


def cell_to_rowcol(cell_ref: str) -> tuple[int, int]:
    """'B3' → (3, 2).  Strips dollar signs."""
    ref = cell_ref.replace("$", "")
    m = re.match(r"([A-Z]{1,3})(\d+)", ref)
    if not m:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return int(m.group(2)), col_to_num(m.group(1))


def rowcol_to_cell(row: int, col: int) -> str:
    """(3, 2) → 'B3'."""
    return f"{num_to_col(col)}{row}"

# ---------------------------------------------------------------------------
# String-literal masking (so we don't confuse "A1" inside a string with a ref)
# ---------------------------------------------------------------------------

_STRING_RE = re.compile(r'"[^"]*"')

def _mask_strings(formula: str) -> tuple[str, list[str]]:
    strings: list[str] = []
    def _repl(m: re.Match) -> str:
        strings.append(m.group())
        return f"\x00S{len(strings)-1}\x00"
    return _STRING_RE.sub(_repl, formula), strings

def _unmask_strings(text: str, strings: list[str]) -> str:
    for i, s in enumerate(strings):
        text = text.replace(f"\x00S{i}\x00", s)
    return text

# ---------------------------------------------------------------------------
# Reference regex
# ---------------------------------------------------------------------------

# Matches optional sheet prefix (quoted or plain) followed by cell ref.
# Captures: (1) full_sheet_or_empty  (2)$? (3)COL (4)$? (5)ROW
# Optionally followed by a range part  :(6)$? (7)COL (8)$? (9)ROW
_SHEET_PREFIX = r"(?:(?:'(?:[^']|'')*'|[A-Za-z0-9_\\.]+)!)?"
_SINGLE_CELL  = r"(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})"
_RANGE_TAIL   = r"(?::(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6}))?"

# Full pattern — we use a verbose version with named groups below for clarity;
# here we compile a simpler one for performance.
_REF_RE = re.compile(
    r"(?P<sheet>(?:'(?:[^']|'')*'|[A-Za-z0-9_\\.]+)!)?"
    r"(?P<ca1>\$?)(?P<c1>[A-Z]{1,3})(?P<ra1>\$?)(?P<r1>[1-9]\d{0,6})"
    r"(?::(?P<ca2>\$?)(?P<c2>[A-Z]{1,3})(?P<ra2>\$?)(?P<r2>[1-9]\d{0,6}))?"
)

# ---------------------------------------------------------------------------
# Public: extract_references
# ---------------------------------------------------------------------------

class CellReference:
    """One parsed cell reference (or range)."""
    __slots__ = (
        "sheet", "col", "row", "col_abs", "row_abs",
        "end_col", "end_row", "end_col_abs", "end_row_abs",
        "is_range", "original",
    )

    def __init__(self, *, sheet: str | None, col: int, row: int,
                 col_abs: bool, row_abs: bool,
                 end_col: int | None = None, end_row: int | None = None,
                 end_col_abs: bool = False, end_row_abs: bool = False,
                 is_range: bool = False, original: str = ""):
        self.sheet = sheet
        self.col = col
        self.row = row
        self.col_abs = col_abs
        self.row_abs = row_abs
        self.end_col = end_col
        self.end_row = end_row
        self.end_col_abs = end_col_abs
        self.end_row_abs = end_row_abs
        self.is_range = is_range
        self.original = original

    # Convenience: iterate individual cells in a range
    def cells(self) -> list[tuple[int, int]]:
        """Return list of (row, col) for every cell in this reference."""
        if not self.is_range:
            return [(self.row, self.col)]
        out = []
        for r in range(self.row, (self.end_row or self.row) + 1):
            for c in range(self.col, (self.end_col or self.col) + 1):
                out.append((r, c))
        return out


def extract_references(formula: str) -> list[CellReference]:
    """Extract every cell/range reference from an Excel A1-style formula."""
    if not formula or not formula.startswith("="):
        return []
    masked, _strs = _mask_strings(formula)
    refs: list[CellReference] = []
    for m in _REF_RE.finditer(masked):
        start, end = m.start(), m.end()
        # Skip if preceded by alpha/underscore (part of a name)
        if start > 0 and (masked[start - 1].isalpha() or masked[start - 1] == '_'):
            continue
        # Skip if followed by '(' — function name
        if end < len(masked) and masked[end] == '(':
            continue

        sheet_raw = m.group("sheet")
        sheet = sheet_raw.rstrip("!").strip("'") if sheet_raw else None

        col1 = col_to_num(m.group("c1"))
        row1 = int(m.group("r1"))
        ca1 = m.group("ca1") == "$"
        ra1 = m.group("ra1") == "$"

        is_range = m.group("c2") is not None
        if is_range:
            col2 = col_to_num(m.group("c2"))
            row2 = int(m.group("r2"))
            ca2 = m.group("ca2") == "$"
            ra2 = m.group("ra2") == "$"
        else:
            col2, row2, ca2, ra2 = None, None, False, False

        refs.append(CellReference(
            sheet=sheet, col=col1, row=row1, col_abs=ca1, row_abs=ra1,
            end_col=col2, end_row=row2, end_col_abs=ca2, end_row_abs=ra2,
            is_range=is_range, original=m.group(),
        ))
    return refs

# ---------------------------------------------------------------------------
# Public: to_r1c1_pattern  /  r1c1_to_a1
# ---------------------------------------------------------------------------

def _single_to_r1c1(col_letter: str, row: int, col_abs: bool, row_abs: bool,
                     from_row: int, from_col: int) -> str:
    """Convert one cell part to R1C1 notation relative to (from_row, from_col)."""
    col_num = col_to_num(col_letter)
    if row_abs:
        r_part = f"R{row}"
    else:
        dr = row - from_row
        r_part = f"R[{dr}]" if dr != 0 else "R"
    if col_abs:
        c_part = f"C{col_num}"
    else:
        dc = col_num - from_col
        c_part = f"C[{dc}]" if dc != 0 else "C"
    return r_part + c_part


def to_r1c1_pattern(formula: str, cell_row: int, cell_col: int) -> str:
    """Convert an A1-style formula to an R1C1-style *pattern* string.

    The pattern is relative to (cell_row, cell_col) so that identical
    dragged formulas produce the same pattern string.
    """
    if not formula or not formula.startswith("="):
        return formula or ""

    masked, strs = _mask_strings(formula)

    # Collect spans to replace (right-to-left safe)
    replacements: list[tuple[int, int, str]] = []
    for m in _REF_RE.finditer(masked):
        start, end = m.start(), m.end()
        if start > 0 and (masked[start - 1].isalpha() or masked[start - 1] == '_'):
            continue
        if end < len(masked) and masked[end] == '(':
            continue

        sheet_raw = m.group("sheet") or ""

        ca1 = m.group("ca1") == "$"
        c1  = m.group("c1")
        ra1 = m.group("ra1") == "$"
        r1  = int(m.group("r1"))
        part1 = _single_to_r1c1(c1, r1, ca1, ra1, cell_row, cell_col)

        if m.group("c2") is not None:
            ca2 = m.group("ca2") == "$"
            c2  = m.group("c2")
            ra2 = m.group("ra2") == "$"
            r2  = int(m.group("r2"))
            part2 = _single_to_r1c1(c2, r2, ca2, ra2, cell_row, cell_col)
            r1c1 = f"{sheet_raw}{part1}:{part2}"
        else:
            r1c1 = f"{sheet_raw}{part1}"

        replacements.append((start, end, r1c1))

    result = masked
    for s, e, rep in reversed(replacements):
        result = result[:s] + rep + result[e:]
    return _unmask_strings(result, strs)


# ---------------------------------------------------------------------------
# R1C1 → A1 conversion (used to regenerate A1 formulas from a pattern)
# ---------------------------------------------------------------------------

_R1C1_CELL_RE = re.compile(
    r"(?P<sheet>(?:'(?:[^']|'')*'|[A-Za-z0-9_\\.]+)!)?"
    r"R(?P<rabs>\d+|\[[-+]?\d+\])?"
    r"C(?P<cabs>\d+|\[[-+]?\d+\])?"
    r"(?::R(?P<rabs2>\d+|\[[-+]?\d+\])?C(?P<cabs2>\d+|\[[-+]?\d+\])?)?"
)


def _rc_part(part: str | None, base: int) -> tuple[int, bool]:
    """Parse one R or C component.  Returns (absolute_number, is_absolute)."""
    if part is None:
        return base, False       # bare R or C → offset 0
    if part.startswith("["):
        return base + int(part[1:-1]), False
    return int(part), True


def r1c1_to_a1(pattern: str, target_row: int, target_col: int) -> str:
    """Convert an R1C1 pattern to an A1 formula for the given cell position."""
    masked, strs = _mask_strings(pattern)

    def _repl(m: re.Match) -> str:
        sheet = m.group("sheet") or ""
        row, rabs = _rc_part(m.group("rabs"), target_row)
        col, cabs = _rc_part(m.group("cabs"), target_col)
        col_str = ("$" if cabs else "") + num_to_col(col)
        row_str = ("$" if rabs else "") + str(row)
        cell_a1 = f"{sheet}{col_str}{row_str}"
        if m.group("rabs2") is not None or m.group("cabs2") is not None:
            row2, rabs2 = _rc_part(m.group("rabs2"), target_row)
            col2, cabs2 = _rc_part(m.group("cabs2"), target_col)
            col2_str = ("$" if cabs2 else "") + num_to_col(col2)
            row2_str = ("$" if rabs2 else "") + str(row2)
            cell_a1 += f":{col2_str}{row2_str}"
        return cell_a1

    result = _R1C1_CELL_RE.sub(_repl, masked)
    return _unmask_strings(result, strs)


# ---------------------------------------------------------------------------
# Dependency helpers
# ---------------------------------------------------------------------------

def get_referenced_cells(formula: str, src_sheet: str,
                         defined_names: dict[str, list[tuple[str, int, int]]] | None = None,
                         ) -> list[tuple[str, int, int]]:
    """Return a flat list of (sheet, row, col) that *formula* depends on.

    *defined_names* maps name → [(sheet, row, col), ...] for named ranges.
    """
    refs = extract_references(formula)
    cells: list[tuple[str, int, int]] = []
    for ref in refs:
        sheet = ref.sheet or src_sheet
        if ref.is_range:
            for r in range(ref.row, (ref.end_row or ref.row) + 1):
                for c in range(ref.col, (ref.end_col or ref.col) + 1):
                    cells.append((sheet, r, c))
        else:
            cells.append((sheet, ref.row, ref.col))
    return cells
