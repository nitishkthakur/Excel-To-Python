"""
Formula pattern detection, grouping, and vectorization.

Detects formulas that are "dragged" copies of the same pattern (vertically
or horizontally) and groups them so the code generator can emit compact loops
instead of one line per cell.

Also detects external workbook references ([Book.xlsx]Sheet!Cell).
"""

import re
from collections import defaultdict

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from formula_converter import col_letter_to_index, index_to_col_letter


# ---------------------------------------------------------------------------
# Reference extraction (preserves absolute / relative info and external refs)
# ---------------------------------------------------------------------------

# External workbook reference patterns:
#   [Book.xlsx]Sheet!$A$1   or   '[path\Book.xlsx]Sheet Name'!A1:B2
_EXT_RANGE_PATTERN = re.compile(
    r"'?\[([^\]]+)\]([^'!]+)'?!"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
    r":"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)
_EXT_CELL_PATTERN = re.compile(
    r"'?\[([^\]]+)\]([^'!]+)'?!"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)

# Internal cross-sheet range:  Sheet1!A1:B2  or  'Sheet Name'!A1:B2
_INT_RANGE_PATTERN = re.compile(
    r"(?:'([^'\[\]]+)'|([A-Za-z_]\w*))!"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
    r":"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)

# Internal cross-sheet cell:  Sheet1!A1  or  'Sheet Name'!A1
_INT_CELL_PATTERN = re.compile(
    r"(?:'([^'\[\]]+)'|([A-Za-z_]\w*))!"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)

# Same-sheet range:  A1:B2
_LOCAL_RANGE_PATTERN = re.compile(
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
    r":"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)

# Same-sheet cell:  A1
_LOCAL_CELL_PATTERN = re.compile(
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
)

# Table structured reference – not vectorized specially
_TABLE_REF_PATTERN = re.compile(r"(\w+)\[([^\]]*)\]")


class Reference:
    """A single cell or range reference extracted from a formula."""

    __slots__ = (
        "kind",            # 'cell' | 'range'
        "external_file",   # str or None
        "sheet",           # resolved sheet name
        "col_abs", "col", "col_idx",
        "row_abs", "row",
        "end_col_abs", "end_col", "end_col_idx",
        "end_row_abs", "end_row",
        "start", "end", "raw",
    )

    def __init__(self, **kw):
        for k in self.__slots__:
            setattr(self, k, kw.get(k))
        if self.col:
            self.col_idx = col_letter_to_index(self.col)
        if self.end_col:
            self.end_col_idx = col_letter_to_index(self.end_col)


def _in_string(formula, pos):
    """Return True if *pos* falls inside a double-quoted string literal."""
    in_str = False
    i = 0
    while i < pos:
        if formula[i] == '"':
            if in_str and i + 1 < len(formula) and formula[i + 1] == '"':
                i += 2
                continue
            in_str = not in_str
        i += 1
    return in_str


def extract_references(formula, current_sheet):
    """Extract every cell/range reference with absolute-marker info.

    Returns a list of Reference objects sorted by start position.
    """
    if formula.startswith("="):
        formula = formula[1:]

    refs = []
    occupied = set()  # character positions already claimed by a match

    def _occupied(start, end):
        return any(p in occupied for p in range(start, end))

    def _claim(start, end):
        occupied.update(range(start, end))

    # --- external ranges ---
    for m in _EXT_RANGE_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        refs.append(Reference(
            kind="range", external_file=m.group(1), sheet=m.group(2),
            col_abs=m.group(3) == "$", col=m.group(4),
            row_abs=m.group(5) == "$", row=int(m.group(6)),
            end_col_abs=m.group(7) == "$", end_col=m.group(8),
            end_row_abs=m.group(9) == "$", end_row=int(m.group(10)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    # --- external cells ---
    for m in _EXT_CELL_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        refs.append(Reference(
            kind="cell", external_file=m.group(1), sheet=m.group(2),
            col_abs=m.group(3) == "$", col=m.group(4),
            row_abs=m.group(5) == "$", row=int(m.group(6)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    # --- internal cross-sheet ranges ---
    for m in _INT_RANGE_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        sheet = m.group(1) or m.group(2)
        refs.append(Reference(
            kind="range", external_file=None, sheet=sheet,
            col_abs=m.group(3) == "$", col=m.group(4),
            row_abs=m.group(5) == "$", row=int(m.group(6)),
            end_col_abs=m.group(7) == "$", end_col=m.group(8),
            end_row_abs=m.group(9) == "$", end_row=int(m.group(10)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    # --- internal cross-sheet cells ---
    for m in _INT_CELL_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        sheet = m.group(1) or m.group(2)
        refs.append(Reference(
            kind="cell", external_file=None, sheet=sheet,
            col_abs=m.group(3) == "$", col=m.group(4),
            row_abs=m.group(5) == "$", row=int(m.group(6)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    # --- same-sheet ranges ---
    for m in _LOCAL_RANGE_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        # Avoid matching inside function names / table refs
        if m.start() > 0 and (formula[m.start() - 1].isalpha() or formula[m.start() - 1] == '_'):
            continue
        refs.append(Reference(
            kind="range", external_file=None, sheet=current_sheet,
            col_abs=m.group(1) == "$", col=m.group(2),
            row_abs=m.group(3) == "$", row=int(m.group(4)),
            end_col_abs=m.group(5) == "$", end_col=m.group(6),
            end_row_abs=m.group(7) == "$", end_row=int(m.group(8)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    # --- same-sheet cells ---
    for m in _LOCAL_CELL_PATTERN.finditer(formula):
        if _in_string(formula, m.start()) or _occupied(m.start(), m.end()):
            continue
        if m.start() > 0 and (formula[m.start() - 1].isalpha() or formula[m.start() - 1] == '_'):
            continue
        refs.append(Reference(
            kind="cell", external_file=None, sheet=current_sheet,
            col_abs=m.group(1) == "$", col=m.group(2),
            row_abs=m.group(3) == "$", row=int(m.group(4)),
            start=m.start(), end=m.end(), raw=m.group(),
        ))
        _claim(m.start(), m.end())

    refs.sort(key=lambda r: r.start)
    return refs


# ---------------------------------------------------------------------------
# Pattern normalisation
# ---------------------------------------------------------------------------

def compute_pattern(formula, current_sheet, cell_col_idx, cell_row):
    """Return a hashable *pattern key* for a formula.

    Two formulas that are copies of each other (dragged in Excel) will have
    the same pattern key.  The key encodes:
      * the formula skeleton (operators, function names, literals)
      * for each reference: whether col/row is absolute or relative, and
        the offset for relative parts.

    Returns (pattern_key, refs).
    """
    raw = formula[1:] if formula.startswith("=") else formula
    refs = extract_references(formula, current_sheet)

    tokens = []
    for ref in refs:
        if ref.kind == "cell":
            col_part = ("abs", ref.col) if ref.col_abs else ("rel", ref.col_idx - cell_col_idx)
            row_part = ("abs", ref.row) if ref.row_abs else ("rel", ref.row - cell_row)
            ext = ref.external_file
            sheet_part = ref.sheet if ref.sheet != current_sheet or ext else None
            tokens.append(("cell", ext, sheet_part, col_part, row_part))
        else:  # range
            sc = ("abs", ref.col) if ref.col_abs else ("rel", ref.col_idx - cell_col_idx)
            sr = ("abs", ref.row) if ref.row_abs else ("rel", ref.row - cell_row)
            ec = ("abs", ref.end_col) if ref.end_col_abs else ("rel", ref.end_col_idx - cell_col_idx)
            er = ("abs", ref.end_row) if ref.end_row_abs else ("rel", ref.end_row - cell_row)
            ext = ref.external_file
            sheet_part = ref.sheet if ref.sheet != current_sheet or ext else None
            tokens.append(("range", ext, sheet_part, sc, sr, ec, er))

    # Build skeleton by replacing refs with numbered placeholders
    skeleton = raw
    for i, ref in enumerate(reversed(refs)):
        idx = len(refs) - 1 - i
        skeleton = skeleton[:ref.start] + f"@{idx}" + skeleton[ref.end:]

    return (skeleton, tuple(tokens)), refs


# ---------------------------------------------------------------------------
# Grouping
# ---------------------------------------------------------------------------

def _find_contiguous_runs(cells, is_adjacent_fn):
    """Split a sorted list into maximal contiguous runs."""
    if not cells:
        return []
    runs = [[cells[0]]]
    for c in cells[1:]:
        if is_adjacent_fn(runs[-1][-1], c):
            runs[-1].append(c)
        else:
            runs.append([c])
    return runs


def group_formulas(formula_cells):
    """Detect vectorisable groups among *formula_cells*.

    Returns (groups, singles) where
      * groups: list of dicts  { 'cells': [...], 'direction': 'vertical'|'horizontal' }
      * singles: list of (sheet, col, row, formula, cell_info)
    """
    # compute patterns
    cell_pats = []
    for sheet, col, row, formula, cell_info in formula_cells:
        col_idx = col_letter_to_index(col)
        try:
            pkey, refs = compute_pattern(formula, sheet, col_idx, row)
        except Exception:
            pkey, refs = None, []
        cell_pats.append({
            "sheet": sheet, "col": col, "row": row,
            "col_idx": col_idx, "formula": formula, "cell_info": cell_info,
            "pattern_key": (sheet, pkey) if pkey else None,
            "refs": refs,
        })

    # bucket by pattern
    buckets = defaultdict(list)
    no_pattern = []
    for cp in cell_pats:
        if cp["pattern_key"] is None:
            no_pattern.append(cp)
        else:
            buckets[cp["pattern_key"]].append(cp)

    groups = []
    singles = []

    for _pkey, cells in buckets.items():
        if len(cells) < 2:
            cp = cells[0]
            singles.append((cp["sheet"], cp["col"], cp["row"],
                            cp["formula"], cp["cell_info"]))
            continue

        used = set()

        # --- vertical runs (same column, consecutive rows) ---
        col_buckets = defaultdict(list)
        for cp in cells:
            col_buckets[cp["col"]].append(cp)

        for _col, col_cells in col_buckets.items():
            col_cells.sort(key=lambda c: c["row"])
            runs = _find_contiguous_runs(
                col_cells, lambda a, b: b["row"] - a["row"] == 1)
            for run in runs:
                if len(run) >= 2:
                    groups.append({
                        "cells": [(c["sheet"], c["col"], c["row"],
                                   c["formula"], c["cell_info"]) for c in run],
                        "direction": "vertical",
                    })
                    for c in run:
                        used.add((c["sheet"], c["col"], c["row"]))

        # --- horizontal runs (same row, consecutive columns) ---
        remaining = [cp for cp in cells
                     if (cp["sheet"], cp["col"], cp["row"]) not in used]
        row_buckets = defaultdict(list)
        for cp in remaining:
            row_buckets[cp["row"]].append(cp)

        for _row, row_cells in row_buckets.items():
            row_cells.sort(key=lambda c: c["col_idx"])
            runs = _find_contiguous_runs(
                row_cells, lambda a, b: b["col_idx"] - a["col_idx"] == 1)
            for run in runs:
                if len(run) >= 2:
                    groups.append({
                        "cells": [(c["sheet"], c["col"], c["row"],
                                   c["formula"], c["cell_info"]) for c in run],
                        "direction": "horizontal",
                    })
                    for c in run:
                        used.add((c["sheet"], c["col"], c["row"]))

        # leftovers → singles
        for cp in cells:
            if (cp["sheet"], cp["col"], cp["row"]) not in used:
                singles.append((cp["sheet"], cp["col"], cp["row"],
                                cp["formula"], cp["cell_info"]))

    for cp in no_pattern:
        singles.append((cp["sheet"], cp["col"], cp["row"],
                        cp["formula"], cp["cell_info"]))

    return groups, singles


# ---------------------------------------------------------------------------
# Dependency ordering for groups + singles
# ---------------------------------------------------------------------------

def _cells_produced(item):
    """Set of (sheet, col, row) that *item* computes."""
    if isinstance(item, dict):  # group
        return {(s, c, r) for s, c, r, _f, _ci in item["cells"]}
    else:  # single tuple
        return {(item[0], item[1], item[2])}


def _cells_needed(item, tables):
    """Set of (sheet, col, row) that *item* depends on."""
    from formula_converter import FormulaConverter
    needed = set()
    cells_list = item["cells"] if isinstance(item, dict) else [item]
    for sheet, col, row, formula, _ci in cells_list:
        conv = FormulaConverter(sheet, tables)
        conv.convert(formula)
        for s, c, r in conv.referenced_cells:
            needed.add((s, c, r))
        for s, c1, r1, c2, r2 in conv.referenced_ranges:
            ci1 = col_letter_to_index(c1)
            ci2 = col_letter_to_index(c2)
            for rr in range(r1, r2 + 1):
                for cc in range(ci1, ci2 + 1):
                    needed.add((s, index_to_col_letter(cc), rr))
    return needed


def order_items(groups, singles, tables):
    """Topologically sort groups and singles by dependency.

    Returns a list of items (each is either a group dict or a single tuple).
    """
    items = list(groups) + list(singles)
    if not items:
        return items

    # map produced cell → item index
    producer = {}
    for idx, item in enumerate(items):
        for cell in _cells_produced(item):
            producer[cell] = idx

    # build adjacency
    n = len(items)
    in_deg = [0] * n
    fwd = defaultdict(list)

    for idx, item in enumerate(items):
        deps = _cells_needed(item, tables)
        own = _cells_produced(item)
        dep_indices = set()
        for cell in deps - own:
            pi = producer.get(cell)
            if pi is not None and pi != idx:
                dep_indices.add(pi)
        for pi in dep_indices:
            fwd[pi].append(idx)
            in_deg[idx] += 1

    # Kahn's algorithm
    queue = [i for i in range(n) if in_deg[i] == 0]
    ordered = []
    while queue:
        node = queue.pop(0)
        ordered.append(node)
        for nb in fwd[node]:
            in_deg[nb] -= 1
            if in_deg[nb] == 0:
                queue.append(nb)

    # circular leftovers
    remaining = set(range(n)) - set(ordered)
    ordered.extend(remaining)

    return [items[i] for i in ordered]


# ---------------------------------------------------------------------------
# External-file discovery
# ---------------------------------------------------------------------------

def discover_external_files(formula_cells):
    """Return a dict  { filename: set of (sheet, col, row, formula) }."""
    ext_files = defaultdict(set)
    for sheet, col, row, formula, _ci in formula_cells:
        raw = formula[1:] if formula.startswith("=") else formula
        for m in _EXT_RANGE_PATTERN.finditer(raw):
            if not _in_string(raw, m.start()):
                ext_files[m.group(1)].add((sheet, col, row, formula))
        for m in _EXT_CELL_PATTERN.finditer(raw):
            if not _in_string(raw, m.start()):
                ext_files[m.group(1)].add((sheet, col, row, formula))
    return dict(ext_files)


# ---------------------------------------------------------------------------
# Cross-sheet / external analysis helpers
# ---------------------------------------------------------------------------

def analyse_references(formula_cells):
    """Produce analysis data for the report.

    Returns a dict with:
        cross_sheet: list of dicts (sheet, col, row, formula, target_sheet)
        external:    list of dicts (sheet, col, row, formula, ext_file, ext_sheet)
    """
    cross_sheet = []
    external = []
    for sheet, col, row, formula, _ci in formula_cells:
        refs = extract_references(formula, sheet)
        for ref in refs:
            if ref.external_file:
                external.append({
                    "sheet": sheet, "col": col, "row": row,
                    "cell": f"{col}{row}",
                    "formula": formula,
                    "external_file": ref.external_file,
                    "external_sheet": ref.sheet,
                    "ref_type": ref.kind,
                })
            elif ref.sheet != sheet:
                cross_sheet.append({
                    "sheet": sheet, "col": col, "row": row,
                    "cell": f"{col}{row}",
                    "formula": formula,
                    "target_sheet": ref.sheet,
                    "ref_type": ref.kind,
                })
    return {"cross_sheet": cross_sheet, "external": external}
