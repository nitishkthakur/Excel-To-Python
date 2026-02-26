# Excel-to-Mapping: Documentation

This module analyses Excel workbooks and produces an **intermediate mapping report** — a flat, tabular Excel file that classifies every cell as an Input, Calculation, or Output.  A human reviewer can edit this report (e.g. toggle `IncludeFlag`, add rows) and then feed it back to recreate an Excel workbook that includes only the selected items.

---

## Table of Contents

1. [Overview](#overview)
2. [Workflow](#workflow)
3. [Intermediate Excel Format](#intermediate-excel-format)
4. [CLI Reference](#cli-reference)
5. [Python API](#python-api)
6. [How Formula Groups Work](#how-formula-groups-work)
7. [Human Reviewer Guide](#human-reviewer-guide)
8. [Regeneration Logic](#regeneration-logic)
9. [Testing](#testing)

---

## Overview

The module has three main components:

| Component | File | Purpose |
|-----------|------|---------|
| **Mapper** | `mapper.py` | Reads an Excel workbook and generates the tabular mapping report |
| **Regenerator** | `regenerator.py` | Reads the mapping report and reconstructs an Excel workbook |
| **CLI** | `main.py` | Command-line interface for all operations |

---

## Workflow

```
Original Excel (.xlsx)
        │
        ▼
   [1] mapper.py ──────────► Intermediate Mapping Excel
                                      │
                              [2] Human Review
                              (set IncludeFlag, add rows)
                                      │
                                      ▼
                              [3] regenerator.py ──────► Regenerated Excel
                                      │
                              (optional) Input Template
                                      │
                              [4] User fills values ──► regenerator.py with overrides
```

### Step 1: Generate the Mapping Report

```bash
python -m excel_to_mapping.main map Indigo.xlsx --output mapping_report.xlsx
```

This produces a tabular Excel file with one sheet per source sheet plus a `_Metadata` sheet.

### Step 2: Human Review

Open `mapping_report.xlsx` in Excel. For each row, the reviewer can:
- Set `IncludeFlag` to `False` to exclude a cell/group from regeneration
- Add new rows for additional calculations
- Modify formulas or values

### Step 3: Regenerate the Workbook

```bash
python -m excel_to_mapping.main regenerate mapping_report.xlsx --output regenerated.xlsx
```

### Step 4 (Optional): Use Input Overrides

```bash
# Generate an input template
python -m excel_to_mapping.main template mapping_report.xlsx --output input_template.xlsx

# Edit input_template.xlsx with new values, then regenerate
python -m excel_to_mapping.main regenerate mapping_report.xlsx --output regenerated.xlsx --inputs input_template.xlsx
```

---

## Intermediate Excel Format

### Data Sheets (one per source sheet)

Each sheet has a header row followed by one data row per cell or formula group:

| Column | Description | Example |
|--------|-------------|---------|
| **Sheet** | Source sheet name | `Assumptions Sheet` |
| **Cell** | Cell address or range | `A1` or `K2:P2` |
| **Type** | Classification | `Input`, `Calculation`, or `Output` |
| **Formula** | Excel formula | `=J2*(1+K3)` |
| **Value** | Hardcoded value | `2010` |
| **GroupID** | Unique ID for dragged-formula groups | `Assumptions Sheet_G1` |
| **GroupDirection** | Direction of the group | `vertical` or `horizontal` |
| **GroupSize** | Number of cells in the group | `6` |
| **PatternFormula** | Representative formula for the group | `=J2*(1+K3)` |
| **NumberFormat** | Excel number format string | `0`, `#,##0.00`, `0%` |
| **FontBold** | Bold font | `True` / `False` |
| **FontItalic** | Italic font | `True` / `False` |
| **FontSize** | Font size in points | `11` |
| **FontColor** | Font colour (ARGB hex) | `FFFF0000` |
| **FillColor** | Cell fill colour (ARGB hex) | `FF4472C4` |
| **HorizAlign** | Horizontal alignment | `center`, `left`, `right` |
| **VertAlign** | Vertical alignment | `top`, `center`, `bottom` |
| **WrapText** | Text wrapping | `True` / `False` |
| **IncludeFlag** | Include in regeneration | `True` (default) |

### _Metadata Sheet

| Column | Description | Example |
|--------|-------------|---------|
| **SheetName** | Source sheet name | `Assumptions Sheet` |
| **MergedCells** | Semicolon-separated merge ranges | `A1:C1;D5:D8` |
| **ColWidths** | JSON dict of column widths | `{"A": 15.0, "B": 12.5}` |
| **RowHeights** | JSON dict of row heights | `{"1": 27.0, "15": 33.0}` |

---

## CLI Reference

### `map` — Generate Mapping Report

```bash
python -m excel_to_mapping.main map <excel_file> [options]

Options:
  --sheets SHEET1 SHEET2   Limit to specific sheets (default: all)
  --config CONFIG.yaml     Path to config YAML
  --output PATH            Output file path (default: ./output/mapping_report.xlsx)
```

### `template` — Generate Input Template

```bash
python -m excel_to_mapping.main template <mapping_file> [options]

Options:
  --output PATH    Output template path (default: ./output/input_template.xlsx)
```

### `regenerate` — Regenerate Workbook

```bash
python -m excel_to_mapping.main regenerate <mapping_file> [options]

Options:
  --output PATH    Output workbook path (default: ./output/regenerated.xlsx)
  --inputs PATH    Input template with overridden values
```

---

## Python API

### `generate_mapping_report(excel_path, sheet_names=None, config_path=None, output_path=None)`

Generate the tabular mapping report from a source Excel workbook.

**Returns:** Path to the generated report file.

### `regenerate_workbook(mapping_path, output_path, input_values_path=None)`

Regenerate an Excel workbook from the intermediate mapping report.

**Parameters:**
- `mapping_path` — Path to the mapping report
- `output_path` — Where to write the regenerated workbook
- `input_values_path` — Optional input template with overridden values

**Returns:** Path to the regenerated workbook.

### `generate_input_template(mapping_path, output_path)`

Generate an input template with only `Input` rows for users to fill in.

**Returns:** Path to the template file.

---

## How Formula Groups Work

### Dragged Formula Detection

Excel's "drag to fill" feature creates formulas that share the same pattern but with shifted references. For example:

| Cell | Formula |
|------|---------|
| C2 | `=A2-B2` |
| C3 | `=A3-B3` |
| C4 | `=A4-B4` |
| C5 | `=A5-B5` |
| C6 | `=A6-B6` |

These are detected as a single **vertical group** and collapsed to one row:

| Cell | Formula | GroupDirection | GroupSize |
|------|---------|---------------|-----------|
| C2:C6 | `=A2-B2` | vertical | 5 |

### Reference Shifting During Regeneration

When the regenerator expands a group, it reconstructs each cell's formula by shifting references:
- **Vertical groups**: row references shift by offset from base row
- **Horizontal groups**: column references shift by offset from base column
- **Absolute references** (e.g. `$A$1`) are never shifted

### Vectorization Benefits

1. **Compact representation**: Thousands of dragged formulas collapse to a few group rows
2. **Efficient regeneration**: Loops instead of individual writes
3. **Easier review**: Domain experts see patterns, not individual cells
4. **Cleaner debugging**: One formula pattern to verify instead of many

---

## Human Reviewer Guide

### Setting IncludeFlag

1. Open the mapping report in Excel
2. Find the `IncludeFlag` column (column S)
3. Set `False` for any row you want to exclude from the regenerated workbook
4. Save the file

### Adding New Rows

You can add rows to represent additional calculations:
1. Fill in `Sheet`, `Cell`, `Type`, and `Formula` at minimum
2. Set `IncludeFlag` to `True`
3. For groups, also fill in `GroupDirection`, `GroupSize`, and `PatternFormula`

### Modifying Values

- Change `Value` for Input rows to use different hardcoded values
- Change `Formula` for formula rows to use different calculations

---

## Regeneration Logic

### Process

1. **Read metadata**: Merged cells, column widths, row heights
2. **Filter rows**: Only include rows where `IncludeFlag` is truthy
3. **Process each row**:
   - **Input**: Write value to cell (use override from input template if provided)
   - **Single formula**: Write formula directly
   - **Group formula**: Expand group, shift references, write each cell
4. **Apply formatting**: Font, fill, alignment, number format
5. **Apply layout**: Merged cells, column widths, row heights

### Formula Shifting Algorithm

For a group with base formula at position (base_col, base_row):
1. Parse all references in the formula using `extract_references()`
2. For each reference, check if row/column is absolute (`$`)
3. Shift non-absolute parts by the offset from the base position
4. Reconstruct the formula by replacing references right-to-left

---

## Testing

### Run Tests

```bash
python -m pytest tests/test_mapping.py -v
```

### Test Coverage

| Test Category | Count | Description |
|---------------|-------|-------------|
| Classification helpers | 5 | `_build_all_referenced_cells`, `_classify_formula_cells` |
| Sheet row building | 5 | `_build_sheet_rows` unit + integration |
| Report generation | 16 | Simple, vertical drag, multi-sheet, horizontal drag, edge cases, sample workbook |
| Formula shifting | 5 | `_shift_formula` with offsets, absolutes, cross-sheet |
| Group expansion | 2 | `_expand_group` vertical + horizontal |
| Regeneration | 5 | Simple, vertical drag, multi-sheet, cross-sheet |
| Input template | 3 | Creation, content, overrides |
| IncludeFlag | 1 | Exclusion filtering |
| **Total** | **44** | |

### Indigo.xlsx Validation

The full pipeline was tested on `Indigo.xlsx` (13 sheets, 7,872 cells):
- **All 3,938 formulas**: Exact match after regeneration
- **All hardcoded values**: Match (187 float precision differences < 1e-6)
- **All merged cells**: Match across all 13 sheets
- **Overall match rate**: 99.99%
