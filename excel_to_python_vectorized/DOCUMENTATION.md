# Excel-to-Python Vectorised Converter

> Converts an Excel workbook into a **vectorised** Python script that is smaller,
> faster, and easier to read than the cell-by-cell approach.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [What It Does](#what-it-does)
3. [How Vectorisation Works](#how-vectorisation-works)
4. [External Workbook References](#external-workbook-references)
5. [Analysis Report](#analysis-report)
6. [Generated Files](#generated-files)
7. [Algorithm Overview](#algorithm-overview)
8. [Configuration](#configuration)
9. [Comparison with Non-Vectorised Version](#comparison-with-non-vectorised-version)

---

## Quick Start

```bash
# From the repository root
python excel_to_python_vectorized/main.py path/to/workbook.xlsx --output-dir output

# The command produces:
#   output/calculate.py           – vectorised calculation script
#   output/input_template.xlsx    – template to fill in input values
#   output/analysis_report.xlsx   – reference analysis report
#   output/input_files_config.json – (only if external file refs exist)

# Fill in the template, then run:
python output/calculate.py output/input_template.xlsx result.xlsx
```

---

## What It Does

| Step | Description |
|------|-------------|
| **Parse** | Reads every cell, formula, format, merge, table, and column width from the workbook. |
| **Classify** | Separates cells into *hardcoded values* (inputs) and *formula cells* (computed). |
| **Detect patterns** | Normalises each formula by replacing cell references with relative offsets to detect "dragged" copies. |
| **Group** | Groups formulas that share the same pattern into contiguous vertical or horizontal runs. |
| **Order** | Topologically sorts groups and individual formulas by their dependencies. |
| **Generate** | Emits a Python script that stores every cell in a dictionary `c[(sheet, col, row)]` and uses compact `for` loops for vectorised groups. |
| **Template** | Creates an input Excel file pre-filled with current hardcoded values. |
| **Report** | Produces an Excel report with cross-sheet and external-file reference analysis. |

---

## How Vectorisation Works

### The Problem

In a typical Excel workbook, formulas are often *dragged* down a column or
across a row.  For example:

| Cell | Formula |
|------|---------|
| D2   | `=B2*C2` |
| D3   | `=B3*C3` |
| D4   | `=B4*C4` |

The non-vectorised converter generates **one Python statement per cell**:

```python
s_Inputs_D2 = s_Inputs_B2 * s_Inputs_C2
s_Inputs_D3 = s_Inputs_B3 * s_Inputs_C3
s_Inputs_D4 = s_Inputs_B4 * s_Inputs_C4
```

For large workbooks with hundreds of dragged rows, this becomes enormous.

### The Solution

The vectorised converter detects that all three formulas share the same
**pattern** (each cell references two cells at the same relative offsets)
and emits a single loop:

```python
# Vectorised: Inputs!D2:D4
for _r in range(2, 5):
    try:
        c[('Inputs', 'D', _r)] = c.get(('Inputs', 'B', _r)) * c.get(('Inputs', 'C', _r))
    except Exception:
        c[('Inputs', 'D', _r)] = None
```

### Pattern Normalisation

For each formula the converter:

1. **Extracts** every cell/range reference, noting whether the column and
   row are **absolute** (`$A$1`) or **relative** (`A1`).
2. **Computes offsets** – for relative parts, the offset from the formula
   cell's own position; for absolute parts, the literal value.
3. **Builds a skeleton** – the formula text with references replaced by
   numbered placeholders (`@0`, `@1`, …) plus offset tuples.

Two formulas produce the **same pattern** if their skeleton and offset
tuples are identical.  This correctly handles:

- Fully relative references (`B2*C2` at D2 → `B3*C3` at D3)
- Mixed references (`$A$1*B2` at D2 → `$A$1*B3` at D3)
- Cross-sheet references (`Sheet2!B2` at D2 → `Sheet2!B3` at D3)
- Range references (`SUM(A2:C2)` at D2 → `SUM(A3:C3)` at D3)

### Grouping

Formulas with the same pattern are then checked for **contiguity**:

- **Vertical group**: same column, consecutive rows → `for _r in …`
- **Horizontal group**: same row, consecutive columns → `for _ci in …`

Groups of 2+ cells become loops; singletons remain as individual statements.

---

## External Workbook References

Excel formulas can reference other files:

```
=[ExtData.xlsx]Prices!A1 * 2
```

The converter handles these by:

1. **Detecting** all `[filename]` references in formulas.
2. **Generating `input_files_config.json`** with every external filename as a key:
   ```json
   {
     "ExtData.xlsx": "",
     "Rates.xlsx": ""
   }
   ```
3. The user fills in the **real paths** on their system:
   ```json
   {
     "ExtData.xlsx": "/data/shared/ExtData.xlsx",
     "Rates.xlsx": "C:\\Users\\me\\Rates.xlsx"
   }
   ```
4. The generated script reads these paths at runtime, opens the external
   workbooks, and loads their cell values into the same dictionary `c`
   under keys like `('ExtData.xlsx|Prices', 'A', 1)`.

---

## Analysis Report

The converter always produces `analysis_report.xlsx` with these sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | Total formulas, vectorised vs individual counts, cross-sheet / external counts. |
| **Vectorised Groups** | Every detected group: direction, start/end cells, cell count, representative formula. |
| **Cross-Sheet Refs** | Every formula that references another sheet: source cell, target sheet, formula text. |
| **External Refs** | Every formula that references an external file: source cell, external file and sheet. |
| **Per-Sheet Breakdown** | Per-sheet counts of formulas, vectorised, individual, cross-sheet, and external references. |

---

## Generated Files

| File | Always? | Description |
|------|---------|-------------|
| `calculate.py` | ✅ | Vectorised Python calculation script. |
| `input_template.xlsx` | ✅ | Pre-filled template for user inputs. |
| `analysis_report.xlsx` | ✅ | Reference analysis report. |
| `input_files_config.json` | ❌ | Only generated when external file references are present. |

---

## Algorithm Overview

```
Excel File (.xlsx)
        │
        ▼
┌───────────────┐
│  Parse cells,  │
│  formulas,     │
│  formatting    │
└───────┬───────┘
        │
        ▼
┌───────────────┐
│  Classify:     │
│  formula vs    │
│  hardcoded     │
└───────┬───────┘
        │
        ▼
┌───────────────────┐
│  Normalise each    │
│  formula → pattern │
│  (offsets + skel.) │
└───────┬───────────┘
        │
        ▼
┌───────────────────┐
│  Group by pattern: │
│  vertical /        │
│  horizontal runs   │
└───────┬───────────┘
        │
        ▼
┌───────────────────┐
│  Topological sort  │
│  groups + singles  │
│  by dependency     │
└───────┬───────────┘
        │
        ▼
┌───────────────┐     ┌──────────────────┐     ┌───────────────────┐
│  calculate.py  │     │ input_template   │     │ analysis_report   │
│  (vectorised)  │     │ .xlsx            │     │ .xlsx             │
└────────────────┘     └──────────────────┘     └───────────────────┘
```

---

## Configuration

The converter accepts the same `config.yaml` as the non-vectorised version:

```yaml
# When true, hardcoded values not referenced by any formula are excluded.
delete_unreferenced_hardcoded_values: false
```

Pass it with `--config`:

```bash
python excel_to_python_vectorized/main.py workbook.xlsx --config config.yaml
```

---

## Comparison with Non-Vectorised Version

| Aspect | `excel_to_python.py` | `excel_to_python_vectorized/` |
|--------|---------------------|-------------------------------|
| Cell storage | Individual variables (`s_Sheet_A1`) | Dictionary (`c[(sheet, col, row)]`) |
| Dragged formulas | One line per cell | Single `for` loop |
| Code size | Grows linearly with cell count | Grows with *unique pattern* count |
| Range building | Builder functions per range | Dynamic `_rng()` helper |
| External file refs | ❌ Not supported | ✅ Via `input_files_config.json` |
| Analysis report | ❌ Not generated | ✅ Always generated |
| Formatting preservation | ✅ | ✅ |
| Cross-sheet refs | ✅ | ✅ |
