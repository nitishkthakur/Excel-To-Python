# Excel-to-Python Pipeline — Documentation

## Overview

This project converts Excel financial models into reproducible Python
pipelines.  Given an `.xlsx` or `.xls` workbook the pipeline produces:

1. A **mapping report** describing every cell (type, formula, value, format,
   dependencies).
2. **Input files** (unstructured *and* structured) containing the editable
   values extracted from the original workbook.
3. **Calculation scripts** that read the inputs, evaluate every formula in
   dependency order, apply formatting, and write an `output.xlsx` that
   reproduces the original workbook.

The pipeline supports **22 test workbooks** spanning typical financial-model
patterns (DCF, balance sheets, revenue models, ratios, etc.).  Current
results: **22/22 execute successfully**, **10/22 produce a byte-perfect
match** against the original cached values.

---

## Architecture

```
ExcelFiles/
  ├── ACC-Ltd.xlsx
  ├── capbudg.xls          ← .xls files are auto-converted via LibreOffice
  └── …                       (22 files total)

src/
  ├── excel_utils.py               Shared helpers (column util, .xls → .xlsx)
  ├── formula_parser.py            Cell-ref parsing, R1C1 patterns, extraction
  ├── generate_mapping_report.py   Layer 1  — mapping_report.xlsx
  ├── generate_unstructured_inputs.py  Layer 2a — unstructured_inputs.xlsx
  ├── generate_structured_inputs.py    Layer 2b — structured_input.xlsx
  ├── generate_unstructured_calculate.py  Layer 3a — unstructured_calculate.py
  ├── generate_structured_calculate.py    Layer 3b — structured_calculate.py
  └── compare_utils.py            Cell-by-cell comparison utility

output/<WorkbookStem>/
  ├── mapping_report.xlsx
  ├── unstructured_inputs.xlsx
  ├── structured_input.xlsx
  ├── unstructured_calculate.py
  ├── structured_calculate.py
  ├── output.xlsx                  ← from unstructured pipeline
  └── output_structured.xlsx       ← from structured pipeline
```

### Data-flow diagram

```
 ┌──────────┐     Layer 1      ┌──────────────────┐
 │ Excel    │ ───────────────► │ mapping_report    │
 │ workbook │                  │ .xlsx             │
 └──────────┘                  └────────┬──────────┘
                                        │
              ┌─────────────────────────┼───────────────────────┐
              │                         │                       │
      Layer 2a│                 Layer 2b│               Layer 3a│ / 3b
              ▼                         ▼                       ▼
 ┌─────────────────┐   ┌────────────────────┐   ┌───────────────────────┐
 │ unstructured     │   │ structured_input   │   │ *_calculate.py        │
 │ _inputs.xlsx     │   │ .xlsx              │   │ (generated script)    │
 └────────┬─────────┘   └────────┬───────────┘   └───────────────────────┘
          │                      │                         │
          │   Layer 3a           │   Layer 3b              │
          └──────────► run ◄─────┘                         │
                        │                                  │
                        ▼                                  │
                  ┌────────────┐                           │
                  │ output.xlsx│ ◄─────────────────────────┘
                  └────────────┘
```

---

## Layer Descriptions

### Layer 1 — Mapping Report (`generate_mapping_report.py`)

**Input:** Original Excel workbook  
**Output:** `mapping_report.xlsx`

Reads the workbook twice (with and without `data_only`) and for every
non-empty cell records:

| Column | Description |
|--------|-------------|
| Cell | A1-style reference |
| Row / Col | 1-based integers |
| Type | `Input`, `Calculation`, `Output` |
| Formula | Raw Excel formula (if any) |
| Value | Cached value |
| NumberFormat, Font*, Fill*, Alignment, WrapText | Formatting metadata |
| GroupID / GroupDirection / GroupSize / PatternFormula | Formula-group info |
| IncludeFlag | Always `True` (reserved for filtering) |
| ReferencedBy / References | Cell-level dependency graph |

Special sheets:

* **`_Metadata`** — source file path, timestamp, etc.
* **`_DefinedNames`** — workbook-level named ranges (Name → Reference).
  Filtered: no external-workbook refs, multi-range, function-based, or
  range refs.

Key implementation details:

* Uses `openpyxl` for `.xlsx` and `xlrd` for `.xls` (with automatic
  LibreOffice conversion when needed).
* `DataTableFormula` objects are detected and replaced with the cached
  value via `isinstance` guard.
* Formula groups are detected by comparing R1C1 patterns across adjacent
  cells in the same row or column.

---

### Layer 2a — Unstructured Inputs (`generate_unstructured_inputs.py`)

**Input:** `mapping_report.xlsx`  
**Output:** `unstructured_inputs.xlsx`

Copies every non-formula cell (Input, Label, Header, etc.) into a new
workbook preserving the **original sheet/row/column layout**.  This file
serves as the direct input for `unstructured_calculate.py`.

---

### Layer 2b — Structured Inputs (`generate_structured_inputs.py`)

**Input:** `mapping_report.xlsx`  
**Output:** `structured_input.xlsx`

Reorganises Input cells into user-friendly **tabular sheets**:

| Sheet | Purpose |
|-------|---------|
| **Index** | Cross-reference: SourceSheet → TargetSheet, Cell, Row, Col |
| **Config** | Scalar / short-vector inputs (SourceSheet, Cell, Label, Value) |
| *Data tabs* | One per rectangular input patch, named `SheetName_N` |

*Auto-transpose rule:* When column headers resemble financial periods
(years, quarters), the table is transposed so rows = periods, columns =
metrics — matching typical analyst expectations.

---

### Layer 3a — Unstructured Calculate (`generate_unstructured_calculate.py`)

**Input:** `mapping_report.xlsx`  
**Output:** `unstructured_calculate.py` (standalone generated script)

This is the core code generator.  It reads the mapping report and emits a
self-contained Python script (~1000–20,000 lines depending on model size)
that:

1. Opens `unstructured_inputs.xlsx`.
2. Evaluates every Calculation/Output formula in **topological order**.
3. Applies formatting.
4. Saves `output.xlsx`.

#### Evaluation-order algorithm

1. **Collect** formula cells and build a cell-level dependency graph from
   the `References` column.
2. **Group** cells sharing the same R1C1 pattern in a contiguous run
   (GroupID ≥ 2 members → vectorised emission unit).
3. **Build a unit-level graph** (group↔group, group↔cell, cell↔cell).
4. **Detect SCCs** via iterative Tarjan's algorithm.  Groups inside
   non-trivial SCCs are "broken" back into individual cells.
5. **Topo-sort** the condensed DAG of SCCs (Kahn's algorithm).
6. **Cell-level Kahn sort** within each SCC; non-broken groups are emitted
   as vectorised `for` loops.

#### Formula translation

The translator converts Excel formulas to Python expressions:

* Cell refs → `_g(sheet, row, col)` (get) / `_s(sheet, row, col, val)` (set)
* Range refs → `_rng(sheet, r1, c1, r2, c2)` (flat list) or `_rng2d(…)` (2-D for INDEX)
* **80+ Excel functions** mapped to Python wrappers (`_xl_sum`, `_xl_if`,
  `_xl_vlookup`, `_xl_irr`, `_xl_pmt`, etc.)
* Named ranges resolved via word-boundary regex before other translation.
* Structured table references (`Table[[…]]`, `Table[Col]`) → `None`.
* Error constants (`#REF!`, `#N/A`, etc.) → `None`.
* `%` postfix → `/100`, `^` → `**`, `&` → `+`, `<>` → `!=`.
* Double-comma cleanup: `,,` → `,None,`.
* Final `compile()` safety check; unparseable expressions → `None`.

#### TypeError retry mechanism

Excel treats blank cells as 0 in arithmetic; Python raises `TypeError` on
`None + 5`.  The generated script uses a `_NM` (numeric mode) flag:

```python
_NM = [False]   # mutable list avoids 'global' keyword issues

def _g(sheet, row, col):
    ...
    return 0 if (_NM[0] and v is None) else v
```

Every formula evaluation is wrapped:

```python
try:
    _s(sheet, row, col, <expression>)
except TypeError:
    _NM[0] = True
    try: _s(sheet, row, col, <expression>)
    except Exception: _s(sheet, row, col, None)
    finally: _NM[0] = False
```

---

### Layer 3b — Structured Calculate (`generate_structured_calculate.py`)

**Input:** `mapping_report.xlsx`  
**Output:** `structured_calculate.py` (standalone generated script)

Uses the **same formula-evaluation engine** as Layer 3a (identical
emission plan, identical runtime helpers, identical function wrappers).

The difference is the **input-loading stage**:

1. Creates a fresh workbook with all original sheets.
2. **Loads constants** — non-formula cell values (labels, headers, default
   inputs) are hardcoded in a `_load_constants()` function.
3. **Loads structured inputs** — reads `structured_input.xlsx`:
   * Config sheet → scalar values placed at original coordinates.
   * Data tabs → values mapped back to original (row, col) positions
     using the Index sheet as a lookup.  Transposed tables have their
     row/col mapping reversed.
4. Evaluates formulas (same as Layer 3a).
5. Saves `output.xlsx`.

---

## Comparison Utility (`compare_utils.py`)

Cell-by-cell comparison between `output.xlsx` and the original Excel:

* Opens both workbooks in `data_only` mode.
* Numeric tolerance: `rel_tol=1e-6`, `abs_tol=1e-9`.
* String comparison: stripped whitespace.
* Reports mismatches as `Mismatch(sheet, cell, expected, actual, kind)`.
* Kinds: `missing`, `extra`, `value`, `type`, `missing_sheet`.

---

## Shared Utilities

### `excel_utils.py`

* `ensure_xlsx(path)` — converts `.xls` to `.xlsx` via LibreOffice if needed
  (caches in `ExcelFiles/_converted/`).
* `safe_sheet_name(name)` — truncates / sanitises for openpyxl's 31-char
  sheet-name limit.
* `col_to_num` / `num_to_col` — column letter ↔ number.

### `formula_parser.py`

* `cell_to_rowcol(ref)` / `rowcol_to_cell(row, col)` — A1 ↔ (row, col).
* `extract_references(formula)` — returns list of referenced cells/ranges.
* `to_r1c1_pattern(formula, row, col)` — converts to R1C1 for group
  detection.
* `r1c1_to_a1(pattern, row, col)` — converts back.

---

## Running the Pipeline

### Prerequisites

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# LibreOffice required for .xls → .xlsx conversion
```

### Single-file execution

```bash
# Layer 1
python -m src.generate_mapping_report ExcelFiles/ACC-Ltd.xlsx output/ACC-Ltd/

# Layer 2a
python -m src.generate_unstructured_inputs output/ACC-Ltd/mapping_report.xlsx output/ACC-Ltd/

# Layer 2b
python -m src.generate_structured_inputs output/ACC-Ltd/mapping_report.xlsx output/ACC-Ltd/

# Layer 3a — generate then run
python -m src.generate_unstructured_calculate output/ACC-Ltd/mapping_report.xlsx output/ACC-Ltd/
cd output/ACC-Ltd && python unstructured_calculate.py unstructured_inputs.xlsx output.xlsx

# Layer 3b — generate then run
python -m src.generate_structured_calculate output/ACC-Ltd/mapping_report.xlsx output/ACC-Ltd/
cd output/ACC-Ltd && python structured_calculate.py structured_input.xlsx output_structured.xlsx
```

### Batch execution (all 22 files)

```python
import pathlib, os, subprocess, sys
from src.generate_mapping_report import generate_mapping_report
from src.generate_unstructured_inputs import generate_unstructured_inputs
from src.generate_structured_inputs import generate_structured_inputs
from src.generate_unstructured_calculate import generate_unstructured_calculate
from src.generate_structured_calculate import generate_structured_calculate

for ef in sorted(pathlib.Path("ExcelFiles").glob("*")):
    if not ef.is_file() or ef.suffix not in (".xlsx", ".xls"):
        continue
    out = f"output/{ef.stem}"
    rp = generate_mapping_report(str(ef), out)
    generate_unstructured_inputs(rp, out)
    generate_structured_inputs(rp, out)
    uc = generate_unstructured_calculate(rp, out)
    sc = generate_structured_calculate(rp, out)
    # Execute generated scripts
    subprocess.run([sys.executable, uc,
                    f"{out}/unstructured_inputs.xlsx", f"{out}/output.xlsx"])
    subprocess.run([sys.executable, sc,
                    f"{out}/structured_input.xlsx", f"{out}/output_structured.xlsx"])
```

### Comparison

```python
from src.compare_utils import compare_workbooks, summarise_mismatches

mm = compare_workbooks("ExcelFiles/ACC-Ltd.xlsx", "output/ACC-Ltd/output.xlsx")
print(summarise_mismatches(mm))
```

---

## Test Results Summary

| # | File | Layer 3a | Layer 3b | Mismatches |
|---|------|----------|----------|------------|
| 1 | ACC-Ltd | ✓ | ✓ | 0 |
| 2 | Aurobindo-Pharma | ✗ | ✗ | 4722 |
| 3 | Automation_Justification | ✗ | ✗ | 184 |
| 4 | Bharti-Airtel(2) | ✗ | ✗ | 202 |
| 5 | Bharti-Airtel(3) | ✗ | ✗ | 202 |
| 6 | Financial_model_1 | ✗ | ✗ | 1669 |
| 7 | Gail-India | ✗ | ✗ | 193 |
| 8 | Indigo | ✗ | ✗ | 397 |
| 9 | TF281b | ✗ | ✗ | 204 |
| 10 | TF368 | ✓ | ✓ | 0 |
| 11 | TF7b5 | ✗ | ✗ | 513 |
| 12 | TFb55 | ✗ | ✗ | 204 |
| 13 | TFd14f | ✗ | ✗ | 488 |
| 14 | TFe68 | ✗ | ✗ | 844 |
| 15 | capbudg | ✓ | ✓ | 0 |
| 16 | carlease | ✓ | ✓ | 0 |
| 17 | fcfest | ✓ | ✓ | 0 |
| 18 | fcffst | ✓ | ✓ | 0 |
| 19 | infltn | ✓ | ✓ | 0 |
| 20 | model | ✓ | ✓ | 0 |
| 21 | ratings | ✓ | ✓ | 0 |
| 22 | statmnts | ✓ | ✓ | 0 |

**Totals:** 22/22 execute, 10/22 perfect, 9822 mismatches.

Layer 3a and Layer 3b produce **identical results** — the structured
pipeline is functionally equivalent to the unstructured pipeline.

### Residual mismatch categories

* **Named ranges / table references** — formulas referencing structured
  table syntax (`Table[[#This Row],[Col]]`) or named ranges pointing to
  external workbooks are replaced with `None`, cascading through dependent
  cells.
* **INDIRECT / dynamic references** — `INDIRECT("A"&ROW())` patterns
  cannot be statically resolved.
* **Array formulas / data tables** — Excel's `{=…}` CSE arrays and
  What-If data tables have limited emulation.

---

## Key Design Decisions

1. **SCC-based evaluation order** — Tarjan's algorithm on the
   unit-level dependency graph handles circular references gracefully
   (breaks groups, falls back to cell-level ordering within SCCs).

2. **Two formula translators** — `_formula_to_python` (static) and
   `_formula_to_python_parametric` (loop variable `_row` / `_col`) enable
   both individual cell emission and vectorised group loops.

3. **TypeError retry with `_NM[0]`** — avoids global `_gn()` replacement
   that would break string-valued cells.  Only activates numeric-coercion
   mode on the specific expression that raised `TypeError`.

4. **Non-serialisable value guard** — `DataTableFormula` and other
   openpyxl objects are filtered out during constant emission
   (`isinstance(val, (str, int, float, bool))` check).

5. **Structured input reverse-mapping** — data-tab values are mapped back
   to original cell coordinates using `min_row` / `min_col` offsets
   derived from the Index sheet, with transposition awareness.
