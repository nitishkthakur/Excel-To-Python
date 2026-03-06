# Excel-to-Python Conversion Pipeline — Full Requirements

This document is a complete specification. Given only this file, the entire project should be reproducible from scratch.

---

## Project Goal

Build a Python pipeline that converts Excel financial models (`.xlsx`) into executable Python code. The pipeline must:

1. **Analyse** an Excel workbook — classify every cell, extract formulas, detect formula patterns
2. **Generate** a human-readable input template users can edit
3. **Generate** a standalone Python script that, when run, reads the input template and produces an Excel output that matches the original workbook
4. **Support** a single-sheet mode that targets one sheet while still honouring cross-sheet formula references

The use case is financial modelling: analysts maintain models in Excel; this pipeline extracts the logic so it can be run programmatically (scenario analysis, Monte Carlo, API integration).

---

## Technology Stack

- **Python 3.10+**
- **openpyxl** — Excel I/O (`.xlsx` only; `.xls` is not supported)
- **pandas / numpy** — vectorised calculation in generated code
- **PyYAML** — configuration
- **formulas** (pip package) — reserved for future formula parsing; not yet used
- **pytest** — testing framework
- Virtual environment at `venv/`

---

## Repository Layout

```
ClaudeCode/
├── run_this.py                    # Full-workbook pipeline runner
├── run_this_onesheet.py           # Single-sheet pipeline runner
├── run_pipeline.py                # Low-level CLI (individual layers)
├── config.yaml                    # Runtime configuration
├── requirements.txt
├── CLAUDE.md                      # Guidance for Claude Code
├── prompt.md                      # This file — full rebuild specification
├── ONESHEET.md                    # Single-sheet pipeline user docs
├── STATUS.md                      # Implementation status tracker
├── ExcelFiles/                    # Sample Excel workbooks for testing
│   ├── Indigo.xlsx
│   ├── Bharti-Airtel(2).xlsx
│   ├── ACC-Ltd.xlsx
│   └── ... (other .xlsx and .xls files)
├── excel_pipeline/
│   ├── core/                      # Reusable analysis modules
│   │   ├── dependency_graph.py
│   │   ├── cell_classifier.py
│   │   ├── formula_analyzer.py
│   │   └── excel_io.py
│   ├── layer1/                    # Mapping report generation
│   │   ├── parser.py
│   │   ├── cell_extractor.py
│   │   └── mapping_writer.py
│   ├── layer2/                    # Input file generation
│   │   ├── unstructured_generator.py
│   │   └── structured_generator.py
│   ├── layer3/                    # Runtime calculation
│   │   ├── unstructured_calculator.py
│   │   └── structured_calculator.py
│   ├── layer4a/                   # Python code generation
│   │   ├── __main__.py
│   │   ├── code_generator.py
│   │   ├── mapping_reader.py
│   │   ├── dependency_graph.py
│   │   ├── formula_translator.py
│   │   ├── vectorization_engine.py
│   │   └── code_emitter.py
│   ├── onesheet/                  # Single-sheet pipeline (isolated)
│   │   ├── freezer.py
│   │   └── pipeline.py
│   ├── runtime/
│   │   └── formula_engine.py
│   ├── validation/
│   └── utils/
│       ├── config.py
│       ├── logging_setup.py
│       └── helpers.py
├── tests/
│   ├── unit/
│   ├── integration/
│   ├── performance/
│   └── e2e/
└── output/                        # Generated at runtime
```

---

## config.yaml

```yaml
pipeline:
  input_folder: "ExcelFiles/"
  output_folder: "output/"
  temp_folder: "temp/"

logging:
  level: "INFO"
  file: "pipeline.log"

performance:
  vectorization_threshold: 10   # minimum cells in a dragged-formula group to vectorize
  chunk_size: 10000              # max cells to expand from a range reference

validation:
  tolerance: 1e-9
  check_formatting: true
  check_formulas: true

version: "1.0.0"
```

---

## Core Concepts

### Cell Classification

Every non-empty cell is classified as exactly one of:

| Type | Rule |
|------|------|
| **Input** | Cell has no formula (hardcoded value or empty) |
| **Calculation** | Cell has a formula AND is referenced by at least one other formula |
| **Output** | Cell has a formula AND is NOT referenced by any other formula (terminal node) |

### `mapping_report.xlsx` — The Contract

This file is the **single source of truth** between all pipeline stages. Every downstream layer reads from it exclusively. Users can also edit it manually to control processing.

**Schema — one sheet per source sheet, plus a `_Metadata` sheet:**

Each data sheet has these columns (exact header names matter):

```
RowNum | ColNum | Cell | Type | Formula | Value |
NumberFormat | FontBold | FontItalic | FontSize | FontColor | FillColor |
Alignment | WrapText |
GroupID | GroupDirection | GroupSize | PatternFormula | Vectorizable | IncludeFlag
```

- `Cell` — coordinate string (`"A1"`) or range (`"D5:O5"`) for consolidated groups
- `Type` — `"Input"`, `"Calculation"`, or `"Output"`
- `GroupID` — integer > 0 if part of a dragged-formula group; 0 otherwise
- `GroupDirection` — `"horizontal"` or `"vertical"`
- `PatternFormula` — template with `{col}` or `{row}` placeholder, e.g. `=B{row}*C{row}`
- `Vectorizable` — `True` if `GroupSize >= vectorization_threshold`
- `IncludeFlag` — `True` by default; user sets to `False` to exclude cell

### Cell Store — Runtime Format

The generated Python scripts use a dictionary `c` as the cell store:

```python
c[(sheet_name, column_letter, row_number)] = value
# e.g.
c[('Income statement', 'D', 5)] = 44927
c[('Assumptions Sheet', 'B', 12)] = 0.08
```

Keys are always `(str, str, int)`. This is consistent across `load_inputs`, `calculate`, and `write_output` in the generated script.

Cross-sheet references translate to:
```python
c.get(('Other Sheet', 'A', 1), 0)
```

---

## Pipeline Stages

### Stage 0 — Entry Points

**`run_this.py <excel_file>`**
Runs the full pipeline for all sheets. Outputs to `output/<stem>/`.

**`run_this_onesheet.py <excel_file> "<sheet_name>"`**
Runs the single-sheet pipeline. Outputs to `output/<stem>_<sheet>/`.

**`run_pipeline.py --layer N --input <file> --output <file>`**
Low-level CLI. Supports `--layer 1`, `--log-level DEBUG`.

---

### Layer 1 — Mapping Report Generator (`layer1/parser.py`)

**Input:** original `.xlsx` file
**Output:** `mapping_report.xlsx`
**Entry point:** `generate_mapping_report(input_path, output_path)`

6-step process:
1. Load workbook (`data_only=False` to preserve formulas)
2. Build `DependencyGraph` — parse all formulas, extract cell references, build precedent/dependent maps, detect circular references
3. Run `FormulaAnalyzer` — detect dragged formula groups, compute vectorizability
4. Extract all cell metadata via `CellExtractor` → list of `CellMetadata` dataclasses
5. Annotate cells with group information
6. Write `mapping_report.xlsx` via `MappingWriter`

**`CellMetadata` dataclass** (from `layer1/cell_extractor.py`):
```python
@dataclass
class CellMetadata:
    sheet_name: str
    row_num: int
    col_num: int
    col_letter: str
    cell_coordinate: str
    cell_type: str          # "Input", "Calculation", "Output"
    formula: str
    value: Any
    number_format: str
    font_bold: bool
    font_italic: bool
    font_size: int
    font_color: str
    fill_color: str
    alignment: str
    wrap_text: bool
    group_id: int           # 0 if not in group
    group_direction: str    # "horizontal", "vertical", or ""
    group_size: int
    pattern_formula: str    # e.g. "=B{row}*C{row}"
    is_vectorizable: bool
    include_flag: bool      # default True
```

**`DependencyGraph`** (from `core/dependency_graph.py`):
- Cell IDs use format `"SheetName!A1"` for cross-sheet, `"A1"` for same-sheet
- `build()` parses all formulas using regex; extracts `Sheet!A1`, `'Sheet Name'!A1`, and plain `A1` patterns
- `classify_cell(sheet_name, coordinate)` applies the Input/Calculation/Output rules
- `has_circular_refs()` detects cycles; logs as warnings, does not abort
- `get_stats()` returns dict with total_formula_cells, circular_references, etc.
- Range expansion capped at 10,000 cells (`chunk_size`) to prevent memory explosion

**`FormulaAnalyzer`** (from `core/formula_analyzer.py`):
- Detects groups of structurally identical formulas that were created by dragging
- A "dragged" group: 2+ adjacent cells in the same row (horizontal) or column (vertical) where each formula is identical except for row/col incrementing
- `vectorization_threshold` (default 10): groups of this size or larger get `is_vectorizable=True`
- Pattern formulas replace concrete references with `{row}` or `{col}` placeholders
- `FormulaGroup` dataclass: `group_id`, `direction`, `cells` (list of coords), `pattern`, `size`, `sheet_name`, `is_vectorizable`

**`MappingWriter`** (from `layer1/mapping_writer.py`):
- Creates one sheet per source sheet (named identically)
- Creates a `_Metadata` sheet with aggregate statistics
- Header row is frozen and bold
- Consolidated groups write one row with the range in `Cell` column (e.g. `"D5:O5"`) and the pattern formula
- All 20 column headers must be present in exact order for downstream layers to parse correctly

---

### Layer 2a — Unstructured Input Generator (`layer2/unstructured_generator.py`)

**Input:** `mapping_report.xlsx`
**Output:** `unstructured_inputs.xlsx`
**Entry point:** `generate_unstructured_inputs(mapping_report_path, output_path)`

- Reads mapping report; filters rows where `Type == "Input"` AND `IncludeFlag == True`
- Creates a new workbook with one sheet per source sheet
- Writes each Input cell to its **original position** (same row, same column)
- Preserves: value, number format, font (bold, italic, size, color), fill color, alignment, wrap_text
- Color strings in the mapping report are hex (with or without `#`); must be converted to 8-char ARGB (prepend `FF` to 6-char RGB)
- Does NOT include Calculation or Output cells
- The result is a clean editable template users can modify with new input values

---

### Layer 3 — Runtime Formula Engine (`runtime/formula_engine.py`)

**Input:** `unstructured_inputs.xlsx` + `mapping_report.xlsx`
**Output:** dict `{(sheet_name, coordinate): value}`
**Entry point:** `calculate_workbook(input_path, mapping_path)`

- Loads input values into `cell_values` dict keyed by `(sheet_name, coordinate)`
- Loads formula metadata (Calculation + Output cells) from mapping report
- Builds dependency graph and topological order (Kahn's algorithm)
- Evaluates formulas in dependency order using a simple `eval()`-based engine
- Groups flagged as vectorizable use `_evaluate_group_vectorized()` (currently falls back to cell-by-cell; true numpy vectorisation is TODO)
- Partial: complex Excel functions not fully implemented; use Layer 4a code generation for production use

---

### Layer 4a — Python Code Generator (`layer4a/code_generator.py`)

**Input:** `mapping_report.xlsx` + `unstructured_inputs.xlsx` (for reference)
**Output:** `unstructured_calculate.py` (standalone Python script)
**Entry point:** `generate_unstructured_code(mapping_report_path, unstructured_inputs_path, output_script_path)`
**Module CLI:** `python -m excel_pipeline.layer4a <mapping_report> <inputs> [output_script]`

5-step process:
1. Read mapping report via `MappingReader` → `cells_by_sheet` dict + vectorizable groups
2. Build `DependencyGraph` (layer4a local version) for calculation order
3. Initialise `FormulaTranslator`, `VectorizationEngine`, `CodeEmitter`
4. Generate calculation code: vectorizable groups → pandas code; non-vectorizable groups → `for col/row in [...]` loop; individual cells → single assignment
5. Assemble and write final Python script

**Generated script structure:**
```python
#!/usr/bin/env python3
"""Auto-generated by Layer 4a..."""

import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
...
import math

# ===== HELPER FUNCTIONS =====
def xl_sum(*args): ...
def xl_average(*args): ...
def xl_eomonth(start_date, months): ...
# ... (all helper functions embedded inline)

def load_inputs(input_path):
    c = {}  # (sheet_name, column_letter, row_int) -> value
    wb_input = load_workbook(input_path, data_only=True)
    for sheet in wb_input.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    c[(sheet.title, cell.column_letter, cell.row)] = cell.value
    return c

def calculate(c):
    # Vectorized group N: M cells (horizontal/vertical)
    for col in ['D','E',...]:
        c[('Sheet', col, 5)] = xl_eomonth(c.get(('Sheet', col, 5), 0), 12)
    ...
    # Individual cells
    c[('Sheet', 'B', 10)] = c.get(('Sheet', 'A', 10), 0) * c.get(('Sheet', 'A', 5), 0)
    ...

def write_output(c, output_path):
    wb_out = Workbook()
    # ... groups cells by sheet, writes all values, saves

def main():
    c = load_inputs("unstructured_inputs.xlsx")
    calculate(c)
    write_output(c, "output.xlsx")

if __name__ == "__main__":
    main()
```

The generated script is standalone. Run it from the directory containing `unstructured_inputs.xlsx`.

---

#### `formula_translator.py` — Excel → Python Translation

**`_translate_cell_by_cell(formula, sheet)`** — the core translation path:

1. **Pre-expand ranges** (call `_expand_ranges` first): `D7:D8` → `D7, D8`
   This must happen before cell-reference substitution to avoid `xl_sum(c.get(...):c.get(...))` invalid syntax.

2. **Cross-sheet references** (regex, apply before same-sheet):
   - `'Sheet Name'!A1` → `c.get(('Sheet Name', 'A', 1), 0)`
   - `SheetName!A1` → `c.get(('SheetName', 'A', 1), 0)`

3. **Same-sheet references** (regex `\b([A-Z]+)(\d+)\b`):
   - `A1` → `c.get(('SheetName', 'A', 1), 0)`

4. **Replace Excel operators**: `^` → `**`

5. **Replace Excel functions** using `EXCEL_FUNCTIONS` dict (simple string replace `FUNC(` → `xl_func(`).
   `IF` is handled separately (see below).

6. **Translate IF** using `_translate_if_function`:
   `IF(cond, true_val, false_val)` → `(true_val if cond else false_val)`
   **CRITICAL:** must use a parenthesis-aware argument splitter (`_split_args`), NOT a simple `([^,]+)` regex — the regex breaks when arguments contain `c.get(('Sheet', 'A', 1), 0)` which contains commas.

**`_expand_ranges(formula, sheet)`**:
```python
# Replaces A1:B3 with comma-separated individual cell references
re.sub(r'\b([A-Z]+)(\d+):([A-Z]+)(\d+)\b', expand_fn, formula)
# e.g. D7:D8 → D7, D8
#      D7:F9 → D7, D8, D9, E7, E8, E9, F7, F8, F9
```

**`EXCEL_FUNCTIONS` mapping** (all must be present):
```python
EXCEL_FUNCTIONS = {
    'SUM': 'xl_sum', 'AVERAGE': 'xl_average', 'COUNT': 'xl_count',
    'MIN': 'min', 'MAX': 'max', 'ABS': 'abs', 'ROUND': 'round',
    'IF': None,  # handled by _translate_if_function
    'AND': 'all', 'OR': 'any', 'NOT': 'not',
    'EOMONTH': 'xl_eomonth', 'EDATE': 'xl_edate',
    'YEAR': 'xl_year', 'MONTH': 'xl_month', 'DAY': 'xl_day',
    'DATE': 'xl_date', 'TODAY': 'xl_today', 'DAYS': 'xl_days',
    'LEN': 'len', 'IFERROR': 'xl_iferror',
    'SUMIF': 'xl_sumif', 'SUMPRODUCT': 'xl_sumproduct',
    'COUNTA': 'xl_counta', 'ISNUMBER': 'xl_isnumber',
    'ISBLANK': 'xl_isblank', 'CONCATENATE': 'xl_concatenate',
}
```

---

#### `vectorization_engine.py` — Code Generation for Groups

**Vertical groups** (same column, `{row}` placeholder):
- Build a pandas DataFrame with one column per dependency
- Apply the vectorised pandas expression
- Write back to `c` via a `for r in df.index:` loop

**Horizontal groups** (same row, `{col}` placeholder):
- Generate `for col in ['D', 'E', ..., 'O']:` loop
- **CRITICAL:** Do NOT use string concatenation to substitute `{col}`. Instead: replace `{col}` with the first column letter (e.g. `'D'`), translate the full formula, then replace `'D'` (the string literal) with the variable name `col`. This avoids generating broken Python like `YEAR(' + col + '5)`.

**Non-vectorizable groups** (loop fallback):
- Same pattern as horizontal groups above (translate with first col/row, then replace with loop variable)

---

#### Helper Functions Embedded in Generated Script

All these must be defined in `HELPER_FUNCTIONS_CODE` (a string constant in `formula_translator.py`) and embedded verbatim into every generated script:

```python
xl_sum(*args)             # flat-list aware SUM
xl_average(*args)         # AVERAGE
xl_count(*args)           # COUNT (numeric only)
xl_if(condition, t, f)    # IF
xl_iferror(value, alt)    # IFERROR
xl_sumif(rng, crit, srng) # SUMIF (simplified)
xl_sumproduct(*arrays)    # SUMPRODUCT
xl_counta(*args)          # COUNTA
xl_isnumber(v)            # ISNUMBER
xl_isblank(v)             # ISBLANK
xl_concatenate(*args)     # CONCATENATE

# Date functions (all handle Excel serial numbers via _to_date helper)
xl_eomonth(start_date, months)   # last day of month N months out
xl_edate(start_date, months)     # same day N months out
xl_year(v), xl_month(v), xl_day(v)
xl_date(year, month, day)
xl_today()
xl_days(end_date, start_date)
```

The `_to_date` helper converts Excel serial numbers (integers) to Python `date` objects using the Excel epoch `date(1899, 12, 30)`.

---

### Single-Sheet Pipeline (`excel_pipeline/onesheet/`)

This module is **completely isolated** — it does not modify any other layer's code. It calls the existing layer entry points as black boxes.

#### Problem

Naively stripping all sheets except the target silently breaks cross-sheet formulas (`='Assumptions Sheet'!B5` resolves to 0 instead of the real value).

#### Solution: Workbook Freezing

**`freezer.py` — `freeze_other_sheets(input_path, target_sheet, frozen_path)`:**

1. Load workbook twice:
   - `wb_f = load_workbook(data_only=False)` — preserves formula strings and special formula objects
   - `wb_v = load_workbook(data_only=True)` — loads last-cached values
2. For every sheet that is NOT the target sheet:
   - Iterate all cells
   - Skip `MergedCell` objects (they have no writable `.value`)
   - Replace any formula cell with its cached value. A formula cell is identified by EITHER:
     - `isinstance(cell.value, str) and cell.value.startswith('=')`
     - `cell.value is not None and not isinstance(cell.value, _SIMPLE_TYPES)`  ← catches `DataTableFormula` / `ArrayFormula` objects which openpyxl returns for data-table cells when `data_only=False`
   - The cached replacement value: `ws_v[cell.coordinate].value`. If this is also not a plain type, use `None`.
3. Save `wb_f` to `frozen_path`
4. Return stats dict: `{other_sheets, frozen_cells, null_cached_values}`

`_SIMPLE_TYPES = (int, float, str, bool, datetime.datetime, datetime.date, type(None))`

**Why this works end-to-end:**
- The target sheet retains all formulas → Layer 1 classifies its cells as Input/Calculation/Output
- Other sheets have only plain values → Layer 1 classifies ALL their cells as Input
- Layer 2a includes other-sheet Input cells in `unstructured_inputs.xlsx`
- The generated script loads everything into `c`; cross-sheet refs like `c.get(('Assumptions Sheet', 'B', 5), 0)` resolve to the frozen cached values

**`pipeline.py` — `run(input_path, sheet_name, output_dir, log_level)`:**

```
Step 1: freeze_other_sheets → intermediate/frozen_workbook.xlsx
Step 2: generate_mapping_report(frozen_path, ...)
Step 3: generate_unstructured_inputs(mapping_path, ...)
Step 4: generate_unstructured_code(mapping_path, inputs_path, script_path)
Step 5: subprocess.run([sys.executable, "unstructured_calculate.py"], cwd=output_dir)
```

The generated script is run via subprocess from `output_dir` so that the hardcoded relative paths `"unstructured_inputs.xlsx"` and `"output.xlsx"` resolve correctly.

---

### Utils

**`config.py`** — Singleton `Config` class. Load via `config.load("config.yaml")`. Properties: `input_folder`, `output_folder`, `log_level`, `log_file`, `vectorization_threshold`, `chunk_size`, `validation_tolerance`, `check_formatting`, `check_formulas`, `version`. Global instance: `from excel_pipeline.utils.config import config`.

**`logging_setup.py`** — `setup_logging(level, log_file)` configures root logger + file handler. `get_logger(__name__)` is used in every module. Log format: `YYYY-MM-DD HH:MM:SS [LEVEL] module.name: message`.

**`excel_io.py`** — `load_workbook(filepath, data_only, read_only)`: wraps openpyxl with logging; `save_workbook(wb, filepath, atomic=True)`: writes to temp file then renames for safety.

---

## Output Directory Layout

```
output/
  <stem>/                          # run_this.py full-workbook output
    mapping_report.xlsx
    unstructured_inputs.xlsx
    unstructured_calculate.py
    output.xlsx
    pipeline.log

  <stem>_<sheet>/                  # run_this_onesheet.py output
    intermediate/
      frozen_workbook.xlsx         # inspect to verify cross-sheet values
    mapping_report.xlsx
    unstructured_inputs.xlsx
    unstructured_calculate.py
    output.xlsx
    pipeline.log
```

---

## Known Pitfalls (Must Avoid When Rebuilding)

1. **Horizontal `{col}` substitution**: never use string concatenation `formula.replace('{col}', "' + col + '")`. This produces `YEAR(' + col + '5)` which is invalid Python. Correct approach: substitute the first concrete column letter, translate the complete formula, then replace that column letter string with the variable name `col`.

2. **Range references in formulas**: `D7:D8` must be pre-expanded to `D7, D8` (comma-separated) BEFORE the cell-reference substitution regex runs. Otherwise `xl_sum(D7:D8)` → `xl_sum(c.get(...):c.get(...))` which is invalid Python (colon is not valid between function call expressions as a Python slice outside `[]`).

3. **IF() argument parsing**: the `IF(cond, true, false)` translation must use a parenthesis-depth-aware splitter, not `([^,]+)` regex. The condition and value arguments frequently contain `c.get(('Sheet', 'COL', ROW), 0)` which has nested commas.

4. **`DataTableFormula` / `ArrayFormula` objects**: openpyxl returns these special objects (not strings) as `cell.value` when loading with `data_only=False` for data-table cells. If not handled they appear as `<openpyxl.worksheet.formula.DataTableFormula object at 0x...>` in generated Python. Detect them with `not isinstance(cell.value, _SIMPLE_TYPES)` and replace with `None` or the cached plain value.

5. **MergedCell proxies**: `MergedCell` objects in openpyxl cannot have their `.value` written. Always check `isinstance(cell, MergedCell)` and skip.

6. **Generated script runs from `output_dir`**: `unstructured_calculate.py` hardcodes `"unstructured_inputs.xlsx"` and `"output.xlsx"` as relative paths. Run it via `subprocess.run([...], cwd=output_dir)`.

7. **`data_only=True` gives stale/None values** when the workbook was never opened in Excel after formula entry. The `null_cached_values` stat exposes this. Cross-sheet references to such cells resolve to `None`/`0`.

---

## Test Commands

```bash
# Install
pip install -r requirements.txt

# Full-workbook pipeline
python run_this.py ExcelFiles/Indigo.xlsx
# → output/Indigo/output.xlsx

# Single-sheet pipeline
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Income statement"
# → output/Indigo_Income_statement/output.xlsx  (~29s on Indigo.xlsx)

python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Balance sheet"
python run_this_onesheet.py ExcelFiles/ACC-Ltd.xlsx "P&L"

# List sheets in a workbook
python -c "import openpyxl; wb=openpyxl.load_workbook('ExcelFiles/Indigo.xlsx',read_only=True); print(wb.sheetnames)"

# Re-run a generated script (after editing unstructured_inputs.xlsx)
cd output/Indigo_Income_statement/
python unstructured_calculate.py

# Layer 1 only
python run_pipeline.py --layer 1 --input ExcelFiles/Indigo.xlsx --output output/test_mapping.xlsx --log-level DEBUG
```

---

## Implementation Status at Last Update

| Component | Status |
|-----------|--------|
| Layer 1 (mapping report) | Complete, tested |
| Layer 2a (unstructured inputs) | Complete |
| Layer 2b (structured inputs) | Partial |
| Layer 3 (runtime formula engine) | Partial (basic eval; complex functions TODO) |
| Layer 4a (Python code generation) | Complete, tested end-to-end |
| `onesheet/` module | Complete, tested (`Indigo.xlsx` → 29s) |
| Validation framework | Structure only; no tests written |
| Layer 4b (structured code gen) | Not started |

---

## What "Complete" Means for the Full Pipeline

The pipeline is considered fully complete when:
1. `run_this.py <any .xlsx in ExcelFiles/>` runs without error
2. `output.xlsx` contains all sheets with values matching what Excel would calculate
3. `run_this_onesheet.py <file> <sheet>` runs in under 60s for a 10-sheet workbook
4. Numerical values in `output.xlsx` match the original within `validation.tolerance` (1e-9)
5. Tests exist for: cell classification, range expansion, IF translation, cross-sheet reference resolution, freeze-other-sheets logic
