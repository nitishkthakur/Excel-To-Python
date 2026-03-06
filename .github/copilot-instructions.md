# GitHub Copilot Instructions — Excel-to-Python

This file guides AI coding agents working in this repository.

---

## Mandatory Documentation Updates

**Every time any change is made to this codebase — new features, bug fixes, refactors, new modules, CLI changes, schema changes, or behaviour changes — the agent MUST update both documentation files before finishing the task:**

| File | What to update |
|------|---------------|
| `.github/copilot-instructions.md` | Update the affected section(s): module table, CLI commands, architecture diagram, schema definitions, code-style conventions, or key function tables. Add a new section if the change introduces a new concept not yet covered. |
| `.github/objective-instructions.md` | Update the affected stage description, design principles, or financial-date / transpose rules. If the change affects the user-facing workflow or the shape of any output file, describe the new behaviour explicitly. |

**Rules:**
- Do not defer documentation to a follow-up step — update both files in the same response as the code change.
- Be specific: name the functions, files, columns, or CLI flags that changed.
- If a section already covers the topic, edit it in place. If no section covers it, add one.
- Never leave either file in a state that contradicts the current code.

---

## Project Overview

This repository contains multiple subsystems for analysing, converting, and regenerating Excel workbooks programmatically.

| Subsystem | Entry point | Purpose |
|-----------|------------|---------|
| **Cell-by-cell converter** | `excel_to_python.py` | Parse + classify cells; generate `calculate.py` + `input_template.xlsx` |
| **Vectorised converter** | `excel_to_python_vectorized/main.py` | Same as above but groups dragged formulas into compact `for` loops |
| **Mapper + Regenerator** | `excel_to_mapping/main.py` | Produce a tabular `mapping_report.xlsx`; regenerate an Excel from it |
| **Structured input generator** | `excel_to_mapping/structured_input_generator.py` | Produce a user-friendly `structured_input.xlsx` from the mapping report |
| **Lineage** | `lineage/lineage_builder.py` | Extract data-flow lineage; render dependency graphs |
| **MCP server** | `mcp_server/server.py` | MCP server that exposes Excel content (formulas + values) to LLMs |

---

## Python Environment

Always use the project virtualenv:

```bash
# Activate implicitly by calling the venv python directly:
.venv/bin/python -m <module>

# Or activate first:
source .venv/bin/activate
python -m <module>
```

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Running Tests

```bash
# All tests
python -m pytest tests/ -v

# Single test module
python -m pytest tests/test_mapping.py -v
python -m pytest tests/test_structured_input.py -v
```

Test files live in `tests/`. Each test module targets one subsystem.

---

## Key CLI Commands

```bash
# Generate mapping report from an Excel file
python -m excel_to_mapping.main map Indigo.xlsx --output output/mapping_report.xlsx

# Generate structured input Excel (requires mapping report + original Excel)
python -m excel_to_mapping.main structured-input \
    --excel Indigo.xlsx --output output/structured_input.xlsx

# Regenerate an Excel workbook from the mapping report
python -m excel_to_mapping.main regenerate output/mapping_report.xlsx \
    --output output/regenerated.xlsx

# Generate input template from a mapping report
python -m excel_to_mapping.main template output/mapping_report.xlsx

# Vectorised converter
python excel_to_python_vectorized/main.py Indigo.xlsx --output-dir output

# Cell-by-cell converter
python excel_to_python.py Indigo.xlsx --config config.yaml --output-dir output

# Start MCP server
cd mcp_server && python server.py
```

---

## Architecture: Core Data Flow

```
Original Excel (.xlsx)
        │
        ▼
 excel_to_python.parse_workbook()          → sheets dict, formula_cells list
 excel_to_python.classify_cells()          → hardcoded_cells list
        │
        ▼
 excel_to_python_vectorized.vectorizer.group_formulas()
                                           → groups (patterns), singles
        │
        ▼
 excel_to_mapping.mapper.generate_mapping_report()
                                           → mapping_report.xlsx  [Layer 1]
        │
        ├─────────────────────────────────────────────┐
        ▼  (Path a)                                   ▼  (Path b)
 unstructured_inputs.xlsx  [Layer 2a]    structured_input.xlsx  [Layer 2b]
        │                                             │
        ▼  [Layer 3a — Engine Creator]                ▼  [Layer 3b — Engine Creator]
 python_engine_creator_unstructured.py    python_engine_creator_structured.py
   reads mapping_report.xlsx               reads mapping_report.xlsx +
   ↓ generates                             structured_input.xlsx
 unstructured_calculate.py                 ↓ generates
   pure Python calculations               structured_calculate.py
   reads unstructured_inputs.xlsx          pure Python calculations
        │                                  reads structured_input.xlsx
        └──────────────────┬───────────────────────────┘
                           ▼
                      output.xlsx
              (matches original workbook)
```

> **Python is the only execution engine.** The generated `calculate.py` files contain real Python
> arithmetic — no formula strings are passed to Excel, no COM automation, no openpyxl formula
> evaluation. The input template files (`unstructured_inputs.xlsx` / `structured_input.xlsx`)
> are data stores only; all calculations happen in Python.

### Layer 3 — Two-Stage Code Generation Pattern

Layer 3 follows a **generate-then-run** pattern:

| Stage | Script | Role |
|-------|--------|------|
| Code generation (run once per workbook) | `python_engine_creator_unstructured.py` or `python_engine_creator_structured.py` | Reads `mapping_report.xlsx`; writes a bespoke `calculate.py` containing all formulas translated to Python |
| Runtime (run each time inputs change) | `unstructured_calculate.py` or `structured_calculate.py` | Reads the input template; executes Python calculations; writes `output.xlsx` |

The `calculate.py` files are **generated artifacts** — they are workbook-specific and must never be hand-written. They must not call any Excel API or evaluate any formula string.

---

## Cell Classification

Three cell types are used throughout the codebase:

| Type | Meaning | Colour (in mapper) |
|------|---------|-------------------|
| `Input` | Plain value cell (not a formula) | Green |
| `Calculation` | Formula cell referenced by at least one other formula | Yellow |
| `Output` | Formula cell *not* referenced by any other formula (terminal) | Blue |

The classification logic lives in `excel_to_mapping/mapper.py` → `_classify_formula_cells()`.

---

## Mapping Report Schema

`mapping_report.xlsx` has one sheet per source sheet plus a `_Metadata` sheet.
Each data sheet has exactly 19 columns defined by `COLUMNS` in `mapper.py`:

```python
COLUMNS = [
    "Sheet", "Cell", "Type", "Formula", "Value",
    "GroupID", "GroupDirection", "GroupSize", "PatternFormula",
    "NumberFormat", "FontBold", "FontItalic", "FontSize", "FontColor",
    "FillColor", "HorizAlign", "VertAlign", "WrapText", "IncludeFlag",
]
```

Never reorder or rename these columns — the regenerator reads them by name.

---

## Important openpyxl Pattern: Storing Formulas as Text

When writing a formula string (e.g. `=SUM(A1:A5)`) to a cell **without** Excel evaluating it, bypass openpyxl's type detection:

```python
# CORRECT — stores formula as plain text string
cell._value = formula_string   # use private attribute
cell.data_type = "s"           # mark as string, not formula

# WRONG — openpyxl auto-detects '=' prefix and marks it as a formula
cell.value = formula_string
```

This pattern is used in `mapper.py` to prevent the Formula column from being executable.

---

## Structured Input Generator

`excel_to_mapping/structured_input_generator.py` reads `mapping_report.xlsx` + the original Excel and produces a user-facing input file.

**Sheet layout in `structured_input.xlsx`:**

| Sheet | Content |
|-------|---------|
| `Index` | Cross-reference: Input file sheet → table → source range → vector length |
| `Config` | Scalar inputs (single cells) and short (label + 1 value) vectors |
| `<SourceSheet>` | One sheet per source sheet containing ≥1 vector input; row 1 = headers, col A = metric labels, remaining cols = time-series data |

**Key functions:**

| Function | Purpose |
|----------|---------|
| `_group_into_vectors_and_scalars(inputs)` | Split Input cells into contiguous horizontal runs (≥2 = vector) vs. singles |
| `_split_label_from_vector(vec)` | If the first cell of a vector is a string, it is the row label, not data |
| `_extract_header_vector(vectors)` | Detect and remove year-label header rows (all values look like `2018E`, `2020`, etc.) |
| `_find_row_label(ws, row, start_col)` | Scan leftward in the source sheet to find a text label for the row |
| `_find_col_headers_in_source(ws, col_indices, max_row)` | Find best year/period header row above the data; prefers integer years over datetimes |
| `_is_financial_date(val)` | Return `True` for any recognised financial date/period value (int, datetime, or string) |
| `_are_date_headers(col_headers)` | Return `True` when ≥ 50 % of a sheet's column headers are financial dates |
| `generate_structured_input(mapping_path, excel_path, output_path)` | Main entry point |

**Post-processing rule:** Vectors consisting of (label + exactly 1 data cell) are moved to Config as scalars.

**Auto-transpose rule:** When `_are_date_headers` is `True`, the sheet is transposed — col A = period labels (dates grow downward), row 1 = metric names.  When `False`, original orientation is kept (col A = metric labels, row 1 = period/column headers).

**Line-N fallback rule:** When a row or metric label cannot be resolved from the source workbook — `_find_row_label` returns `None` and no embedded string label is present in the vector — the generator assigns `Line1`, `Line2`, `Line3`, … (counter resets per sheet) instead of the legacy `Row {row_number}` fallback.  This applies in both transposed and original layouts.  Named rows always retain their real label.

---

## Lineage Module

`lineage/lineage_builder.py` builds two levels of lineage:

- **Simple**: sheet-level view (inputs, formulas, outputs per sheet + cross-sheet edges)
- **Complex**: column-level view with every unique formula pattern and full dependency tracking

Uses `mcp_server/smart_formula_sampler.py` to normalise and deduplicate dragged formulas.

Graph rendering: `lineage/lineage_graph.py` (uses `networkx` + `matplotlib`).

---

## MCP Server

`mcp_server/server.py` exposes Excel content to LLM clients via the Model Context Protocol.

**Sampling strategies** (passed as `mode` to fetch tools):

| Mode | Use case |
|------|----------|
| `smart_random` | First-time exploration / overview |
| `full` | Deep analysis of a specific sheet |
| `head` / `tail` | First/last N rows |
| `keyword` | Search for specific text/formula patterns |
| `column_head` / `column_n` | Header-only or Nth column sampling |

**Philosophy:** Always fetch *formulas* first. Formulas reveal business logic; values are just one snapshot.

---

## Configuration

`config.yaml` controls global converter behaviour:

```yaml
delete_unreferenced_hardcoded_values: false
# When true, hardcoded cells not referenced by any formula are excluded
# from the generated code and input template.
```

---

## Module Import Structure

```
excel_to_mapping/
  mapper.py           ← imports excel_to_python, excel_to_python_vectorized.vectorizer
  structured_input_generator.py  ← imports mapper.COLUMNS, excel_to_python
  regenerator.py      ← imports excel_to_python, vectorizer
  main.py             ← CLI; imports mapper, regenerator, structured_input_generator

excel_to_python_vectorized/
  vectorizer.py       ← core grouping logic (no internal project deps)
  converter.py        ← formula conversion
  code_generator.py   ← Python script generation
  main.py             ← CLI

lineage/
  lineage_builder.py  ← imports mcp_server.smart_formula_sampler
  lineage_graph.py    ← imports lineage_builder outputs
```

All subsystems share `excel_to_python.py` as the workbook parser and cell classifier.

---

## Code Style & Conventions

- **All formatting** for the mapping report uses `openpyxl.styles.PatternFill` / `Font` / `Alignment` — never xlwt or xlrd.
- **Column letters** are converted with `col_letter_to_index` / `index_to_col_letter` from `excel_to_python.py` — do not use `openpyxl.utils` column helpers unless necessary.
- **Formula normalisation** (offsetting cell refs by row/col delta) lives in `excel_to_python_vectorized/vectorizer.py` → `normalise_formula()`.
- **Test data** for integration tests is `Indigo.xlsx` in the workspace root. Tests that require it should use `unittest.skipUnless(os.path.exists("Indigo.xlsx"), "Indigo.xlsx not found")`.
- Keep new functions private (prefix `_`) unless they are part of a public API used by `main.py` or tests.
