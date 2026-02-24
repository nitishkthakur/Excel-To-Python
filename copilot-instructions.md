# Copilot Instructions — Excel-To-Python Code Structure

This document provides a comprehensive description of the repository layout, module responsibilities, data flow, and conventions used throughout the **Excel-To-Python** project. It is intended to help an AI assistant (or any new developer) quickly orient themselves and resume productive development.

---

## 1  Repository Overview

The project has two main subsystems that share low-level Excel-reading utilities:

| Subsystem | Entry point | Purpose |
|---|---|---|
| **Excel → Python converter** | `excel_to_python.py` | Converts an Excel workbook (formulas, formatting, cross-sheet references) into a standalone Python script that reproduces the same calculations. |
| **MCP Server** | `mcp_server/server.py` | Exposes the workbook's structure and data to LLM clients (e.g. Claude, Copilot Chat) via the **Model Context Protocol**, enabling interactive analysis of Excel files. |

---

## 2  Top-Level Files

### `excel_to_python.py`
The main converter. It:
1. Loads a workbook with `openpyxl`.
2. Parses every sheet (cell values, formulas, merged cells, formatting, column widths, named tables).
3. Classifies cells into *formula* cells and *hardcoded* cells using `classify_cells()`.
4. Builds a dependency graph among formula cells via `find_all_references()` and topologically sorts them with `build_dependency_order()` (Kahn's algorithm).
5. Translates each formula into a Python expression using `FormulaConverter` (from `formula_converter.py`).
6. Emits a complete, runnable Python script (`generate_python_script()`) that reads inputs from a template workbook, evaluates all formulas in dependency order, and writes a new output workbook with matching formatting.

### `formula_converter.py`
A self-contained module that translates Excel formula strings into Python code.

* **`FormulaConverter`** — main class. Handles cell references (`A1`, `$A$1`, `Sheet1!A1`), range references, table structured references, 80+ Excel function names → Python helper mappings, and operators.
* **`HELPER_FUNCTIONS_CODE`** — a long string constant containing the Python implementations of all mapped Excel functions (`xl_sum`, `xl_vlookup`, `xl_if`, …). This is injected verbatim into the generated script.
* Helper utilities: `col_letter_to_index`, `index_to_col_letter`, `cell_to_var_name`, `range_to_var_name`, `table_ref_to_var_name`.

### `config.yaml`
A small YAML configuration file consumed by the converter. Currently contains one setting (`delete_unreferenced_hardcoded_values`).

### `algorithm.md`
Documentation describing the high-level algorithm used by the converter.

### `requirements.txt` (root)
Top-level dependencies: `openpyxl>=3.1.0`, `pyyaml>=6.0`.

---

## 3  `mcp_server/` Package

This directory is a self-contained MCP server. It can be started with `python server.py` (working directory must be `mcp_server/` or the package must be importable). The MCP JSON configuration is in `mcp.json`.

### 3.1 Sampling Strategies

Reading large Excel sheets in full would overwhelm an LLM's context window. The server therefore supports **three pluggable sampling strategies**, each implemented in its own module:

| Mode name | Module | Description |
|---|---|---|
| `smart_random` | `excel_reader_smart_sampler.py` | The original strategy. Detects rectangular *data regions* in each sheet, then picks a representative sample: header rows, formula rows, and a spread of head / middle / tail data rows. Controlled by `max_sample_rows`. |
| `full` | `excel_sample_full.py` | Loads every cell in each detected region. Parameters `nrows` and `ncols` (both default to `None`) allow optional capping. `None` = load everything. |
| `column_n` | `column_n.py` | Extracts a **vertical strip**: locates the first data column in each region (the label column, which typically holds line-item names like "Revenue", "COGS") and loads that column plus the next *N* columns (default 10). All rows within the region are included. |

All three modules expose the same function signature:

```python
def extract_sheet_data(path: str, sheet_name: str, **kwargs) -> dict[str, Any]
```

The returned dict always has the shape:

```python
{
    "sheet_name": str,
    "regions": [
        {
            "region": str,        # human-readable description
            "min_row": int,
            "max_row": int,
            "min_col": int,
            "max_col": int,
            "headers": [str, ...],
            "rows": [{"row_number": int, "values": [Any, ...]}, ...],
            "formulas": [{"address": str, "formula": str, "cached_value": Any}, ...],
        },
        ...
    ],
    "sampled": bool,
    "total_rows": int,
    "sampled_rows": int,
}
```

This consistent shape means every strategy's output is directly consumable by the formatters (`to_markdown`, `to_json`, `to_xml`).

### 3.2 `excel_reader_smart_sampler.py`

Core low-level reading module. In addition to the `extract_sheet_data` entry-point it provides:

* **`DataRegion`** — value class representing a rectangular data block (row/col bounds, optional header row).
* **`detect_regions(ws)`** — scans a worksheet top-to-bottom, grouping contiguous non-blank rows into regions and detecting header rows.
* **`sample_row_indices(region, ws_formula, max_rows)`** — the "smart random" logic that selects representative rows.
* **`extract_formulas(path, sheet_name)`** — returns every formula cell in a sheet.
* **`workbook_summary(path)`** — lightweight metadata about all sheets.
* **`sheet_names(path)`** — simple list of sheet names.
* Helpers: `open_workbook`, `open_workbook_values`, `_cell_info`, `_cell_addr`.

The other two sampling modules (`excel_sample_full.py`, `column_n.py`) import region detection and cell helpers from this module to avoid code duplication.

### 3.3 `server.py`

The MCP server itself, built on `mcp.server.fastmcp.FastMCP`. It exposes five tools:

1. **`list_sheets`** — returns sheet names (JSON array).
2. **`get_workbook_summary`** — returns structural metadata.
3. **`get_sheet_data`** — main data-reading tool. Accepts a `mode` parameter (`smart_random` | `full` | `column_n`) and delegates to the correct sampler via the internal `_dispatch_extract()` helper. Also accepts mode-specific parameters (`max_sample_rows`, `nrows`, `ncols`, `num_columns`).
4. **`get_sheet_formulas`** — extracts all formulas from a sheet.
5. **`get_sheet_sample`** — convenience wrapper around `get_sheet_data` with a smaller default sample size (30 rows). Also accepts the `mode` parameter.

**Adding a new sampling strategy** requires:
1. Creating a new module in `mcp_server/` with an `extract_sheet_data(path, sheet_name, **kwargs)` function.
2. Importing it in `server.py`.
3. Adding the mode name to `VALID_MODES`.
4. Adding an `elif` branch in `_dispatch_extract()`.

### 3.4 `formatters.py`

Stateless conversion of the extracted-data dict to three output formats:

* `to_markdown(data)` — Markdown tables + formula lists.
* `to_json(data, pretty=True)` — JSON.
* `to_xml(data)` — XML via `xml.etree.ElementTree`.

### 3.5 `mcp.json`

MCP client configuration. Tells the client how to start the server:

```json
{ "mcpServers": { "excel-analysis": { "command": "python", "args": ["server.py"], "cwd": "mcp_server" } } }
```

---

## 4  `tests/` Directory

* **`create_sample_workbook.py`** — programmatically builds a multi-sheet `.xlsx` file used as a test fixture. Three sheets: *Inputs* (prices, quantities, subtotals), *Summary* (cross-sheet refs, conditionals), *Rates* (table-like data). Saved to `tests/test_data/sample.xlsx` (git-ignored).
* **`test_excel_to_python.py`** — end-to-end tests for the converter pipeline: config loading, workbook parsing, cell classification, dependency ordering.
* **`test_formula_converter.py`** — unit tests for `FormulaConverter`: column/index helpers, cell variable naming, formula translation.
* **`test_sampling.py`** — tests for the three sampling strategies (`smart_random`, `full`, `column_n`) and the MCP-server dispatch logic. Uses the sample workbook created by `create_sample_workbook.py`.

Tests are run with **pytest** from the repository root:

```bash
python -m pytest tests/ -v
```

---

## 5  Data Flow Summary

```
Excel file (.xlsx)
        │
        ▼
  openpyxl loads workbook (formula view + value view)
        │
        ├──▶ detect_regions(ws)  →  list[DataRegion]
        │
        ├──▶ Sampling strategy picks rows/cols to read
        │       ├── smart_random: sample_row_indices()
        │       ├── full: all rows, optional nrows/ncols cap
        │       └── column_n: label col + N adjacent cols
        │
        ├──▶ _cell_info() reads each selected cell
        │
        └──▶ Result dict  ──▶  Formatter (md / json / xml)  ──▶  MCP client
```

---

## 6  Conventions & Patterns

* **Import style** — relative imports within `mcp_server/`; the server adds its own directory to `sys.path` implicitly via `cwd`.
* **Type hints** — used throughout (`list[str]`, `dict[str, Any]`, `int | None`).
* **Workbook lifecycle** — every function that opens a workbook closes it before returning. Formula workbook (`data_only=False`) and value workbook (`data_only=True`) are always opened in tandem.
* **Region-centric design** — all samplers operate per-region, not per-sheet. A sheet may have multiple disjoint data blocks (separated by blank rows), and each is processed independently.
* **Consistent return shape** — every `extract_sheet_data` function returns the same dict structure so formatters and server code never need to check which sampler was used.

---

## 7  Extending the Project

* **New sampling strategy** — see §3.3 above for the four-step checklist.
* **New output format** — add a `to_<format>(data)` function in `formatters.py` and a branch in `server.py`.
* **New Excel function** — add a `xl_<name>` implementation in `HELPER_FUNCTIONS_CODE` (inside `formula_converter.py`) and map the Excel name in `FormulaConverter._function_map`.
