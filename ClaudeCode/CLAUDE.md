# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# ── End-to-end runners (preferred entry points) ──────────────────────────────

# Full pipeline for an entire workbook
python run_this.py ExcelFiles/Indigo.xlsx

# Single-sheet pipeline with cross-sheet reference support
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Income statement"
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Balance sheet"

# ── Individual layers (lower-level) ──────────────────────────────────────────

# Layer 1: generate mapping report
python run_pipeline.py --layer 1 --input ExcelFiles/Indigo.xlsx --output output/mapping_report.xlsx
python run_pipeline.py --layer 1 --input ExcelFiles/Indigo.xlsx --output output/report.xlsx --log-level DEBUG

# Layer 4a: generate Python calculation script
python -m excel_pipeline.layer4a output/mapping_report.xlsx output/unstructured_inputs.xlsx generated_calculate.py

# Re-run a generated script (after editing unstructured_inputs.xlsx)
cd output/Indigo_Income_statement/
python unstructured_calculate.py     # produces output.xlsx from unstructured_inputs.xlsx

# ── Tests ─────────────────────────────────────────────────────────────────────
pytest tests/
pytest tests/unit/
pytest tests/integration/
pytest -x tests/            # stop on first failure

# ── Code quality ──────────────────────────────────────────────────────────────
black excel_pipeline/
pylint excel_pipeline/
mypy excel_pipeline/
```

## Architecture

Excel financial models → Python code pipeline, structured in layers where each layer's output feeds the next.

### Data Flow

```
Excel file (.xlsx)
  → [freeze other sheets if single-sheet mode]
  → Layer 1:  mapping_report.xlsx        ← single source of truth / contract
  → Layer 2a: unstructured_inputs.xlsx   ← Input cells, original layout
  → Layer 2b: structured_input.xlsx      ← tabular format (not yet complete)
  → Layer 3:  runtime formula engine     ← evaluates intermediate values
  → Layer 4a: unstructured_calculate.py  ← generated Python script (vectorized)
  → execute:  output.xlsx                ← fully populated Excel output
```

### Key Design: `mapping_report.xlsx` is the CONTRACT

Every downstream layer reads from it; users can edit it to control what gets processed. Cell classification:
- **Input** — no formula (user-editable values)
- **Calculation** — has formula AND referenced by other cells
- **Output** — has formula AND NOT referenced by other cells

`IncludeFlag` column: set to `FALSE` to exclude a cell from all downstream processing.

### Entry Point Scripts

| Script | Purpose |
|--------|---------|
| `run_this.py <file>` | Full pipeline for all sheets → `output/<stem>/` |
| `run_this_onesheet.py <file> "<sheet>"` | Single-sheet pipeline → `output/<stem>_<sheet>/` |
| `run_pipeline.py` | Low-level CLI; runs individual layers |

### Core Modules (`excel_pipeline/core/`)

- **`dependency_graph.py`** — Builds precedent/dependent maps from all formulas, detects circular refs, topological sort for calc order. Cell IDs are `"SheetName!A1"`.
- **`cell_classifier.py`** — Classifies cells Input/Calculation/Output via the dependency graph.
- **`formula_analyzer.py`** — Detects dragged formula groups. Groups `>= vectorization_threshold` (default: 10) get `is_vectorizable=True`. Pattern formulas use `{row}`/`{col}` placeholders.
- **`excel_io.py`** — Centralized Excel I/O with atomic saves (temp + rename). Load with `data_only=False` to preserve formulas.

### Layer Modules

- **`layer1/`** — `parser.py` runs 6-step mapping report generation; `cell_extractor.py` produces `CellMetadata` dataclasses; `mapping_writer.py` writes the xlsx.
- **`layer2/`** — `unstructured_generator.py` (2a) and `structured_generator.py` (2b). Partially implemented.
- **`layer3/`** — Runtime calculators via `runtime/formula_engine.py`. Partially implemented.
- **`layer4a/`** — Full Python code generator:
  - `mapping_reader.py` — reads mapping report, produces `CellMetadata` / `GroupMetadata`
  - `dependency_graph.py` — layer-local graph for calculation ordering
  - `formula_translator.py` — Excel formulas → Python `c.get(('Sheet', 'COL', ROW), 0)` expressions; cell store key is `(sheet_name, column_letter, row_int)`
  - `vectorization_engine.py` — pandas/numpy code for vectorizable groups; loop-based fallback for non-vectorizable groups
  - `code_emitter.py` — assembles final standalone script

### Single-Sheet Module (`excel_pipeline/onesheet/`)

Isolated logic for `run_this_onesheet.py`. Does **not** modify any other layer.

- **`freezer.py`** — `freeze_other_sheets(input_path, target_sheet, frozen_path)`:
  - Loads the workbook twice: once `data_only=False` (formulas), once `data_only=True` (cached values)
  - For every non-target sheet, replaces formula strings (`=…`) AND `DataTableFormula`/`ArrayFormula` objects with their last-cached plain values
  - Result: target sheet keeps all formulas; other sheets become value-only → classified as Input by Layer 1
  - Saves frozen workbook to `intermediate/frozen_workbook.xlsx` for inspection

- **`pipeline.py`** — `run(input_path, sheet_name, output_dir, log_level)`:
  - Step 1: freeze → Step 2: Layer 1 → Step 3: Layer 2a → Step 4: Layer 4a → Step 5: execute script
  - Because non-target sheets are Input-only, Layer 2a includes their frozen values in `unstructured_inputs.xlsx`; cross-sheet references in the generated script resolve via `c.get(('OtherSheet', col, row), 0)`

**Why this is faster than the full pipeline:** code generation runs only for the target sheet's Calculation/Output cells (typically ~200 cells vs 20,000+ for a full workbook).

**Null cached value warning:** if `N were None` appears in the log, the workbook was saved before those formulas were calculated. Press `Ctrl+Alt+F9` in Excel, save, and re-run.

### Utils (`excel_pipeline/utils/`)

- **`config.py`** — Singleton `Config`, loaded from `config.yaml`. Use `from excel_pipeline.utils.config import config`.
- **`logging_setup.py`** — Use `get_logger(__name__)` in every module.
- **`helpers.py`** — Financial date detection, cell reference parsing.

### Configuration (`config.yaml`)

Key settings: `performance.vectorization_threshold` (default: 10), `performance.chunk_size` (10000), `validation.tolerance` (1e-9).

### Known Bugs Fixed in `layer4a/`

These were pre-existing and fixed during onesheet pipeline development:

| File | Bug | Fix |
|------|-----|-----|
| `vectorization_engine.py` | Horizontal `{col}` substitution used string concat: `YEAR(' + col + '5)` — invalid Python | Translate with first column letter, then replace that string with `col` variable |
| `formula_translator.py` | Ranges like `D7:D8` inside SUM produced `xl_sum(c.get(...):c.get(...))` — invalid Python | `_expand_ranges()` pre-processor converts `D7:D8` → `D7, D8` before cell substitution |
| `formula_translator.py` | `IF()` regex `([^,]+)` broke on nested commas in `c.get(...)` arguments | Replaced with `_split_args()` parenthesis-aware splitter |
| `formula_translator.py` | Missing Excel functions caused `NameError` at runtime | Added `EOMONTH`, `EDATE`, `YEAR`, `MONTH`, `DAY`, `DATE`, `TODAY`, `DAYS`, `IFERROR`, `SUMIF`, `SUMPRODUCT`, `COUNTA`, `ISNUMBER`, `ISBLANK`, `CONCATENATE` to both `EXCEL_FUNCTIONS` and `HELPER_FUNCTIONS_CODE` |

### Output Layout

```
output/
  <stem>/                          # run_this.py output
    mapping_report.xlsx
    unstructured_inputs.xlsx
    unstructured_calculate.py
    output.xlsx
    pipeline.log

  <stem>_<sheet>/                  # run_this_onesheet.py output
    intermediate/
      frozen_workbook.xlsx         # other sheets value-only (inspect for cross-sheet values)
    mapping_report.xlsx
    unstructured_inputs.xlsx       # target inputs + all frozen other-sheet values
    unstructured_calculate.py
    output.xlsx
    pipeline.log
```

### Performance Considerations

- Range expansion in `dependency_graph.py` is capped at 10,000 cells to prevent memory explosion.
- Vectorizable groups (dragged formulas >= threshold) emit pandas/numpy code rather than per-cell assignments.
- Single-sheet pipeline is significantly faster for iteration on one sheet because dependency analysis and code generation skip all other sheets' formula cells.
- Target: 100MB+ Excel files in under 5 minutes.

### Implementation Status

- Layer 1: Complete and tested
- Layer 2a (unstructured inputs): Complete
- Layer 2b (structured inputs): Partially implemented
- Layer 3 (runtime engine): Partially implemented
- Layer 4a (unstructured code gen): Complete and tested end-to-end
- `onesheet/` module: Complete and tested (`Indigo.xlsx "Income statement"` → 29s, output.xlsx 84KB)
- Tests: Directory structure exists; no tests written yet
