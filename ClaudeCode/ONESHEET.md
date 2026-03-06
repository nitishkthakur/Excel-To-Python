# Single-Sheet Pipeline with Cross-Sheet Reference Support

Run the Excel-to-Python pipeline on one target sheet while preserving all
cross-sheet formula dependencies.

---

## Quick start

```bash
python run_this_onesheet.py <excel_file> "<sheet_name>"
```

**Example — Indigo "Income statement" sheet:**
```bash
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Income statement"
```

**Other examples:**
```bash
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Balance sheet"
python run_this_onesheet.py ExcelFiles/Indigo.xlsx "Assumptions Sheet"
python run_this_onesheet.py ExcelFiles/ACC-Ltd.xlsx "P&L"
python run_this_onesheet.py "ExcelFiles/Bharti-Airtel(2).xlsx" "Income Statement"
```

> **Note:** `.xls` files are not supported. Open them in Excel and
> save as `.xlsx` first.

---

## What it produces

All artifacts land in `output/<filename>_<sheet>/`:

| File | Description |
|------|-------------|
| `intermediate/frozen_workbook.xlsx` | The original workbook with every non-target sheet's formulas replaced by their last-cached values. Inspect this to understand what cross-sheet values are being used. |
| `mapping_report.xlsx` | Layer 1 output: full cell metadata (type, formula, formatting, vectorization groups) for all sheets. Only the target sheet has Calculation/Output cells. |
| `unstructured_inputs.xlsx` | Layer 2a output: Input cells for the target sheet **plus** all frozen values from every other sheet. This is what the generated script reads. |
| `unstructured_calculate.py` | Generated Python calculation script. Standalone — can be run or edited independently. |
| `output.xlsx` | Fully populated Excel output for the target sheet. |
| `pipeline.log` | Detailed run log for all 5 steps. |

---

## How cross-sheet references are preserved

The original approach (strip all other sheets) silently broke formulas like:
```
='Assumptions Sheet'!B5
```
because the referenced sheet no longer existed, defaulting to 0.

This pipeline solves it in `excel_pipeline/onesheet/`:

```
Original workbook
      │
      ▼  freezer.py
Frozen workbook
  ├── target sheet:   formulas intact  (e.g. =EOMONTH('Assumptions Sheet'!B5, 12))
  └── other sheets:   cached values    (e.g. Assumptions Sheet!B5 = 44927)
      │
      ▼  Layer 1
mapping_report.xlsx
  ├── target sheet cells: Input / Calculation / Output
  └── other sheet cells:  all Input  (no formulas → just values)
      │
      ▼  Layer 2a
unstructured_inputs.xlsx
  ├── target sheet Input cells
  └── ALL other-sheet cells  ←  cross-sheet values live here
      │
      ▼  Layer 4a (code generation)
unstructured_calculate.py
  └── c = load_inputs("unstructured_inputs.xlsx")   ← loads everything
      # cross-sheet ref in formula translates to:
      # c.get(('Assumptions Sheet', 'B', 5), 0)   ← resolves correctly
      │
      ▼  execute script
output.xlsx   (fully populated target sheet)
```

### Null cached values
If the frozen workbook reports `N were None` for a sheet, it means the
workbook was last saved before those formulas were calculated in Excel.
Cross-sheet references to those cells will resolve to 0 in the output.
**Fix:** open the original file in Excel, press `Ctrl+Alt+F9` to force-recalculate,
save, then re-run.

---

## New module: `excel_pipeline/onesheet/`

This logic is fully isolated from the rest of the pipeline.

| File | Role |
|------|------|
| `freezer.py` | `freeze_other_sheets(input_path, target_sheet, frozen_path)` — replaces formula cells in non-target sheets with cached values, handles `DataTableFormula` / `ArrayFormula` objects, logs per-sheet stats. |
| `pipeline.py` | `run(input_path, sheet_name, output_dir, log_level)` — orchestrates the 5-step pipeline and is the only entry point used by `run_this_onesheet.py`. |

These files do not modify or import from any other pipeline layer;
they call the existing layer entry points (`generate_mapping_report`,
`generate_unstructured_inputs`, `generate_unstructured_code`) unchanged.

---

## Bug fixes applied to the code generator

Three pre-existing bugs in `layer4a/` were found and fixed during testing:

| File | Bug | Fix |
|------|-----|-----|
| `vectorization_engine.py` | Horizontal `{col}` substitution used broken string-concatenation: `YEAR(' + col + '5)` | Translate with first column letter, then replace the column string with `col` variable |
| `formula_translator.py` | Range references (`D7:D8`) inside `SUM` created invalid Python: `xl_sum(c.get(...):c.get(...))` | Pre-expand `D7:D8` → `D7, D8` before cell-reference substitution |
| `formula_translator.py` | `IF()` regex `([^,]+)` broke on nested commas in `c.get(...)` arguments | Replaced regex with a parenthesis-aware argument splitter (`_split_args`) |

A set of missing Excel helper functions was also added to `HELPER_FUNCTIONS_CODE`
and `EXCEL_FUNCTIONS`:
`EOMONTH`, `EDATE`, `YEAR`, `MONTH`, `DAY`, `DATE`, `TODAY`, `DAYS`,
`IFERROR`, `SUMIF`, `SUMPRODUCT`, `COUNTA`, `ISNUMBER`, `ISBLANK`,
`CONCATENATE`.

---

## Comparison: `run_this.py` vs `run_this_onesheet.py`

| | `run_this.py` | `run_this_onesheet.py` |
|--|--|--|
| Sheets processed | All sheets in the workbook | All sheets (frozen) + target sheet (formulas) |
| Code generated for | All Calculation/Output cells across all sheets | Only target sheet Calculation/Output cells |
| Cross-sheet references | Fully preserved | Fully preserved via frozen values |
| Speed | Slower (full workbook analysis) | Faster (code gen for 1 sheet only) |
| Use case | Full workbook conversion | Iterating on one sheet at a time |

---

## Running the generated script independently

After `run_this_onesheet.py` completes, the generated script is standalone:

```bash
cd output/Indigo_Income_statement/
python unstructured_calculate.py
# re-creates output.xlsx from unstructured_inputs.xlsx
```

You can edit `unstructured_inputs.xlsx`, change input values, and re-run the
script to see how the target sheet recalculates — without re-running the
full pipeline.

---

## Inspecting the intermediate frozen workbook

```bash
# See exactly what values were used for cross-sheet references:
#   output/Indigo_Income_statement/intermediate/frozen_workbook.xlsx
```

Open it in Excel or with openpyxl. The target sheet still has formulas;
every other sheet contains only plain values.
