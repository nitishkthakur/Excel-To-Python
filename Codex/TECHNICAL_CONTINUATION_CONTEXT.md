# Technical Continuation Context

## 1. Purpose
This file is the persistent handoff state for continuing development of the Excel-to-Python conversion pipeline without re-deriving prior decisions.

Core rule: `mapping_report.xlsx` is the source-of-truth contract for all downstream stages.

## 2. Current Status (Checkpoint)
- Pipeline implemented in `excel_pipeline/` with CLI entry at `run_pipeline.py`.
- `.xls` normalization supports:
  - Windows + Excel COM (`pywin32`) primary path.
  - `soffice --headless` fallback path.
- Stage verification implemented at 2 levels:
  - Intermediate verification (Layer1/Layer2a/Layer2b).
  - Final output verification (cell-by-cell comparison to normalized original).
- Latest aggregate validation in `artifacts/summary.json`:
  - `total_files=22`, `success_count=22`, `failure_count=0`.

## 3. Repo Technical Map
- `excel_pipeline/normalize.py`
  - Normalizes source workbook to `.xlsx`.
- `excel_pipeline/mapping.py`
  - Builds Layer 1 model and writes `mapping_report.xlsx`.
  - Includes dragged-formula grouping and `_FormulaGroups` sheet.
- `excel_pipeline/mapping_io.py`
  - Reads `mapping_report.xlsx` robustly by header names.
- `excel_pipeline/layer2_unstructured.py`
  - Generates `unstructured_inputs.xlsx` from mapping.
- `excel_pipeline/layer2_structured.py`
  - Generates `structured_input.xlsx` (`Index`, `Config`, table sheets).
- `excel_pipeline/reconstruct.py`
  - Rebuilds workbook using mapping + input overrides.
- `excel_pipeline/runtime.py`
  - Runtime calculation for unstructured/structured modes.
- `excel_pipeline/codegen.py`
  - Emits `unstructured_calculate.py` and `structured_calculate.py`.
- `excel_pipeline/compare.py`
  - Final workbook compare (formula/value/style-signature checks).
- `excel_pipeline/verify.py`
  - Intermediate stage verification + stage diagnosis.
- `excel_pipeline/runner.py`
  - Orchestrates full pipeline per workbook / directory, parallel supported.

## 4. Data Contracts

### 4.1 `mapping_report.xlsx`
Sheets:
- One sheet per source sheet.
- `_Metadata`.
- `_FormulaGroups`.

Per-source-sheet columns (`MAPPING_COLUMNS`):
- `Sheet`, `Cell`, `Row`, `Column`, `Type`
- `Formula` (display-safe apostrophe prefix, e.g. `'=SUM(B30:B31)`)
- `Value`, `ValueJSON` (currently kept empty for overflow safety)
- `NumberFormat`, `FontBold`, `FontItalic`, `FontSize`, `FontColor`, `FillColor`
- `HorizontalAlignment`, `VerticalAlignment`, `WrapText`
- `IncludeFlag`
- `GroupID`, `GroupDirection`, `GroupSize`
- `IsDragged`, `GroupRange`, `DragCount`, `DragSummary`
- `PatternFormula`, `StyleJSON`

`_FormulaGroups` columns:
- `GroupID`, `Sheet`, `Direction`, `GroupSize`, `DragCount`
- `CellRange`
- `AnchorCell`, `AnchorFormula`
- `PatternFormula`
- `VectorizationHint`

Example now captured for user request (GAIL BALANCESHEET):
- Anchor formula `'=SUM(B30:B31)` grouped across `BALANCESHEET!B27:J27` with `DragCount=9`.

### 4.2 `unstructured_inputs.xlsx`
- Same sheet structure as source.
- Keeps only included Input values in original positions.
- Formula cells are blanked.

### 4.3 `structured_input.xlsx`
- `Index`: Source↔Input cell mapping contract for runtime.
- `Config`: scalars/small vectors.
- Table sheets: patch-based input tables; transpose if period-like headers detected.

## 5. Key Algorithms and Decisions

### 5.1 Cell classification (Layer 1)
- `Input`: non-formula captured cell.
- `Calculation`: formula referenced by at least one other formula.
- `Output`: formula not referenced by other formulas.

### 5.2 Dragged formula grouping
- Uses pandas vectorized segmentation:
  - Group by `(sheet, canonical_pattern_formula)`.
  - Detect horizontal runs (`diff`, `cumsum`), then vertical runs.
  - Block fallback for full rectangular unassigned groups.
- Group metadata written both per-cell and in `_FormulaGroups` summary.

### 5.3 Reconstruction strategy
- Uses normalized workbook as template for formatting fidelity.
- Applies mapped inputs/formulas as overrides.
- Special formula objects (e.g. `DataTableFormula`) preserved from template when required.

### 5.4 Verification + diagnosis
- Intermediate checks:
  - `verify_mapping_against_original`
  - `verify_unstructured_inputs`
  - `verify_structured_input`
- Final checks:
  - `compare_workbooks`
- Diagnosis routing:
  - `Layer1_Mapping`, `Layer2a_Unstructured`, `Layer2b_Structured`, `Layer3_CodegenOrRuntime`, `clean`.

## 6. Vectorization + Performance Notes
Implemented vectorization/sparse patterns:
- Pandas-based dragged-formula grouping.
- DataFrame-based mapping report row materialization.
- Sparse cell traversal via populated cell store (avoid max-row/max-col full scans).
- Parallel directory processing in runner (`--workers`).

Operational note:
- For large files, use `--workers` and persistent `.cache/normalized`.

## 7. Runbook

### Single workbook
```bash
.venv/bin/python run_pipeline.py \
  --single-file ExcelFiles/Gail-India.xls \
  --output-root artifacts_single \
  --cache-dir .cache/normalized
```

### Full batch
```bash
.venv/bin/python run_pipeline.py \
  --excel-dir ExcelFiles \
  --output-root artifacts \
  --cache-dir .cache/normalized \
  --workers 4
```

### Smoke test
```bash
.venv/bin/python -m unittest tests/test_pipeline.py
```

## 8. Known Constraints
- `openpyxl` cannot preserve all Excel feature extensions identically (warnings expected).
- Style compare intentionally uses a stable subset to avoid false positives from theme/RGB coercion noise.
- `ValueJSON` serialization was de-emphasized to avoid Excel 32,767-char cell limit issues.

## 9. Immediate Next-Session Resume Checklist
1. Read this file and `DOCUMENTATION.md`.
2. Run a single-file check on the target workbook first.
3. If mismatch:
   - Open `artifacts/<Workbook>/validation_report.json`.
   - Check `intermediate_checks.stage_verification` first.
   - Apply fix at earliest failing layer.
4. Re-run full batch before finalizing changes.

## 10. Compressed Context Snapshot (Machine-Readable)
```yaml
context_version: v1
date_local: "2026-03-05"
project: "Excel-to-Python conversion pipeline"
entrypoint: "run_pipeline.py"
golden_rule: "mapping_report.xlsx is the contract"
status:
  pipeline_implemented: true
  stage_verification_2_levels: true
  batch_validation:
    total_files: 22
    success: 22
    failure: 0
key_outputs:
  mapping_report_features:
    - formula_display_safe_with_apostrophe
    - dragged_formula_per_cell_flags
    - _FormulaGroups_summary_sheet
    - drag_count_and_anchor_formula
  layer2:
    - unstructured_inputs
    - structured_input_with_index
  layer3:
    - generated_unstructured_calculate_py
    - generated_structured_calculate_py
key_algorithms:
  - pandas_vectorized_drag_grouping
  - sparse_cell_iteration
  - parallel_directory_execution
critical_files:
  - excel_pipeline/mapping.py
  - excel_pipeline/mapping_io.py
  - excel_pipeline/layer2_structured.py
  - excel_pipeline/verify.py
  - excel_pipeline/runner.py
artifacts:
  summary: "artifacts/summary.json"
  per_workbook_report: "artifacts/<Workbook>/validation_report.json"
next_action_if_resuming:
  - "run single target workbook"
  - "inspect stage_verification on failure"
  - "fix earliest failing layer"
```

## 11. Completion Status Dashboard
Last updated: 2026-03-05

### 11.1 Pipeline Stage Completion
| Stage | Objective | Status | Notes |
|---|---|---|---|
| Layer 1 | Build `mapping_report.xlsx` contract | Complete | Includes formula classification, drag groups, `_FormulaGroups` |
| Layer 2a | Build `unstructured_inputs.xlsx` | Complete | Inputs retained, formulas removed |
| Layer 2b | Build `structured_input.xlsx` | Complete | `Index` + `Config` + sheet tables |
| Layer 3a | Generate and run unstructured calculator | Complete | Produces validated output |
| Layer 3b | Generate and run structured calculator | Complete | Produces validated output |
| Verification | Intermediate + final compare | Complete | Stage-level and final output checks active |

### 11.2 Validation Completion
| Scope | Total | Passed | Failed | Status |
|---|---:|---:|---:|---|
| `ExcelFiles/` batch | 22 | 22 | 0 | Complete |
| Smoke tests | 1 | 1 | 0 | Complete |

### 11.3 Artifact Completion
| Artifact | Location | Status |
|---|---|---|
| Batch summary | `artifacts/summary.json` | Complete |
| Per-workbook reports | `artifacts/<Workbook>/validation_report.json` | Complete |
| Technical handoff | `TECHNICAL_CONTINUATION_CONTEXT.md` | Complete |
| User docs | `DOCUMENTATION.md` | Complete |

## 12. Continuation Protocol (Resume Exactly Where Left Off)
Use this when starting a new session:

1. Read:
   - `TECHNICAL_CONTINUATION_CONTEXT.md`
   - `DOCUMENTATION.md`
2. Confirm baseline status:
   - Open `artifacts/summary.json`
   - Check `success_count`, `failure_count`
3. If user asks a targeted fix:
   - Run single-file pipeline first
   - Inspect `validation_report.json` for that file
   - Fix earliest failing stage (`Layer1` -> `Layer2` -> `Layer3`)
4. After targeted fix:
   - Re-run single-file validation
   - Re-run full-batch validation
5. Update this file:
   - Update “Last updated”
   - Update status tables
   - Add a short entry in the session log

## 13. Open/Pending Work (Continuation Backlog)
These are optional enhancements, not blockers:

- Add richer style fidelity checks when needed for strict formatting audits.
- Add dedicated benchmarks per stage (timing by workbook and layer).
- Add unit tests for formula-group edge cases (mixed horizontal/vertical overlaps).
- Add CSV export for `_FormulaGroups` to support external review.

## 14. Session Log (Append-Only)
Use one line per session to preserve continuity.

| Date | Summary | Result | Next Step |
|---|---|---|---|
| 2026-03-05 | Implemented full pipeline + vectorized formula grouping + stage verification + docs + drag summary fields | 22/22 pass | Continue from user-requested refinements only |
