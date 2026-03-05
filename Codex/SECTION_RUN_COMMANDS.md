# Section Run Commands

Use these exact commands from repo root.

## Common Inputs
- Source workbook: replace with your target file in `ExcelFiles/`
- Output root: `artifacts_sections`
- Cache: `.cache/normalized`

## 1) Only Layer 1
```bash
.venv/bin/python run_stage.py \
  --mode 1 \
  --file ExcelFiles/Gail-India.xls \
  --output-root artifacts_sections \
  --cache-dir .cache/normalized
```

## 2) Layer 1 -> Layer 2a
```bash
.venv/bin/python run_stage.py \
  --mode 1-2a \
  --file ExcelFiles/Gail-India.xls \
  --output-root artifacts_sections \
  --cache-dir .cache/normalized
```

## 3) Layer 1 -> Layer 2b
```bash
.venv/bin/python run_stage.py \
  --mode 1-2b \
  --file ExcelFiles/Gail-India.xls \
  --output-root artifacts_sections \
  --cache-dir .cache/normalized
```

## 4) Layer 1 -> Layer 2a -> Layer 3a
```bash
.venv/bin/python run_stage.py \
  --mode 1-2a-3a \
  --file ExcelFiles/Gail-India.xls \
  --output-root artifacts_sections \
  --cache-dir .cache/normalized
```

## 5) Layer 1 -> Layer 2b -> Layer 3b
```bash
.venv/bin/python run_stage.py \
  --mode 1-2b-3b \
  --file ExcelFiles/Gail-India.xls \
  --output-root artifacts_sections \
  --cache-dir .cache/normalized
```

## Output Location
For input workbook `ExcelFiles/Gail-India.xls`, outputs are written under:
- `artifacts_sections/Gail-India/`

Files appear depending on mode:
- `mapping_report.xlsx`
- `unstructured_inputs.xlsx`
- `structured_input.xlsx`
- `unstructured_calculate.py`
- `structured_calculate.py`
- `output_unstructured.xlsx`
- `output_structured.xlsx`
