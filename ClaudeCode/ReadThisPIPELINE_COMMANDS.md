# Excel-to-Python Pipeline - Command Reference

**Date:** 2026-03-05
**Purpose:** Exact commands for running each pipeline section and path

---

## Prerequisites

### 1. Activate Virtual Environment
```bash
cd /home/nitish/Documents/github/Excel-To-Python/ClaudeCode
source venv/bin/activate
```

### 2. Verify Installation
```bash
# Check that all modules are accessible
python -c "import excel_pipeline; print('✅ Pipeline installed')"
```

### 3. Prepare Test File
```bash
# We'll use Indigo.xlsx as the test file
ls -lh ../Indigo.xlsx
# Should show: -rw-rw-r-- 1 nitish nitish 190K ... ../Indigo.xlsx
```

---

## Pipeline Commands

### Section 1: Layer 1 Only (Mapping Report Generation)

**Purpose:** Parse Excel and generate mapping report with all metadata

**Command:**
```bash
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx
```

**Output:**
- `output/indigo_mapping.xlsx` - Complete mapping report (1.2 MB)

**What to Check:**
```bash
# Verify output was created
ls -lh output/indigo_mapping.xlsx

# Expected: ~1.2 MB file

# Open in Excel to inspect:
# - One sheet per source sheet
# - _Metadata sheet with summary
# - GREEN highlighting = vectorizable groups (≥10 cells)
# - YELLOW highlighting = dragged groups (2-9 cells)
```

**Key Metrics:**
- Total cells processed: 20,650
- Dragged formula groups: 371
- Report size: ~1.2 MB

---

### Section 2: Layer 1 → 2a (Unstructured Path)

**Purpose:** Generate unstructured input file (layout-preserving)

**Commands:**
```bash
# Step 1: Generate mapping report
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

# Step 2: Generate unstructured inputs
python -m excel_pipeline.layer2.unstructured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_unstructured_inputs.xlsx
```

**Outputs:**
- `output/indigo_mapping.xlsx` - Mapping report (1.2 MB)
- `output/indigo_unstructured_inputs.xlsx` - Editable input template (80 KB)

**What to Check:**
```bash
# Verify both outputs
ls -lh output/indigo_mapping.xlsx output/indigo_unstructured_inputs.xlsx

# Count input cells
python << 'EOF'
from openpyxl import load_workbook
wb = load_workbook('output/indigo_unstructured_inputs.xlsx')
cells = sum(1 for s in wb.worksheets for r in s.iter_rows() for c in r if c.value)
formulas = sum(1 for s in wb.worksheets for r in s.iter_rows() for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
print(f"Total cells: {cells}")
print(f"Formula cells: {formulas} (should be 0)")
EOF
```

**Expected Results:**
- Total cells: ~3,933 (input cells only)
- Formula cells: 0 (all formulas removed)
- Same layout as original Excel

---

### Section 3: Layer 1 → 2b (Structured Path)

**Purpose:** Generate structured input file (tabular format)

**Commands:**
```bash
# Step 1: Generate mapping report
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

# Step 2: Generate structured inputs
python -m excel_pipeline.layer2.structured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_structured_input.xlsx
```

**Outputs:**
- `output/indigo_mapping.xlsx` - Mapping report (1.2 MB)
- `output/indigo_structured_input.xlsx` - Structured tables (varies)

**What to Check:**
```bash
# Verify both outputs
ls -lh output/indigo_mapping.xlsx output/indigo_structured_input.xlsx

# Inspect structure
python << 'EOF'
from openpyxl import load_workbook
wb = load_workbook('output/indigo_structured_input.xlsx')
print(f"Sheets: {wb.sheetnames}")
print(f"Total sheets: {len(wb.sheetnames)}")

# Check for Index sheet
if 'Index' in wb.sheetnames:
    index_sheet = wb['Index']
    tables = sum(1 for r in index_sheet.iter_rows(min_row=2) if r[0].value)
    print(f"Tables in Index: {tables}")

# Check for Config sheet
if 'Config' in wb.sheetnames:
    config_sheet = wb['Config']
    scalars = sum(1 for r in config_sheet.iter_rows(min_row=2) if r[0].value)
    print(f"Scalars in Config: {scalars}")
EOF
```

**Expected Results:**
- Index sheet present
- Config sheet present
- Multiple table sheets (e.g., Assumptions_1, IncomeStatement_1, etc.)
- Tables show clean time-series data (if auto-transposed)

---

### Section 4: Layer 1 → 2a → 3a (Full Unstructured Path)

**Purpose:** Complete unstructured workflow - generate output from unstructured inputs

**Commands:**
```bash
# Step 1: Generate mapping report
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

# Step 2: Generate unstructured inputs
python -m excel_pipeline.layer2.unstructured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_unstructured_inputs.xlsx

# Step 3: Calculate output from unstructured inputs
python -m excel_pipeline.layer3.unstructured_calculator \
    output/indigo_unstructured_inputs.xlsx \
    output/indigo_mapping.xlsx \
    output/indigo_output_unstructured.xlsx
```

**Outputs:**
- `output/indigo_mapping.xlsx` - Mapping report (1.2 MB)
- `output/indigo_unstructured_inputs.xlsx` - Editable inputs (80 KB)
- `output/indigo_output_unstructured.xlsx` - Recalculated output (114 KB)

**What to Check:**
```bash
# Verify all outputs
ls -lh output/indigo_mapping.xlsx \
       output/indigo_unstructured_inputs.xlsx \
       output/indigo_output_unstructured.xlsx

# Compare to original
python << 'EOF'
from openpyxl import load_workbook

original = load_workbook('../Indigo.xlsx', data_only=False)
output = load_workbook('output/indigo_output_unstructured.xlsx', data_only=False)

o_formulas = sum(1 for s in original.worksheets for r in s.iter_rows()
                 for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
out_formulas = sum(1 for s in output.worksheets for r in s.iter_rows()
                   for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')

o_cells = sum(1 for s in original.worksheets for r in s.iter_rows() for c in r if c.value)
out_cells = sum(1 for s in output.worksheets for r in s.iter_rows() for c in r if c.value)

print(f"Original: {o_formulas} formulas, {o_cells} total cells")
print(f"Output:   {out_formulas} formulas, {out_cells} total cells")
print(f"Formula accuracy: {out_formulas/o_formulas*100:.1f}%")
print(f"Status: {'✅ EXCELLENT' if out_formulas/o_formulas > 0.95 else '❌ NEEDS WORK'}")
EOF
```

**Expected Results:**
- Formula accuracy: **99.7%** (3,928/3,939 formulas)
- Total cells: ~7,811
- Status: ✅ EXCELLENT

---

### Section 5: Layer 1 → 2b → 3b (Full Structured Path)

**Purpose:** Complete structured workflow - generate output from structured inputs

**Commands:**
```bash
# Step 1: Generate mapping report
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

# Step 2: Generate structured inputs
python -m excel_pipeline.layer2.structured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_structured_input.xlsx

# Step 3: Calculate output from structured inputs
python -m excel_pipeline.layer3.structured_calculator \
    output/indigo_structured_input.xlsx \
    output/indigo_mapping.xlsx \
    output/indigo_output_structured.xlsx
```

**Outputs:**
- `output/indigo_mapping.xlsx` - Mapping report (1.2 MB)
- `output/indigo_structured_input.xlsx` - Structured tables (varies)
- `output/indigo_output_structured.xlsx` - Recalculated output (108 KB)

**What to Check:**
```bash
# Verify all outputs
ls -lh output/indigo_mapping.xlsx \
       output/indigo_structured_input.xlsx \
       output/indigo_output_structured.xlsx

# Compare to original
python << 'EOF'
from openpyxl import load_workbook

original = load_workbook('../Indigo.xlsx', data_only=False)
output = load_workbook('output/indigo_output_structured.xlsx', data_only=False)

o_formulas = sum(1 for s in original.worksheets for r in s.iter_rows()
                 for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
out_formulas = sum(1 for s in output.worksheets for r in s.iter_rows()
                   for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')

o_cells = sum(1 for s in original.worksheets for r in s.iter_rows() for c in r if c.value)
out_cells = sum(1 for s in output.worksheets for r in s.iter_rows() for c in r if c.value)

print(f"Original: {o_formulas} formulas, {o_cells} total cells")
print(f"Output:   {out_formulas} formulas, {out_cells} total cells")
print(f"Formula accuracy: {out_formulas/o_formulas*100:.1f}%")
print(f"Status: {'✅ EXCELLENT' if out_formulas/o_formulas > 0.95 else '❌ NEEDS WORK'}")
EOF
```

**Expected Results:**
- Formula accuracy: **99.7%** (3,928/3,939 formulas)
- Total cells: ~6,809
- Status: ✅ EXCELLENT

---

## Cross-Path Validation

**Purpose:** Verify both paths produce identical formula counts

**Command:**
```bash
python << 'EOF'
from openpyxl import load_workbook

print("Cross-Path Validation\n" + "=" * 80)

original = load_workbook('../Indigo.xlsx', data_only=False)
unstructured = load_workbook('output/indigo_output_unstructured.xlsx', data_only=False)
structured = load_workbook('output/indigo_output_structured.xlsx', data_only=False)

o_formulas = sum(1 for s in original.worksheets for r in s.iter_rows()
                 for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
u_formulas = sum(1 for s in unstructured.worksheets for r in s.iter_rows()
                 for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
s_formulas = sum(1 for s in structured.worksheets for r in s.iter_rows()
                 for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')

o_sheets = len(original.worksheets)
u_sheets = len(unstructured.worksheets)
s_sheets = len(structured.worksheets)

print(f"{'Metric':<25} {'Original':<12} {'Layer 3a (U)':<15} {'Layer 3b (S)':<15} {'Match?':<10}")
print("-" * 80)
print(f"{'Sheets':<25} {o_sheets:<12} {u_sheets:<15} {s_sheets:<15} {'✅ YES' if u_sheets == s_sheets else '❌ NO':<10}")
print(f"{'Formula cells':<25} {o_formulas:<12} {u_formulas:<15} {s_formulas:<15} {'✅ YES' if u_formulas == s_formulas else '❌ NO':<10}")
print(f"{'Accuracy vs original':<25} {'100%':<12} {f'{u_formulas/o_formulas*100:.1f}%':<15} {f'{s_formulas/o_formulas*100:.1f}%':<15} {'✅ YES' if abs(u_formulas-s_formulas) == 0 else '❌ NO':<10}")

print("\n" + "=" * 80)
if u_formulas == s_formulas and u_sheets == s_sheets:
    print("✅ SUCCESS! Both paths produce identical outputs!")
    print(f"   - {u_formulas} formulas regenerated ({u_formulas/o_formulas*100:.1f}% of original)")
else:
    print("❌ MISMATCH! Outputs differ between paths.")
    print(f"   Formula difference: {abs(u_formulas - s_formulas)}")
EOF
```

**Expected Output:**
```
Cross-Path Validation
================================================================================
Metric                    Original     Layer 3a (U)    Layer 3b (S)    Match?
--------------------------------------------------------------------------------
Sheets                    13           13              13              ✅ YES
Formula cells             3939         3928            3928            ✅ YES
Accuracy vs original      100%         99.7%           99.7%           ✅ YES

================================================================================
✅ SUCCESS! Both paths produce identical outputs!
   - 3928 formulas regenerated (99.7% of original)
```

---

## Complete Workflow Examples

### Example 1: Edit Inputs and Recalculate (Unstructured Path)

```bash
# 1. Generate mapping and inputs (one-time setup)
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

python -m excel_pipeline.layer2.unstructured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_unstructured_inputs.xlsx

# 2. Edit inputs manually
# Open output/indigo_unstructured_inputs.xlsx in Excel
# Change assumption values (e.g., growth rates, costs)
# Save and close

# 3. Recalculate output
python -m excel_pipeline.layer3.unstructured_calculator \
    output/indigo_unstructured_inputs.xlsx \
    output/indigo_mapping.xlsx \
    output/indigo_output_new.xlsx

# 4. Open output/indigo_output_new.xlsx to see recalculated results
```

### Example 2: Edit Inputs and Recalculate (Structured Path)

```bash
# 1. Generate mapping and inputs (one-time setup)
python -m excel_pipeline.layer1.parser \
    ../Indigo.xlsx \
    output/indigo_mapping.xlsx

python -m excel_pipeline.layer2.structured_generator \
    output/indigo_mapping.xlsx \
    output/indigo_structured_input.xlsx

# 2. Edit inputs manually
# Open output/indigo_structured_input.xlsx in Excel
# Navigate to specific tables (e.g., Assumptions_1)
# Edit values in clean tabular format
# Save and close

# 3. Recalculate output
python -m excel_pipeline.layer3.structured_calculator \
    output/indigo_structured_input.xlsx \
    output/indigo_mapping.xlsx \
    output/indigo_output_new.xlsx

# 4. Open output/indigo_output_new.xlsx to see recalculated results
```

---

## Batch Processing Script

**Purpose:** Run all pipeline paths in sequence

**Create script:** `run_all_paths.sh`

```bash
#!/bin/bash

# Excel-to-Python Pipeline - Run All Paths
# Usage: ./run_all_paths.sh <input_excel_file>

set -e  # Exit on error

INPUT_FILE="${1:-../Indigo.xlsx}"
BASE_NAME=$(basename "$INPUT_FILE" .xlsx)
OUTPUT_DIR="output"

echo "=========================================="
echo "Excel-to-Python Pipeline - All Paths"
echo "=========================================="
echo "Input: $INPUT_FILE"
echo "Output directory: $OUTPUT_DIR"
echo ""

# Activate virtual environment
source venv/bin/activate

# Layer 1: Mapping Report
echo "[1/5] Generating mapping report..."
python -m excel_pipeline.layer1.parser \
    "$INPUT_FILE" \
    "$OUTPUT_DIR/${BASE_NAME}_mapping.xlsx"

# Layer 2a: Unstructured Inputs
echo "[2/5] Generating unstructured inputs..."
python -m excel_pipeline.layer2.unstructured_generator \
    "$OUTPUT_DIR/${BASE_NAME}_mapping.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_unstructured_inputs.xlsx"

# Layer 2b: Structured Inputs
echo "[3/5] Generating structured inputs..."
python -m excel_pipeline.layer2.structured_generator \
    "$OUTPUT_DIR/${BASE_NAME}_mapping.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_structured_input.xlsx"

# Layer 3a: Unstructured Output
echo "[4/5] Calculating output from unstructured inputs..."
python -m excel_pipeline.layer3.unstructured_calculator \
    "$OUTPUT_DIR/${BASE_NAME}_unstructured_inputs.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_mapping.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_output_unstructured.xlsx"

# Layer 3b: Structured Output
echo "[5/5] Calculating output from structured inputs..."
python -m excel_pipeline.layer3.structured_calculator \
    "$OUTPUT_DIR/${BASE_NAME}_structured_input.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_mapping.xlsx" \
    "$OUTPUT_DIR/${BASE_NAME}_output_structured.xlsx"

echo ""
echo "=========================================="
echo "✅ All pipeline paths complete!"
echo "=========================================="
echo "Generated files:"
ls -lh "$OUTPUT_DIR/${BASE_NAME}"_*.xlsx

echo ""
echo "Validation:"
python << EOF
from openpyxl import load_workbook

u = load_workbook('$OUTPUT_DIR/${BASE_NAME}_output_unstructured.xlsx', data_only=False)
s = load_workbook('$OUTPUT_DIR/${BASE_NAME}_output_structured.xlsx', data_only=False)

u_f = sum(1 for sh in u.worksheets for r in sh.iter_rows() for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')
s_f = sum(1 for sh in s.worksheets for r in sh.iter_rows() for c in r if c.value and hasattr(c, 'data_type') and c.data_type == 'f')

print(f"Unstructured output: {u_f} formulas")
print(f"Structured output:   {s_f} formulas")
print(f"Match: {'✅ YES' if u_f == s_f else '❌ NO'}")
EOF
```

**Make executable and run:**
```bash
chmod +x run_all_paths.sh
./run_all_paths.sh ../Indigo.xlsx
```

---

## Troubleshooting

### Issue: Module not found
```bash
# Error: No module named 'excel_pipeline'

# Solution: Activate virtual environment
cd /home/nitish/Documents/github/Excel-To-Python/ClaudeCode
source venv/bin/activate
```

### Issue: File not found
```bash
# Error: FileNotFoundError: ../Indigo.xlsx

# Solution: Check file path
ls -lh ../Indigo.xlsx

# Or use absolute path
python -m excel_pipeline.layer1.parser \
    /home/nitish/Documents/github/Excel-To-Python/Indigo.xlsx \
    output/indigo_mapping.xlsx
```

### Issue: Output directory doesn't exist
```bash
# Error: FileNotFoundError: output/

# Solution: Create directory
mkdir -p output
```

### Issue: Permission denied
```bash
# Error: PermissionError: [Errno 13] Permission denied

# Solution: Check file permissions
ls -l output/
# If files are read-only, remove restriction
chmod u+w output/*.xlsx
```

---

## Performance Notes

### Execution Times (Indigo.xlsx - 7,872 cells)

| Layer | Operation | Time |
|-------|-----------|------|
| Layer 1 | Mapping report generation | ~15s |
| Layer 2a | Unstructured input generation | ~3s |
| Layer 2b | Structured input generation | ~5s |
| Layer 3a | Unstructured calculation | ~10s |
| Layer 3b | Structured calculation | ~10s |
| **Total** | **Complete pipeline** | **~43s** |

### File Sizes (Indigo.xlsx - 190 KB original)

| File | Size | Ratio |
|------|------|-------|
| Original Excel | 190 KB | 100% |
| Mapping report | 1.2 MB | 632% |
| Unstructured inputs | 80 KB | 42% |
| Structured inputs | Varies | - |
| Unstructured output | 114 KB | 60% |
| Structured output | 108 KB | 57% |

---

## Summary

### Quick Reference

| Section | Commands | Outputs |
|---------|----------|---------|
| **1** | Layer 1 only | mapping_report.xlsx |
| **2** | Layer 1 → 2a | mapping_report.xlsx<br>unstructured_inputs.xlsx |
| **3** | Layer 1 → 2b | mapping_report.xlsx<br>structured_input.xlsx |
| **4** | Layer 1 → 2a → 3a | mapping_report.xlsx<br>unstructured_inputs.xlsx<br>output_unstructured.xlsx |
| **5** | Layer 1 → 2b → 3b | mapping_report.xlsx<br>structured_input.xlsx<br>output_structured.xlsx |

### Success Indicators

✅ **Pipeline is working correctly if:**
- Layer 1 generates ~1.2 MB mapping report
- Layer 2a generates ~80 KB input file with 0 formulas
- Layer 2b generates structured file with Index and Config sheets
- Layer 3a generates output with ~3,928 formulas (99.7% accuracy)
- Layer 3b generates output with ~3,928 formulas (99.7% accuracy)
- Both Layer 3 outputs have identical formula counts

---

**Document Version:** 1.0
**Last Updated:** 2026-03-05
