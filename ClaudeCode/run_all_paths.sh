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

# Create output directory if it doesn't exist
mkdir -p "$OUTPUT_DIR"

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
