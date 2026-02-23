# Excel-to-Python Converter

Converts an Excel workbook with formulas, cross-sheet references, and table references into a standalone Python calculation script.

## Features

- **Multi-sheet support**: Parses all sheets in a workbook and preserves cross-sheet references
- **Formula conversion**: Translates Excel formulas (SUM, AVERAGE, IF, VLOOKUP, etc.) to Python
- **Table references**: Supports Excel table structured references (e.g., `Table1[Column]`)
- **Format preservation**: Replicates cell formatting (fonts, fills, number formats, column widths, merged cells) in the output
- **Input template**: Generates an input template Excel file with the same sheet structure for users to provide input values
- **Configurable filtering**: Option to exclude unreferenced hardcoded values via `config.yaml`

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Step 1: Convert Excel to Python

```bash
python excel_to_python.py <input_excel_file> [--config config.yaml] [--output-dir output]
```

This generates two files in the output directory:
- `calculate.py` — The Python script that performs the calculations
- `input_template.xlsx` — An Excel file pre-filled with the original hardcoded values

### Step 2: Edit Input Values

Open `input_template.xlsx` and modify any hardcoded values you want to change. The template follows the same sheet names and cell positions as the original workbook — you only need to enter the input values.

### Step 3: Run the Calculation

```bash
python output/calculate.py output/input_template.xlsx result.xlsx
```

The output `result.xlsx` replicates the original workbook's structure with recalculated values.

## Configuration

Edit `config.yaml` to control converter behavior:

```yaml
# When true, hardcoded values not referenced by any formula are excluded
# from the generated code and input template.
# When false (default), all hardcoded values are preserved.
delete_unreferenced_hardcoded_values: false
```

## Supported Excel Functions

SUM, AVERAGE, COUNT, COUNTA, MIN, MAX, IF, AND, OR, NOT, ABS, ROUND, ROUNDUP, ROUNDDOWN, INT, MOD, POWER, SQRT, LEN, LEFT, RIGHT, MID, UPPER, LOWER, TRIM, CONCATENATE, TEXT, VALUE, VLOOKUP, HLOOKUP, INDEX, MATCH, IFERROR, ISBLANK, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, SUMPRODUCT, ROW, COLUMN, ROWS, COLUMNS, TODAY, NOW, YEAR, MONTH, DAY, DATE, EOMONTH, EDATE, DATEDIF, PI

## Running Tests

```bash
pip install pytest
python -m pytest tests/ -v
```
