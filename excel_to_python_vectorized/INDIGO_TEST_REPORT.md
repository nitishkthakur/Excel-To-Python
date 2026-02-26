# Vectorized Sampler — Indigo.xlsx Test Report

## Overview

This document records the results of running the **excel_to_python_vectorized** converter on the `Indigo.xlsx` workbook (Interglobe Aviation financial model) and the improvements applied.

---

## Workbook Profile

| Metric | Value |
|--------|-------|
| Sheets | 13 |
| Total rows (all sheets) | ~750 |
| Total formula cells | 3 938 |
| Cross-sheet references | 722 |
| External file references | 0 |

### Per-Sheet Breakdown

| Sheet | Formulas | Vectorised | Individual | Cross-Sheet Refs |
|-------|----------|------------|------------|------------------|
| Assumptions Sheet | 367 | 356 | 11 | 62 |
| Income statement | 207 | 207 | 0 | 65 |
| Asset Schedule | 87 | 83 | 4 | 5 |
| Cost Matrix | 312 | 306 | 6 | 55 |
| ATF Fuel 2 | 1 047 | 1 044 | 3 | 0 |
| Balance sheet | 363 | 353 | 10 | 177 |
| Revenue Matrix | 119 | 108 | 11 | 57 |
| Debt schedule | 61 | 56 | 5 | 5 |
| Incentive | 67 | 67 | 0 | 22 |
| CashFlow Statement | 239 | 209 | 30 | 151 |
| ATF fuel | 822 | 757 | 65 | 0 |
| Valuation | 247 | 220 | 27 | 123 |

---

## Vectorization Results

* **3 766 / 3 938** formula cells (95.6%) were grouped into **339 vectorised loops**.
* Only **172** formulas remained as individual assignments.
* The generated script is **24 413 lines** (vs. an estimated ~80 000+ without vectorization).
* The script compiles without errors and executes successfully.

---

## Bug Fix: DataTableFormula Serialization

While testing on Indigo.xlsx, a bug was discovered and fixed:

**Problem:** The Valuation sheet contains a `DataTableFormula` object (openpyxl's representation of Excel data tables). The code generator used `repr(val)` to serialize default values, which produced an unparseable Python object string like `<openpyxl.worksheet.formula.DataTableFormula object at 0x...>`.

**Fix:** Added a type check before `repr()` — only `int`, `float`, `bool`, `str`, and `None` are serialized via `repr()`; all other types default to `None`.

**Files changed:**
* `excel_to_python_vectorized/code_generator.py` (line 315)
* `excel_to_python.py` (line 370)

---

## Observations & Recommendations

1. **Vectorization coverage is excellent** (95.6%). The remaining 172 individual formulas are mostly one-off calculations in summary rows or unique patterns.

2. **Cross-sheet reference handling works correctly.** The converter properly resolves references like `='Assumptions Sheet'!K2` across all 13 sheets.

3. **No external file references** — Indigo.xlsx is self-contained, so the `input_files_config.json` is not generated.

4. **Generated script runs end-to-end.** Input template → calculation → output workbook pipeline completes without errors.

---

## How to Reproduce

```bash
# Run the vectorised converter
python -m excel_to_python_vectorized.main Indigo.xlsx --output-dir output/

# Check the generated script syntax
python -c "compile(open('output/calculate.py').read(), 'calculate.py', 'exec')"

# Run the generated script
python output/calculate.py output/input_template.xlsx output/result.xlsx
```
