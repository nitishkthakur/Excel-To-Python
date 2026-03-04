# Project Objective — Excel-to-Python

## Goal

Convert an **unstructured Excel workbook** (with formulas, cross-sheet references, and hardcoded inputs scattered across cells) into a **fully programmable Python pipeline** that any user can run without knowledge of the underlying Excel structure.

---

## The Three-Stage Workflow

```
Original Excel (.xlsx)
        │
        ▼  [Stage 1 — Analysis]
  Intermediate Files
  ├── mapping_report.xlsx   ← full cell-level audit (for developers / reviewers)
  └── structured_input.xlsx ← clean tabular input template (for users)
        │
        ▼  [Stage 2 — User Edit]
  User fills structured_input.xlsx
        │
        ▼  [Stage 3 — Execution]
  Python regenerates output.xlsx  ← same format as the original workbook
```

---

## Stage 1 — Analysis & Intermediate Files

The original Excel is parsed and two intermediate files are produced:

### `mapping_report.xlsx`
A **full audit table** of every cell in the workbook — its type (Input / Calculation / Output), formula text, value, formatting, and grouping metadata.

- Intended audience: developers and reviewers who want to see exactly what the workbook contains.
- Future automated code reads from this file to drive regeneration.
- One sheet per source sheet, plus a `_Metadata` sheet.
- Schema is fixed (19 columns); never reorder or rename them.

### `structured_input.xlsx`
A **clean, user-facing input template** derived from the mapping report.

- Intended audience: end users who need to supply new input values.
- **Purely tabular** — no merged cells, no embedded formulas, no decoration.
- **Dates / time-periods are row indices, not columns.** Each period (year, quarter, month, etc.) occupies one row so that the user can add or remove as many periods as needed without restructuring the file.
- Scalar inputs (single-cell values) live in the `Config` sheet.
- Vector inputs (time-series data) live in named sheets, one per source sheet.  When the original column headers are recognised as financial date/period labels the table is **automatically transposed**: col A = period label, remaining columns = one metric per column.  When headers are not dates the original orientation is kept (col A = metric, remaining columns = one period per column).
- An `Index` sheet cross-references every table back to its source sheet and cell range.

---

## Financial Date / Period Detection

The generator recognises a wide range of financial date and period formats so it can decide whether to transpose a table.  The following patterns are treated as date-like (handled case-insensitively):

| Category | Examples |
|----------|---------|
| Quarter (number first) | `1Q2021`, `4Q21`, `1Q-2024` |
| Quarter (Q prefix) | `Q12023`, `Q421`, `Q1-2024` |
| Quarter (year first) | `2024Q1`, `24Q4`, `20241Q`, `2024-1Q` |
| Half-year | `H12024`, `H1-24`, `2024H1`, `24-H2` |
| Fiscal year | `FY2024`, `FYE2024`, `FY24E`, `FYE24A` |
| Calendar year tag | `CY2024`, `CY24E` |
| Year + financial suffix | `2024E`, `2024A`, `2024F`, `2024B`, `2024P` |
| Plain year (integer or string) | `2023`, `2024` |
| Month + year | `Jan-24`, `Jan-2024`, `Jan 2024` |
| Year + numeric month | `2024-03`, `2024/3` |
| Full date strings | `02-01-2024`, `2024-01-02`, `01/02/2024` |
| Relative period labels | `LTM`, `NTM`, `TTM`, `YTD`, `LTM 2024` |

The detection logic lives in `excel_to_mapping/structured_input_generator.py`:

- `_is_financial_date(val)` — returns `True` for a single value (int, float, datetime, or string)
- `_are_date_headers(col_headers)` — returns `True` when ≥ 50 % of a vector-sheet's column headers are recognised as financial dates

### Auto-transpose rule

When `_are_date_headers` returns `True` for a sheet's column headers, `_build_vector_sheet` writes the table in **transposed** orientation:

```
Period [Source]   | Revenue  | Costs  | EBITDA  | ...
------------------+----------+--------+---------+----
2022              | 100      | 80     | 20      |
2023              | 110      | 88     | 22      |
2024E             | 120      | 94     | 26      |
```

This lets a user add a new time period simply by **appending a row** — no column restructuring required.  When headers are not dates the original orientation (metrics as rows, periods as columns) is preserved.

---

## Line-N Fallback Labels

When a row or metric label cannot be resolved from the source workbook (i.e., `_find_row_label` returns `None` and no embedded string label is present in the vector), the generator assigns a **context-agnostic sequential label** instead of a raw cell reference:

- `Line1`, `Line2`, `Line3`, … (counter resets to 1 for every vector sheet)

This applies in **both layout orientations**:

| Layout | Where the label appears | Fallback |
|--------|------------------------|----------|
| **Transposed** (dates as rows) | Row 1, column headers (metric names) | `Line1`, `Line2`, … |
| **Original** (metrics as rows) | Col A, row labels | `Line1`, `Line2`, … |

**Rule:** A row receives a `LineN` label only when its label is completely unresolvable — i.e., no text string exists to the left of the data in the source sheet AND the vector's first cell is not a string.  Named rows (e.g., "Revenue from operations") always retain their real label.

The fallback `"Row {row_number}"` label is no longer used anywhere in the output.

---

## Stage 2 — User Edit

The user opens `structured_input.xlsx` and:
1. Reviews the `Index` sheet to understand what each table represents.
2. Edits scalar values in `Config`.
3. Edits or extends time-series rows in the named vector sheets (add rows for new periods, delete rows for periods no longer needed).

No knowledge of the original Excel structure is required.

---

## Stage 3 — Execution

The user runs a single Python command:

```bash
python -m excel_to_mapping.main regenerate output/mapping_report.xlsx \
    --inputs output/structured_input.xlsx \
    --output output/result.xlsx
```

The regenerator:
1. Reads the mapping report to reconstruct the workbook structure and formula logic.
2. Reads the user-edited `structured_input.xlsx` for new input values.
3. Outputs `result.xlsx` in **exactly the same format** as the original workbook — same sheets, same layout, same formatting — but with recalculated values based on the new inputs.

---

## Design Principles

| Principle | Rationale |
|-----------|-----------|
| **Tabular intermediates** | Structured files are readable by both humans and code; no opaque binary state. |
| **Dates as rows, not columns** | When column headers are financial date/period labels the table is transposed so each period occupies one row.  Adding a new time period is simply appending a row; no schema changes are needed. |
| **Formula logic stays in code** | Users never edit formulas; they only supply input values. |
| **Format preservation** | The output must be indistinguishable in layout from the original so existing downstream consumers of the Excel file are unaffected. |
| **Reviewable audit trail** | The mapping report exposes every cell decision (Input / Calculation / Output) so a reviewer can verify the conversion before trusting the output. |
