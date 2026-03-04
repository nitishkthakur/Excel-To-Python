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
- Vector inputs (time-series rows) live in named sheets, one per source sheet, with metric labels in column A and period values in subsequent columns.
- An `Index` sheet cross-references every table back to its source sheet and cell range.

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
| **Dates as rows, not columns** | Allows arbitrary extension of the time dimension without schema changes. |
| **Formula logic stays in code** | Users never edit formulas; they only supply input values. |
| **Format preservation** | The output must be indistinguishable in layout from the original so existing downstream consumers of the Excel file are unaffected. |
| **Reviewable audit trail** | The mapping report exposes every cell decision (Input / Calculation / Output) so a reviewer can verify the conversion before trusting the output. |
