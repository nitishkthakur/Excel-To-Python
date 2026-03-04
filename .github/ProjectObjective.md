# Project Objective: Excel-to-Python Conversion Pipeline

You are working on a system that converts an Excel financial model into a reproducible, editable Python pipeline. The end-to-end flow has five layers.

---

## Layer 1 — Intermediate Mapping Report

Parse the original Excel workbook (`.xlsx` / `.xls`) and produce a comprehensive intermediate file.

- **Classify every cell** into one of three types:
  - `Input` — plain hardcoded value, no formula
  - `Calculation` — formula cell referenced by at least one other formula
  - `Output` — formula cell that is NOT referenced by any other formula (terminal)
- For dragged/repeated formulas, detect the pattern and group them using formula normalisation. Record: `GroupID`, `GroupDirection`, `GroupSize`, `PatternFormula`.
- Record per-cell metadata: `Sheet`, `Cell`, `Type`, `Formula`, `Value`, number format, font (bold / italic / size / color), fill color, alignment, `WrapText`, `IncludeFlag`.
- Write everything into **`mapping_report.xlsx`** — one sheet per source sheet plus a `_Metadata` sheet.
- A human reviewer may open it, flip `IncludeFlag` rows, or add rows before feeding it back to the next stage.
- This file must contain all information needed to reconstruct the original workbook and drive code generation. **It is the contract between the parser and the generator.**

---

## Layer 2a — Unstructured Input File (`unstructured_inputs.xlsx`)

Produce the simplest possible editable input file for the user.

- Read `mapping_report.xlsx`.
- Extract only cells where `Type == "Input"` and `IncludeFlag == True`.
- **Remove all formula cells entirely.** Retain only raw hardcoded input values in their original sheet layout and cell positions.
- The objective: the user edits this file with their new data, and it is used downstream to reproduce the exact output via the calculations recorded in `mapping_report.xlsx`.

---

## Layer 2b — Structured Input File (`structured_input.xlsx`)

Produce a clean, tabular input file organised by sheet, suitable for bulk data entry.

- Read `mapping_report.xlsx` (prefer this alone; avoid reading additional files unless necessary).
- Extract only cells where `Type == "Input"` and `IncludeFlag == True`.
- Scalars (isolated or length-1 cells) always go to the **Config sheet**, with their source sheet reference recorded.
- Identify contiguous rectangular patches of Input cells. Each patch becomes a table in the output.
- A single source sheet may produce **multiple input tables** if it contains patches with different header types (e.g. one patch with financial-date column headers and another with non-date headers).
- **Auto-transpose rule:** if the column headers of a patch are financial dates/periods (integers like `2020`, strings like `"2020E"`, or datetime objects), transpose the table so that rows = periods and columns = metrics. Otherwise keep the original orientation.
- If a row/column label cannot be resolved from the source sheet, assign `Line1`, `Line2`, … as the label.
- Write **`structured_input.xlsx`** with:
  - **Index sheet** — cross-reference between this file and `mapping_report.xlsx`; must be useful both for the human user and for downstream code generation
  - **Config sheet** — all scalars and short (label + 1 data value) vectors
  - **Per-source-sheet tabs** — one or more tabular input tabs per source sheet that contains ≥ 1 vector input

---

## Layer 3a — Unstructured Code Generation (`unstructured_calculate.py`)

- Write Python code that accepts `mapping_report.xlsx` and generates `unstructured_calculate.py`.
- `unstructured_calculate.py` must accept `unstructured_inputs.xlsx` and produce `output.xlsx` that matches the original workbook in formulas, formatting, and values.
- **Test rigorously.** Run against all provided Excel files. Write helper functions to compare output values cell-by-cell against the original. Understand and resolve every mismatch before moving on.

---

## Layer 3b — Structured Code Generation (`structured_calculate.py`)

- Write Python code that accepts `mapping_report.xlsx` and `structured_input.xlsx` and generates `structured_calculate.py`.
- `structured_calculate.py` must accept `structured_input.xlsx` and produce `output.xlsx` that matches the original workbook in formulas, formatting, and values.
- **Test rigorously.** Run against all provided Excel files. Write helper functions to compare output values cell-by-cell against the original. Understand and resolve every mismatch before moving on.

---

## Two Paths, One Output

After `mapping_report.xlsx` is generated the user has two options:

| Path | Input file edited by user | Generator script |
|------|--------------------------|-----------------|
| Unstructured | `unstructured_inputs.xlsx` | `unstructured_calculate.py` |
| Structured | `structured_input.xlsx` | `structured_calculate.py` |

Both paths must produce an identical `output.xlsx` that matches the original workbook in formulas, formatting, and values.

---

## Process Guidelines

- **Plan deeply.** Anticipate issues, re-plan, review results, and understand every mismatch before proceeding to the next stage.
- `mapping_report.xlsx` is the single source of truth for all downstream stages. Always refer back to it when in doubt.
- **Always test. Always verify results.** 
- Write modular, well-documented code. Produce a documentation file that explains the overall architecture, the purpose of each module, and how they interact.
- For mapping_report.xlsx, the input files, output files - all make sure to read the file and check if the content is as expected. It is meant for human and downstream code understanding. Decide what it must look like and what it must have. Check any information captured in it with the original file. Make each stage bulletproof before moving on. If the content is not as expected, understand why and fix the issue before proceeding to the next stage. Read outputs of every stage till they are perfect. 

---

## Golden Rule

> `mapping_report.xlsx` is the contract between the Excel parser and the code generator. Never bypass it. All decisions about what to include, group, or regenerate are encoded in that file.

## Secondary outputs
- Documentation.md - must contain detailed documentation of the overall architecture, the purpose of each module, and how they interact. It should also include instructions for running the code, testing, and troubleshooting.
- Test scripts - must include comprehensive tests for each stage of the pipeline
- Mermaid diagrams - must include diagrams that illustrate the architecture and flow of the pipeline, as well as the structure of the intermediate files.

