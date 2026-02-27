# Excel Analysis MCP Server

An MCP (Model Context Protocol) server that reads `.xlsx` workbooks and exposes their content — including formulas, cached values, and structural metadata — to LLM clients such as GitHub Copilot. The goal is to let the LLM generate a plain-language business summary of what an Excel file calculates, sheet by sheet.

> **Formula-first philosophy** — unless the user explicitly asks for raw
> values, always fetch *formulas* first.  Formulas reveal the calculation
> logic and business rules; values are just one snapshot of the result.

## Quick Start

```bash
cd mcp_server
pip install -r requirements.txt
python server.py          # starts the stdio MCP server
```

### Connecting from VS Code / GitHub Copilot

Add the following to your workspace `.vscode/mcp.json` (create the file if it doesn't exist):

```jsonc
{
  "mcpServers": {
    "excel-analysis": {
      "command": "python",
      "args": ["/absolute/path/to/mcp_server/server.py"],
      "cwd": "/absolute/path/to/mcp_server"
    }
  }
}
```

### Connecting from Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-analysis": {
      "command": "python",
      "args": ["/absolute/path/to/mcp_server/server.py"],
      "cwd": "/absolute/path/to/mcp_server"
    }
  }
}
```

---

## Sampling Strategies

All five strategies extract **both formulas and values** from every cell
they visit — you never need a separate call just to get formulas from the
sampled rows.

| Mode | When to use | What it loads | Budget applies to |
|---|---|---|---|
| `"smart_random"` | **First-time exploration / overview.** Best balance of token cost and coverage. Use ONLY when exploring the Excel file for the first time or getting an overview. | Headers, all formula rows (up to half the budget), plus head/mid/tail data rows. | Rows per sheet |
| `"full"` | **Detailed analysis of a specific sheet.** Use when comprehensively analysing a particular sheet. Avoid on sheets with thousands of rows. | Every row and column in each detected region (optionally capped by `nrows`/`ncols`). | Rows/cols per region |
| `"column_n"` | Wide sheets (many period columns) where you only need a vertical slice. | The label column + the next *N* data columns, all rows. | Columns per region |
| `"row_head"` | **Scanning non-priority sheets efficiently.** Captures all patch headers and the first few data rows. Perfect for scanning sheets where you need the structure but not all data. | Header row + first N data rows of each detected patch. Budget divided across patches proportionally. | Rows per sheet |
| `"column_head"` | **Financial sheets with date columns.** Columns are dates (Q1 2023, Q2 2023, …) and rows are entities (Income, Profit, Revenue). Captures entity labels + first N date columns. | All rows, first N columns of each patch. Budget divided across patches proportionally. | Columns per sheet |

All five modes are available in **`get_sheet_data`** and **`get_sheet_sample`** via the `mode` parameter.

---

## Tools Reference

### `list_sheets`

List all sheet names in an Excel workbook. Use this as the first step when you receive an unknown file.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |

**Example call:**

```
list_sheets(file_path="/home/user/reports/q4_financials.xlsx")
→ '["Sales", "Costs", "Summary"]'
```

---

### `get_workbook_summary`

Return a lightweight structural summary of the entire workbook — sheet names, dimensions, detected data regions, and formula counts. **Call this first** to decide which sheets deserve deeper analysis.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |

**Example call:**

```
get_workbook_summary(file_path="/data/report.xlsx")
```

---

### `get_sheet_formulas`

Extract every formula from a sheet. **Use this first** when analysing a sheet — formulas encode the business logic.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |
| `sheet_name` | string | yes | Exact sheet name (case-sensitive) | `"Summary"` |

---

### `get_sheet_data`

Read a sheet and return its data (with formulas) in the requested format. Large sheets are automatically sampled. **All five sampling modes return both formulas and values.**

| Argument | Type | Required | Default | Description | When to set |
|---|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file | Always required |
| `sheet_name` | string | yes | — | Exact sheet name (case-sensitive) | Always required |
| `format` | string | no | `"markdown"` | Output format: `"markdown"`, `"json"`, or `"xml"` | Set to `"json"` for programmatic processing, `"xml"` for XML interchange |
| `mode` | string | no | `"smart_random"` | Sampling strategy — see table below | Set based on your analysis goal |
| `max_sample_rows` | int | no | `100` | Row budget **per sheet** (for `smart_random` and `row_head` modes) | Increase for wider coverage, decrease to save tokens |
| `nrows` | int | no | `None` | Max rows per region (only `full` mode) | Set when you want to cap full-mode output |
| `ncols` | int | no | `None` | Max columns per region (only `full` mode) | Set when you want to cap full-mode output |
| `num_columns` | int | no | `10` | Data columns after label column (only `column_n` mode) | Set to 3-5 for narrow views |
| `max_cols` | int | no | `20` | Column budget **per sheet** (only `column_head` mode) | Set to 10 for narrow, 30 for wide coverage |

**Mode selection guide:**

| Mode value | When to set this value |
|---|---|
| `"smart_random"` | Set this when exploring an Excel file for the first time or getting a general overview. This is the default. |
| `"full"` | Set this when you need to comprehensively analyse a specific sheet in detail — all rows, all columns. |
| `"column_n"` | Set this when the sheet is very wide (many columns) and you only need the label column + a few data columns. |
| `"row_head"` | Set this to scan sheets efficiently — captures all patch headers and first few rows. Use for non-priority sheets. |
| `"column_head"` | Set this for financial sheets where columns are dates and rows are entities (Income, Profit, etc.). |

---

### `get_sheet_sample`

Get a very small sample of a sheet — useful for a quick preview. Identical to `get_sheet_data` but defaults to 30 rows instead of 100.

| Argument | Type | Required | Default | Description |
|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file |
| `sheet_name` | string | yes | — | Exact sheet name |
| `max_rows` | int | no | `30` | Row budget **per sheet** (for `smart_random` and `row_head` modes) |
| `format` | string | no | `"markdown"` | Output format |
| `mode` | string | no | `"smart_random"` | Sampling strategy (same 5 modes as `get_sheet_data`) |
| `num_columns` | int | no | `10` | Data columns after label (only `column_n` mode) |
| `max_cols` | int | no | `20` | Column budget **per sheet** (only `column_head` mode) |

---

## Recommended Tool Combinations for LLMs

### Pathway 1: Reading an Unknown Excel File for Overview

Use this when the user provides an Excel file you have never seen before
and you need to understand its structure and content.

```
Step 1 → get_workbook_summary(file_path="/data/unknown.xlsx")
         Read the structural summary: sheet names, dimensions,
         region counts, formula counts.  Identify which sheets
         are large, which have formulas, and which are small.

Step 2 → For each sheet, get a quick sample in smart_random mode:
         get_sheet_sample(file_path="/data/unknown.xlsx",
                          sheet_name="<sheet>",
                          mode="smart_random",
                          max_rows=20)

Step 3 → get_sheet_formulas(file_path="/data/unknown.xlsx",
                            sheet_name="<sheet_with_formulas>")
         Pull formulas from sheets that have formula_count > 0
         to understand the calculation logic.

Step 4 → Synthesise a summary:
         "Based on the workbook summary and samples, here is what
          this Excel file contains and calculates…"
```

### Pathway 2: User Mentions an Important Sheet

Use this when the user says something like *"The P&L sheet is the most
important one"* or *"Focus on the Revenue sheet"*.

```
Step 1 → get_workbook_summary(file_path="/data/financial_model.xlsx")
         Understand the full workbook structure.

Step 2 → Load the important sheet in FULL mode for detailed analysis:
         get_sheet_data(file_path="/data/financial_model.xlsx",
                        sheet_name="P&L",
                        mode="full")

Step 3 → Load all OTHER sheets in ROW_HEAD mode for efficient scanning:
         get_sheet_data(file_path="/data/financial_model.xlsx",
                        sheet_name="<other_sheet>",
                        mode="row_head",
                        max_sample_rows=30)
         This captures all headers and first few rows of every patch
         without consuming excessive tokens.

Step 4 → get_sheet_formulas(file_path="/data/financial_model.xlsx",
                            sheet_name="P&L")
         Trace the formula dependencies in the important sheet.

Step 5 → Produce the detailed analysis of the important sheet,
         referencing context from the other sheets scanned in Step 3.
```

### Pathway 3: Financial Model with Date Columns

Use this for workbooks where sheets have many quarterly/monthly date
columns spanning multiple years and row indices are financial entities.

```
Step 1 → get_workbook_summary(file_path="/data/forecast.xlsx")

Step 2 → For wide sheets with date columns, use column_head mode:
         get_sheet_data(file_path="/data/forecast.xlsx",
                        sheet_name="Revenue",
                        mode="column_head",
                        max_cols=10)
         This reads ALL entity rows but only the first 10 date columns,
         so the LLM can see all entities and relate them to dates.

Step 3 → For the key analysis sheet, use full mode:
         get_sheet_data(file_path="/data/forecast.xlsx",
                        sheet_name="Summary",
                        mode="full")

Step 4 → get_sheet_formulas for sheets with cross-references.
```

### Pathway 4: Formula-Focused Analysis (Trace Calculation Chains)

When analysing an Excel file, **start with formulas** unless the user
explicitly asks for data values.

```
Step 1 → get_workbook_summary(file_path="/data/budget.xlsx")
         Identify sheets that have formulas (formula_count > 0).

Step 2 → get_sheet_formulas(file_path="/data/budget.xlsx",
                            sheet_name="Assumptions")
Step 3 → get_sheet_formulas(file_path="/data/budget.xlsx",
                            sheet_name="P&L")
         Pull formulas from every relevant sheet.

Step 4 → Ask the LLM:
         "Trace the cross-sheet formula references.  Starting from the
          Assumptions sheet, explain how each derived quantity flows into
          the P&L sheet."
```

---

## Handling Large Files

Sheets with more than **100 rows per data region** are automatically sampled when using `smart_random` mode (the default). The sampling strategy ensures the LLM sees enough to understand the sheet:

1. **Header rows** — always included so column names are visible.
2. **Formula rows** — always included (up to half the budget) so calculation logic is captured.
3. **Head / middle / tail** data rows — evenly spread to show the range of values.

**Budget is always per sheet** — for modes that support multiple patches (`row_head`, `column_head`), the budget is divided across patches proportionally to their size.

| Goal | How |
|---|---|
| Bigger sample for more coverage | `get_sheet_data(..., max_sample_rows=200)` |
| Smaller sample to save tokens | `get_sheet_data(..., max_sample_rows=30)` or use `get_sheet_sample` |
| Disable sampling entirely | `get_sheet_data(..., mode="full")` — **use with caution** on huge files |
| Vertical slice of a wide sheet | `get_sheet_data(..., mode="column_n", num_columns=5)` |
| Efficient header scan | `get_sheet_data(..., mode="row_head", max_sample_rows=30)` |
| Date-column financial sheets | `get_sheet_data(..., mode="column_head", max_cols=10)` |

---

## Unstructured Sheets

Real-world Excel files are often messy:

- Company names or logos pasted in the top rows
- Blank rows separating different tables
- Headers that don't start on row 1
- Multiple "patches" of data on the same sheet with their own headers

The server handles all of these by scanning for contiguous non-blank row runs and treating each as an independent **data region**. Each region is reported separately with its own detected headers, row data, and formulas.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `FileNotFoundError` | Wrong path or relative path | Use an absolute path, e.g. `"/home/user/file.xlsx"` |
| `KeyError: '<sheet>'` | Sheet name doesn't exist or is misspelled | Call `list_sheets` first to get exact names |
| Output is too large | Sheet has thousands of rows and `mode="full"` | Switch to `mode="smart_random"` or `mode="row_head"` or reduce `max_sample_rows` |
| `cached_value` is `null` | File was never opened in Excel after formulas were written | Open and save the file in Excel, then re-run |
| Wide sheet too many columns | Financial model with 50+ date columns | Use `mode="column_head"` with `max_cols=10` |
