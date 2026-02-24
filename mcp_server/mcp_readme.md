# Excel Analysis MCP Server

An MCP (Model Context Protocol) server that reads `.xlsx` workbooks and exposes their content — including formulas, cached values, and structural metadata — to LLM clients such as GitHub Copilot. The goal is to let the LLM generate a plain-language business summary of what an Excel file calculates, sheet by sheet.

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

Alternatively, copy or symlink the provided `mcp.json` from this directory into your client's configuration folder.

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

## Tools Reference

### `list_sheets`

List all sheet names in an Excel workbook.  Use this as the first step when you receive an unknown file.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |

**Example call:**

```
list_sheets(file_path="/home/user/reports/q4_financials.xlsx")
```

**Example output:**

```json
["Sales", "Costs", "Summary"]
```

---

### `get_workbook_summary`

Return a lightweight structural summary of the entire workbook — sheet names, dimensions, detected data regions, and formula counts.  **Call this first** to decide which sheets deserve deeper analysis.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |

**Example call:**

```
get_workbook_summary(file_path="/data/report.xlsx")
```

**Example output (abbreviated):**

```json
{
  "file": "/data/report.xlsx",
  "sheets": [
    {
      "name": "Sales",
      "max_row": 500,
      "max_column": 10,
      "regions": ["DataRegion(rows=1-2, cols=1-1, header=1)",
                  "DataRegion(rows=4-500, cols=1-10, header=4)"],
      "region_count": 2,
      "formula_count": 48
    },
    {
      "name": "Costs",
      "max_row": 20,
      "max_column": 5,
      "region_count": 1,
      "formula_count": 8
    }
  ]
}
```

---

### `get_sheet_data`

Read a sheet and return its data (with formulas) in the requested format.  Large sheets are automatically sampled.

| Argument | Type | Required | Default | Description | Example values |
|---|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |
| `sheet_name` | string | yes | — | Exact sheet name (case-sensitive) | `"Sales"`, `"P&L Summary"` |
| `format` | string | no | `"markdown"` | Output format: `"markdown"`, `"json"`, or `"xml"` | `"json"` |
| `max_sample_rows` | int | no | `100` | Max rows per data region | `50`, `200` |
| `full` | bool | no | `false` | Disable sampling and return every row | `true` |

**Example calls:**

```
# Markdown overview (default, sampled)
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")

# JSON with a smaller sample
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Costs",
               format="json", max_sample_rows=50)

# All rows from a small lookup table
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Lookups",
               format="xml", full=True)
```

**Example Markdown output (abbreviated):**

```markdown
## Sheet: Sales

_Sampled 18 of 500 rows._

### Region 1  (rows 1–2, cols A–A)

| Acme Corp |
| --- |
| Q4 2025 Sales Report |

### Region 2  (rows 4–500, cols A–D)

| Product | Units | Price | Revenue |
| --- | --- | --- | --- |
| Widget A | 100 | 9.99 | 999 |
| Widget B | 250 | 14.99 | 3747.5 |
| ... | ... | ... | ... |

**Formulas:**

- `D5`: `=B5*C5`  → 999
- `D6`: `=B6*C6`  → 3747.5
- `D500`: `=SUM(D5:D499)`  → 125000
```

---

### `get_sheet_formulas`

Extract every formula from a sheet.  Best for tracing cross-sheet references and deeply nested derived quantities.

| Argument | Type | Required | Description | Example |
|---|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |
| `sheet_name` | string | yes | Exact sheet name | `"Summary"` |

**Example call:**

```
get_sheet_formulas(file_path="/data/report.xlsx", sheet_name="Summary")
```

**Example output:**

```json
[
  {"address": "B4", "formula": "=SUM(B2:B3)", "cached_value": 4500},
  {"address": "B6", "formula": "=Sales!D500-B4", "cached_value": 120500}
]
```

---

### `get_sheet_sample`

Get a very small sample of a sheet — useful for a quick preview before requesting full data.

| Argument | Type | Required | Default | Description | Example values |
|---|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file | `"/data/report.xlsx"` |
| `sheet_name` | string | yes | — | Exact sheet name | `"Raw Data"` |
| `max_rows` | int | no | `30` | Max rows to include | `10`, `50` |
| `format` | string | no | `"markdown"` | Output format | `"json"`, `"xml"` |

**Example calls:**

```
# Quick Markdown preview
get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales")

# Tiny JSON preview (10 rows) to check column names
get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Raw Data",
                 max_rows=10, format="json")
```

---

## Recommended Tool Combinations

### 1. Full Workbook Analysis  (generate a business summary)

This is the most common workflow.  It gives the LLM enough context to write a plain-language summary of what the workbook calculates.

```
Step 1 → get_workbook_summary(file_path="/data/report.xlsx")
         Understand how many sheets exist, their sizes, and formula counts.

Step 2 → get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")
         Repeat for each sheet (or focus on sheets with formulas).

Step 3 → get_sheet_formulas(file_path="/data/report.xlsx", sheet_name="Sales")
         Pull the full formula list for sheets with cross-sheet references or
         complex nested calculations.  Ask the LLM to trace the dependency
         chain and explain each derived quantity.
```

**Prompt to the LLM after gathering data:**

> "Based on the workbook summary and the sheet data above, write a business
> summary of what this Excel file calculates. For each sheet, explain: the
> purpose, key inputs, key outputs, and how formulas derive the outputs."

---

### 2. Quick Preview of a Large File

When the file is very large and you want to minimise context usage:

```
Step 1 → list_sheets(file_path="/data/big_model.xlsx")
         See all sheet names.

Step 2 → get_sheet_sample(file_path="/data/big_model.xlsx",
                          sheet_name="Revenue", max_rows=20)
         Glance at the structure and column headers.

Step 3 → get_sheet_data(file_path="/data/big_model.xlsx",
                        sheet_name="Revenue", max_sample_rows=50)
         Get a bigger sample with formulas for a more complete picture.
```

---

### 3. Formula-Focused Analysis  (trace calculation chains)

When the goal is to understand the calculation logic rather than the data:

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

### 4. Comparing Formats  (pick the best for your use case)

| Format | Best for |
|---|---|
| `"markdown"` | Human-readable output, quick summaries, conversational use |
| `"json"` | Programmatic processing, feeding into downstream scripts |
| `"xml"` | Structured interchange, systems that consume XML |

```
# Get the same sheet in all three formats to compare
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales",
               format="markdown")
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales",
               format="json")
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales",
               format="xml")
```

---

## Handling Large Files

Sheets with more than **100 rows per data region** are automatically sampled.  The sampling strategy ensures the LLM sees enough to understand the sheet:

1. **Header rows** — always included so column names are visible.
2. **Formula rows** — always included (up to half the budget) so calculation logic is captured.
3. **Head / middle / tail** data rows — evenly spread to show the range of values.

You can control the sample size:

| Goal | How |
|---|---|
| Bigger sample for more coverage | `get_sheet_data(..., max_sample_rows=200)` |
| Smaller sample to save tokens | `get_sheet_data(..., max_sample_rows=30)` or use `get_sheet_sample` |
| Disable sampling entirely | `get_sheet_data(..., full=True)` — **use with caution** on huge files |

---

## Unstructured Sheets

Real-world Excel files are often messy:

- Company names or logos pasted in the top rows
- Blank rows separating different tables
- Headers that don't start on row 1
- Multiple "patches" of data on the same sheet with their own headers

The server handles all of these by scanning for contiguous non-blank row runs and treating each as an independent **data region**.  Each region is reported separately with its own detected headers, row data, and formulas.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `FileNotFoundError` | Wrong path or relative path | Use an absolute path, e.g. `"/home/user/file.xlsx"` |
| `KeyError: '<sheet>'` | Sheet name doesn't exist or is misspelled | Call `list_sheets` first to get exact names |
| Output is too large | Sheet has thousands of rows and `full=True` | Remove `full=True` or reduce `max_sample_rows` |
| `cached_value` is `null` | File was never opened in Excel after formulas were written | Open and save the file in Excel, then re-run |
