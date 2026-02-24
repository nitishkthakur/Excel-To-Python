# Excel Analysis MCP Server

An MCP server that reads `.xlsx` workbooks and exposes formulas, cached values, and structural metadata to LLM clients (GitHub Copilot, Claude Desktop, etc.). The primary goal is to let the LLM generate a plain-language business summary of what an Excel file calculates.

## Quick Start

```bash
cd mcp_server
pip install -r requirements.txt
python server.py          # starts the stdio MCP server
```

Add to your client configuration (VS Code `.vscode/mcp.json` or Claude Desktop `claude_desktop_config.json`):

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

---

## Tools at a Glance

| Tool | Purpose | When to use |
|---|---|---|
| `list_sheets` | List sheet names | First step on an unknown file |
| `get_workbook_summary` | Structural overview (dimensions, regions, formula counts) | Immediately after — to decide which sheets matter |
| **`get_sheet_formulas`** | **Extract every formula** (address, formula string, cached value) | **Primary analysis tool — always call this before `get_sheet_data`** |
| `get_sheet_data` | Return row data with formulas in Markdown / JSON / XML (auto-sampled) | After formulas, to see the surrounding values and context |
| `get_sheet_sample` | Tiny preview (default 30 rows) | Quick glance at structure before committing to a larger read |

> **Key principle:** `get_sheet_formulas` is the most important tool. Formulas encode the business logic — they tell you *what* a sheet calculates. Values alone are just numbers. **Always call `get_sheet_formulas` first**, then use `get_sheet_data` to fill in the surrounding context.

---

## Tools Reference

### `list_sheets`

List all sheet names. Use as the first step when you receive an unknown file.

| Argument | Type | Required | Description |
|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file |

```
list_sheets(file_path="/data/report.xlsx")
→ ["Sales", "Costs", "Summary"]
```

### `get_workbook_summary`

Lightweight structural summary of the entire workbook — sheet names, dimensions, detected data regions, and formula counts. **Call this first** to decide which sheets deserve deeper analysis.

| Argument | Type | Required | Description |
|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file |

```
get_workbook_summary(file_path="/data/report.xlsx")
→ { "sheets": [{ "name": "Sales", "max_row": 500, "max_column": 10,
     "region_count": 2, "formula_count": 48 }, ...] }
```

### `get_sheet_formulas`

**The primary analysis tool.** Extract every formula from a sheet — cell address, raw Excel formula string, and last-cached value. Best for tracing cross-sheet references and nested derived quantities.

| Argument | Type | Required | Description |
|---|---|---|---|
| `file_path` | string | yes | Absolute path to the `.xlsx` file |
| `sheet_name` | string | yes | Exact sheet name (case-sensitive) |

```
get_sheet_formulas(file_path="/data/report.xlsx", sheet_name="Summary")
→ [{"address": "B4", "formula": "=SUM(B2:B3)", "cached_value": 4500},
   {"address": "B6", "formula": "=Sales!D500-B4", "cached_value": 120500}]
```

### `get_sheet_data`

Read a sheet and return its data (including formulas) in the requested format. Large sheets are automatically sampled.

| Argument | Type | Required | Default | Description |
|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file |
| `sheet_name` | string | yes | — | Exact sheet name (case-sensitive) |
| `format` | string | no | `"markdown"` | `"markdown"`, `"json"`, or `"xml"` |
| `max_sample_rows` | int | no | `100` | Max rows per data region |
| `full` | bool | no | `false` | Disable sampling and return every row |

```
# Sampled Markdown (default)
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")

# Full load of a small lookup table
get_sheet_data(file_path="/data/report.xlsx", sheet_name="Lookups", full=True)
```

### `get_sheet_sample`

Very small preview (default 30 rows) — useful for a quick glance before requesting more data.

| Argument | Type | Required | Default | Description |
|---|---|---|---|---|
| `file_path` | string | yes | — | Absolute path to the `.xlsx` file |
| `sheet_name` | string | yes | — | Exact sheet name (case-sensitive) |
| `max_rows` | int | no | `30` | Max rows to include |
| `format` | string | no | `"markdown"` | Output format |

```
get_sheet_sample(file_path="/data/report.xlsx", sheet_name="Sales", max_rows=10)
```

---

## Recommended Tool Combinations

The workflows below cover the most common scenarios. In every case, **formulas come first** — they are the primary source of insight into what a workbook does.

### 1. Full Workbook Analysis (business summary)

The standard workflow. Formulas are retrieved **before** values so the LLM understands the calculation logic first.

```
Step 1 → get_workbook_summary(file_path="/data/report.xlsx")
         Understand sheet sizes and formula counts.

Step 2 → get_sheet_formulas(file_path="/data/report.xlsx", sheet_name="Sales")
         *** Call this FIRST for each sheet. ***
         Pull the complete formula list — cross-sheet references, nested
         calculations, everything.  This is the most important step.

Step 3 → get_sheet_data(file_path="/data/report.xlsx", sheet_name="Sales")
         Now get the surrounding values for context.  With the formulas
         already in hand, the LLM can map values to the logic.

Repeat Steps 2–3 for each sheet with formulas.
```

**Prompt after gathering data:**

> "Based on the formulas and sheet data above, write a business summary of
> what this workbook calculates.  For each sheet explain: purpose, key
> inputs, key outputs, and how formulas derive the outputs."

---

### 2. Important-Sheet Deep Dive

When the user explicitly says a particular sheet is important, give it the most thorough treatment:

- **Small sheet** (total cells ≤ ~5 000, e.g. 100 rows × 50 cols): load the whole sheet so nothing is missed.
- **Large sheet** (total cells > ~5 000): prefer a **column-based approach** instead of random sampling, because a vertical slice gives complete, contiguous data for every line item and eliminates the randomness of scattered samples.

```
Step 1 → get_workbook_summary(file_path="/data/model.xlsx")
         Check the sheet dimensions (max_row × max_column).

Step 2 → get_sheet_formulas(file_path="/data/model.xlsx",
                            sheet_name="P&L")
         *** Always formulas first. ***

Step 3a (small sheet — total cells ≤ ~5 000):
  → get_sheet_data(file_path="/data/model.xlsx", sheet_name="P&L",
                   full=True)
    Load every row so the full picture is available.

Step 3b (large sheet — total cells > ~5 000):
  → get_sheet_data(file_path="/data/model.xlsx", sheet_name="P&L",
                   max_sample_rows=200)
    Use a generous sample budget, and focus on specific columns of
    interest (see "Column-Based Sampling" below) rather than a random
    scatter of rows.
```

> **Only apply this deep-dive mode when the user has flagged the sheet as important.** For all other sheets, the default smart sampling is sufficient.

---

### 3. Column-Based Sampling (vertical slices)

Many financial models are laid out with **line items (rows) × time periods (columns)**. In these sheets:

- Formulas are typically the same across every time-period column (e.g. `=B5*C5` repeats as `=C5*D5`, `=D5*E5`, …).
- A single vertical slice (one or two columns) therefore captures **all the unique formulas** for every line item.

**When to use this approach:**

1. You ran a random sample and want to confirm the **complete set of formulas for every line item** — a vertical slice is more reliable than scattered rows because it covers every row in a contiguous block.
2. Dates or time periods run along the columns — the formulas are repetitive across time, so one column tells the whole story.

```
Step 1 → get_sheet_formulas(file_path="/data/forecast.xlsx",
                            sheet_name="Revenue")
         Get ALL formulas first.  Inspect them to confirm the pattern
         repeats across columns (e.g. same formula in C5, D5, E5, …).

Step 2 → get_sheet_data(file_path="/data/forecast.xlsx",
                        sheet_name="Revenue", max_sample_rows=200)
         Review the sampled data.  Because formulas repeat horizontally,
         a single representative time-period column already contains the
         full logic.  The LLM can note the pattern and generalise.

Step 3 → Ask the LLM:
         "The formulas repeat across columns (time periods).  Using
          column B as the representative slice, explain what each line
          item calculates and how it flows into downstream sheets."
```

> **Contrast with random sampling:** random sampling picks rows from the head, middle, and tail — good for general exploration. Column-based sampling picks a vertical slice — better when you need all formulas for every line item without gaps.

---

### 4. Formula-Focused Analysis (trace calculation chains)

When the goal is purely to understand logic, not data:

```
Step 1 → get_workbook_summary(file_path="/data/budget.xlsx")
         Identify sheets with formulas (formula_count > 0).

Step 2 → get_sheet_formulas(file_path="/data/budget.xlsx",
                            sheet_name="Assumptions")
Step 3 → get_sheet_formulas(file_path="/data/budget.xlsx",
                            sheet_name="P&L")
         Pull formulas from every relevant sheet.

Step 4 → Ask the LLM:
         "Trace the cross-sheet formula references.  Starting from
          Assumptions, explain how each derived quantity flows into P&L."
```

---

### 5. Quick Preview of a Large File

When the file is very large and you want to minimise context usage:

```
Step 1 → list_sheets(file_path="/data/big_model.xlsx")

Step 2 → get_sheet_sample(file_path="/data/big_model.xlsx",
                          sheet_name="Revenue", max_rows=20)
         Glance at structure and column headers.

Step 3 → get_sheet_formulas(file_path="/data/big_model.xlsx",
                            sheet_name="Revenue")
         Even for a quick preview, formulas first.

Step 4 → get_sheet_data(file_path="/data/big_model.xlsx",
                        sheet_name="Revenue", max_sample_rows=50)
         Bigger sample with values for additional context.
```

---

## When to Use Each Mode

| Mode | Flag / Tool | When to use |
|---|---|---|
| **Smart sampling** (default) | `get_sheet_data(...)` | General-purpose exploration of any sheet. Automatically includes headers, formula rows, and head/middle/tail data rows. |
| **Full load** | `get_sheet_data(..., full=True)` | Small sheets (≤ ~5 000 cells) or sheets the user has flagged as important and small enough to fit. **Use with caution on large sheets.** |
| **Column-based approach** | `get_sheet_formulas` + targeted `get_sheet_data` | Sheets where dates/periods run across columns and formulas repeat horizontally. A vertical slice captures all unique formulas per line item. Also preferred for large important sheets where random sampling is too sparse. |
| **Tiny preview** | `get_sheet_sample(..., max_rows=10)` | Quick glance at column headers and structure before deciding on a deeper read. |
| **Formulas only** | `get_sheet_formulas(...)` | When you only need the calculation logic, not the data values. Always the first call for any sheet you plan to analyse. |

---

## Handling Large Files & Unstructured Sheets

Sheets with more than **100 rows per data region** are automatically sampled. The sampling strategy always includes header rows, formula rows (up to half the budget), and evenly spread head / middle / tail data rows.

| Goal | How |
|---|---|
| Bigger sample | `get_sheet_data(..., max_sample_rows=200)` |
| Smaller sample | `get_sheet_data(..., max_sample_rows=30)` or `get_sheet_sample` |
| Disable sampling | `get_sheet_data(..., full=True)` — **use with caution** |

Real-world files are often messy (banners, blank rows, headers not on row 1, multiple data patches). The server detects contiguous non-blank row runs and treats each as an independent **data region** with its own headers, rows, and formulas.

---

## Output Formats

| Format | Best for |
|---|---|
| `"markdown"` | Human-readable summaries, conversational use |
| `"json"` | Programmatic processing, downstream scripts |
| `"xml"` | Structured interchange, XML-consuming systems |

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `FileNotFoundError` | Wrong or relative path | Use an absolute path, e.g. `"/home/user/file.xlsx"` |
| `KeyError: '<sheet>'` | Sheet name misspelled | Call `list_sheets` first to get exact names |
| Output too large | Thousands of rows + `full=True` | Remove `full=True` or reduce `max_sample_rows` |
| `cached_value` is `null` | File never opened in Excel after formulas were written | Open and save in Excel, then re-run |
