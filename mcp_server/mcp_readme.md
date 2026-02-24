# Excel Analysis MCP Server

An MCP (Model Context Protocol) server that reads `.xlsx` workbooks and exposes their content — including formulas and cached values — to LLM clients such as GitHub Copilot.

## Quick Start

```bash
cd mcp_server
pip install -r requirements.txt
python server.py          # starts the stdio MCP server
```

### Connecting from an MCP Client

Copy (or symlink) `mcp.json` into your client's configuration directory.
For **VS Code / GitHub Copilot** add the server entry to your workspace `.vscode/mcp.json`:

```jsonc
{
  "mcpServers": {
    "excel-analysis": {
      "command": "python",
      "args": ["<absolute-path-to>/mcp_server/server.py"],
      "cwd": "<absolute-path-to>/mcp_server"
    }
  }
}
```

---

## Tools

| Tool | Purpose |
|---|---|
| `list_sheets` | List all sheet names in a workbook. |
| `get_workbook_summary` | Lightweight overview — sheet dimensions, region count, formula count. |
| `get_sheet_data` | Full sheet data (auto-sampled) as **markdown**, **json**, or **xml**. |
| `get_sheet_formulas` | Extract every formula from a sheet with cached values. |
| `get_sheet_sample` | Quick small preview (default 30 rows) of a sheet. |

---

## Recommended Tool Combinations

### 1. Full Workbook Analysis

```
Step 1 → get_workbook_summary(file_path)
         Understand sheets, sizes, formula counts.

Step 2 → get_sheet_data(file_path, sheet_name, format="markdown")
         For each sheet (or the interesting ones).

Step 3 → get_sheet_formulas(file_path, sheet_name)
         Dive deeper into formula logic when cross-sheet
         references or deeply nested formulas need explaining.
```

### 2. Quick Preview of a Large File

```
Step 1 → list_sheets(file_path)

Step 2 → get_sheet_sample(file_path, sheet_name, max_rows=20)
         Glance at what the sheet contains.

Step 3 → get_sheet_data(file_path, sheet_name, max_sample_rows=50)
         Get a bigger sample with formulas.
```

### 3. Formula-Focused Analysis

```
Step 1 → get_workbook_summary(file_path)

Step 2 → get_sheet_formulas(file_path, sheet_name)
         For every sheet that has formulas.
         Ask the LLM to trace cross-sheet references and
         explain derived quantities.
```

---

## Handling Large Files

Sheets with more than **100 rows per data region** are automatically sampled.
The sampling strategy:

1. **Header rows** — always included.
2. **Formula rows** — always included (up to a budget).
3. **Head / middle / tail** data rows — evenly spread.

You can control the sample size:

- `get_sheet_data(..., max_sample_rows=200)` — larger sample.
- `get_sheet_data(..., full=True)` — disable sampling entirely (use with caution on huge files).

## Unstructured Sheets

The server auto-detects rectangular **data regions** — even when headers are not on the first row, or when blank rows and company banners sit above the actual data.  Each region is reported separately with its own headers.
