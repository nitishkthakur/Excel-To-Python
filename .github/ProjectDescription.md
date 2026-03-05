Write a project description for the following project:
1. you will be given a bunch of excel files. these files are financial files with 10 - 100 sheets primarily. Each sheet has tabular patches, and small areas of information - like value_name, value pairs scattered across the sheet. Some patches in the same sheet may have dates as columns, or dates as rows or no dates - just 2 categorical entities as headers and row indices. It may just be headers and nothing as row index - like a series which has no index - with thousands of rows of hardcoded data and / or formulas.

The objectives are as follows:



















## Build an excel mcp server with 
I want you to build an excel mcp server. It will consist of multiple different ways of sampling from an excel file - each sampling method, capable of doing different things and meant for a different purpose. However, each sampling method will have the following common features:
1. It can read both formulas and values
2. It can read from one sheet or multiple sheets

There will be certain generic functions like:
1. get_sheet_names() - returns the names of all sheets in the file

## Retrievers - each retriever returns values as markdown to the LLM.

### Retriever 1 - used for fetching ALL DATA from a sheet or multiple sheets. 
1. Identify the limits of the sheet and simply fetch all data and return. 
2. Make sure this is fast - using libraries which can read large amounts of data fast (this must hold true for every code you write here)

### Retriever 2 - used for fetching representative data from a sheet or multiple sheets - from each "patch" within a sheet.
1. Identify patches within a sheet. 
2. Identify non patches within a sheet which have data - like value_name, value pairs.
3. If the patch is Too wide or too long, then fetch the header row and column, and 5 rows and columns of data from the top left corner of the patch. By default, fetch the entire patch.
4. If truncating the patch, do it in the following way: The user sets a threshold for the maximum number of rows and columns to fetch from a patch. If the patch exceeds this threshold, apply the truncation method described above. But fetch both the top threshold rows, and the bottom threshold rows, and the left threshold columns and the right threshold columns. This way, we can capture any patterns that may exist in the top or bottom rows, or the left or right columns of the patch. default threshold is 20 rows and 20 columns. The user can adjust this threshold as needed - separately for rows and columns.
5. The returned markdowns must be returned as a dictionary with the sheet_name, patch_range(of the original patch - not the truncated patch), and the markdown content.
6. This needs to be as fast as possible. Really think on how to make this fast.

Tool docs:
1. Write documentation for each tool - explaining primarily when to use the tool.
2. For each argument, write a one line documentation inline Using Annotated[--, Field(description="--")]
3. Provide a few examples in te tool description or docstring.
---

------------------------
# Task: Build an Excel MCP Server

## Context

You are working with large financial Excel files with the following structure:

- **10–100 sheets** per file
- Each sheet contains a mix of:
  - **Tabular patches** — contiguous rectangular data blocks. Column/row headers may be dates, categorical labels, or absent entirely (e.g. flat series with thousands of rows)
  - **Scattered key-value pairs** — isolated `label | value` cells sitting outside any table
- Cells may contain **hardcoded values** or **Excel formulas**

---

## What to Build

A Python **MCP server** that exposes Excel file content to LLMs via structured retrieval tools. The server must implement the tools described below.

### Global Constraints

- All tools must support reading both **raw formulas** and **computed values** (parameter-controlled; default: formulas)
- All tools must accept either a **single sheet name** or a **list of sheet names**
- **Performance is non-negotiable** — use the fastest available libraries (e.g. `openpyxl` with `read_only=True`, or `calamine`/`fastexcel` for value-only reads). Justify any choice that affects throughput
- All tools return content as **Markdown**

---

## Tool Documentation Requirements

Every tool must follow this pattern:

```python
from typing import Annotated
from pydantic import Field

def tool_name(
    file_path: Annotated[str, Field(description="Absolute path to the .xlsx file")],
    sheets:    Annotated[str | list[str], Field(description="Sheet name or list of sheet names to read")],
    formulas:  Annotated[bool, Field(description="Return raw formulas if True, computed values if False. Default: True")],
    # ... other params
) -> ...:
    """
    One-line summary of what the tool does.

    **When to use:**
    Describe the primary use case. Explain when to prefer this tool over others.
    Note any performance trade-offs (e.g. 'use for small sheets only', 'O(cells) scan').

    **Examples:**
    # Example 1 — common case
    result = tool_name('financials.xlsx', 'Income_Statement')

    # Example 2 — multiple sheets, values only
    result = tool_name('financials.xlsx', ['P&L', 'Balance_Sheet'], formulas=False)

    # Example 3 — show expected output shape
    # Returns: { 'sheet_name': '...', 'patch_range': 'B3:M47', 'markdown': '...' }
    """
```

**Rules:**
- Parameter descriptions: one line, include valid values and default behaviour
- `**When to use:**` block is mandatory in every docstring
- Include 2–3 `**Examples:**` covering common usage, optional params, and output shape

---

## Generic Utilities

| Tool | Description |
|------|-------------|
| `get_sheet_names(file_path)` | Returns all sheet names in the workbook as a list of strings |

---

## Retrievers

### Retriever 1 — Full Sheet Fetch (`fetch_full`)

**When to use:** When you need the complete, unsampled contents of a sheet — e.g. for thorough analysis or small sheets where truncation would lose context. Expect higher latency on large sheets.

**Behaviour:**
1. Detect the used range of the sheet (min/max row and column bounds)
2. Read all cells within that range
3. Return as a single Markdown table per sheet

**Returns:** `dict[sheet_name, markdown_string]`

---

### Retriever 2 — Patch-Aware Sample Fetch (`fetch_patches`)

**When to use:** When exploring a large or unfamiliar sheet. This tool maps every data region in the sheet individually, giving the LLM a structured overview without overwhelming token budgets. Prefer this over `fetch_full` for sheets with more than ~500 rows.

**Behaviour:**

1. **Detect patches** — find all contiguous rectangular data blocks via a fast sparse scan
2. **Detect isolated key-value pairs** — find `label | value` cells that sit outside any patch
3. **Return each patch in full** if it fits within the row/column thresholds
4. **Truncate large patches** as follows:
   - Default thresholds: `row_threshold=20`, `col_threshold=20` (caller-configurable independently)
   - Always include: the header row and index column of the patch
   - Include: the **top N** and **bottom N** data rows (N = `row_threshold`)
   - Include: the **left M** and **right M** data columns (M = `col_threshold`)
   - This dual-sided slice captures structure at both ends of wide/tall tables

**Returns:** A list of dicts, one per patch (and one per isolated key-value region):

```python
{
    "sheet_name":  str,   # Source sheet name
    "patch_range": str,   # Full original range before truncation, e.g. "B3:M47"
    "markdown":    str,   # Markdown table of the (possibly truncated) patch
}
```

**Performance note:** Patch detection must be vectorised — load the used-range data into a 2-D array first, then apply a connected-components or bounding-box scan. Avoid cell-by-cell openpyxl iteration.