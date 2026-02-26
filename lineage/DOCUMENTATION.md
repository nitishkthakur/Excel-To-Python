# Lineage Module — Documentation

## Overview

The **lineage** module extracts data-flow lineage from Excel workbooks and produces both structured Excel reports and visual graph images. It builds on the **smart formula sampler** (`mcp_server/smart_formula_sampler.py`) to de-duplicate dragged formulas — so a column with 10 000 identical `=B{R}*C{R}` rows appears as a single pattern rather than 10 000 separate entries.

Two levels of lineage are produced:

| Level | Description | Output |
|-------|-------------|--------|
| **Simple** | Sheet-level view: input columns, unique calculation patterns, output columns, cross-sheet data flow. | `<name>_simple_lineage.xlsx` + PNG graph |
| **Complex** | Column-level view: every unique formula pattern with its dependencies, cross-sheet and external-file references, column-to-column dependency edges. | `<name>_complex_lineage.xlsx` + PNG graph |

---

## Modules

### `lineage_builder.py`

Core extraction logic.

| Function | Purpose |
|----------|---------|
| `build_simple_lineage(path)` | Scan the workbook, classify cells as inputs/calculations/outputs at the column level, detect cross-sheet edges. |
| `build_complex_lineage(path)` | Same scan, but preserves every unique formula pattern per column with full dependency tracking. |
| `write_simple_lineage(lineage, output_path)` | Persist the simple lineage dict to an Excel file with sheets: Overview, Inputs, Calculations, Outputs, Cross-Sheet Edges. |
| `write_complex_lineage(lineage, output_path)` | Persist the complex lineage dict with sheets: Summary, All Patterns, Dependency Edges, Cross-Sheet Refs, External Refs. |

**Key design decisions:**

* **Inherits from the smart formula sampler** — `normalise_formula()` and `deduplicate_workbook_formulas()` are reused so that dragged formulas are collapsed into canonical patterns.
* **Inputs** = value cells that are referenced by at least one formula.
* **Outputs** = formula cells that are *not* referenced by any other formula (terminal calculations).
* **Cross-sheet edges** are detected by parsing formula references and checking if the target sheet differs from the source.

### `lineage_graph.py`

Reads the lineage Excel files and renders them as PNG graphs using `networkx` and `matplotlib`.

| Function | Purpose |
|----------|---------|
| `render_simple_graph(excel_path, output_png)` | Sheet-level graph. Each node = one sheet; edges = cross-sheet data flow. |
| `render_complex_graph(excel_path, output_png, max_nodes=80)` | Column-level graph. Nodes = `Sheet!Column`; edges = formula dependencies. Automatically prunes to keep the graph readable when there are many nodes. |

**Visual conventions:**
* **Green** nodes = inputs
* **Orange** nodes = calculations
* **Salmon** nodes = outputs
* Labels are word-wrapped to fit inside nodes
* Multipartite layout separates inputs → calculations → outputs into layers

---

## Usage

### CLI

```bash
# Build lineage Excel files
python -m lineage.lineage_builder Indigo.xlsx output_dir/

# Render graphs from the Excel files
python -m lineage.lineage_graph simple output_dir/Indigo_simple_lineage.xlsx simple.png
python -m lineage.lineage_graph complex output_dir/Indigo_complex_lineage.xlsx complex.png
```

### Python API

```python
from lineage.lineage_builder import (
    build_simple_lineage, write_simple_lineage,
    build_complex_lineage, write_complex_lineage,
)
from lineage.lineage_graph import render_simple_graph, render_complex_graph

# Build and save simple lineage
simple = build_simple_lineage("Indigo.xlsx")
write_simple_lineage(simple, "simple_lineage.xlsx")
render_simple_graph("simple_lineage.xlsx", "simple_lineage.png")

# Build and save complex lineage
cplx = build_complex_lineage("Indigo.xlsx")
write_complex_lineage(cplx, "complex_lineage.xlsx")
render_complex_graph("complex_lineage.xlsx", "complex_lineage.png")
```

---

## Indigo.xlsx Results

When run on the included `Indigo.xlsx` (Interglobe Aviation financial model):

| Metric | Value |
|--------|-------|
| Sheets | 13 |
| Total formula cells | 3 938 |
| Unique formula patterns | ~3 380 |
| Cross-sheet reference pairs | 32 |
| Input columns (referenced values) | 147 |
| Output columns (unreferenced formulas) | 149 |

The simple lineage graph shows a sheet-level data-flow diagram with the Valuation sheet as the primary output and Assumptions Sheet as the primary input. The complex lineage graph highlights the most-connected cross-sheet column dependencies.

---

## Dependencies

* `openpyxl` ≥ 3.1.0 (workbook I/O)
* `networkx` ≥ 3.0 (graph construction)
* `matplotlib` ≥ 3.5 (graph rendering)

---

## Tests

Tests are in `tests/test_lineage.py` and cover:
* Simple/complex lineage builder output structure
* Excel file generation with correct sheet names
* Graph PNG rendering (non-empty output)

Run with:

```bash
python -m pytest tests/test_lineage.py -v
```
