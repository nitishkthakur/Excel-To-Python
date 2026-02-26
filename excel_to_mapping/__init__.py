"""Excel-to-Mapping Converter.

Analyses Excel workbooks and produces a mapping report that separates
each sheet's cells into three categories:

  * **Inputs** – hardcoded (non-formula) values.
  * **Calculations** – formula cells, with continuous dragged formulas
    collapsed into compact group representations.
  * **Outputs** – formula cells whose results are *not* referenced by
    any other formula (i.e. terminal / final results).

One output sheet is generated per input sheet.
"""

from .mapper import generate_mapping_report

__all__ = ["generate_mapping_report"]
