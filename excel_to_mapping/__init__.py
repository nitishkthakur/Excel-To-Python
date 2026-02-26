"""Excel-to-Mapping Converter.

Analyses Excel workbooks and produces a **tabular** mapping report that
classifies each sheet's cells into three categories via a ``Type`` column:

  * **Input** – hardcoded (non-formula) values.
  * **Calculation** – formula cells (intermediate results referenced by
    other formulas), with dragged formulas collapsed into groups.
  * **Output** – formula cells whose results are *not* referenced by any
    other formula (terminal / final results).

An ``IncludeFlag`` column (default ``True``) lets a human reviewer mark
which rows to keep.  The companion :mod:`regenerator` module reads the
mapping report back and reconstructs an Excel workbook that honours the
reviewer's selections.
"""

from .mapper import generate_mapping_report
from .regenerator import regenerate_workbook, generate_input_template

__all__ = [
    "generate_mapping_report",
    "regenerate_workbook",
    "generate_input_template",
]
