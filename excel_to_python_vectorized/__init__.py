"""Excel-to-Python Vectorized Converter.

Converts Excel workbooks to vectorized Python scripts that are:
- Smaller in size (repeated formulas become loops)
- Faster to execute (batch operations)
- Easier to read and modify

Also supports:
- Cross-sheet references
- External workbook references (via input_files_config.json)
- Generates an analysis report
"""

from .converter import convert_excel_to_python_vectorized

__all__ = ["convert_excel_to_python_vectorized"]
