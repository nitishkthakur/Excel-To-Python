"""FormulaTranslator: Translate Excel formulas to Python/pandas expressions."""

import re
from typing import Optional, Dict, List, Tuple
from excel_pipeline.utils.logging_setup import get_logger

logger = get_logger(__name__)


class FormulaTranslator:
    """Translate Excel formulas to Python expressions."""

    # Vectorizable pattern detection (for pattern formulas with {col}/{row})
    VECTORIZABLE_PATTERNS = [
        # Simple arithmetic: =A{row}+B{row}
        (r'^=?([A-Z]+)\{row\}([+\-*/^])([A-Z]+)\{row\}$',
         lambda m: f"df['{m.group(1)}'] {m.group(2).replace('^', '**')} df['{m.group(3)}']"),

        # Scalar multiplication: =A{row}*2
        (r'^=?([A-Z]+)\{row\}\*([0-9.]+)$',
         lambda m: f"df['{m.group(1)}'] * {m.group(2)}"),

        # Scalar operation: =A{row}/100
        (r'^=?([A-Z]+)\{row\}([*/])([0-9.]+)$',
         lambda m: f"df['{m.group(1)}'] {m.group(2)} {m.group(3)}"),

        # Simple function: =ABS(A{row})
        (r'^=?ABS\(([A-Z]+)\{row\}\)$',
         lambda m: f"df['{m.group(1)}'].abs()"),

        # MAX/MIN: =MAX(A{row}, 0)
        (r'^=?MAX\(([A-Z]+)\{row\},\s*([0-9.]+)\)$',
         lambda m: f"df['{m.group(1)}'].clip(lower={m.group(2)})"),

        (r'^=?MIN\(([A-Z]+)\{row\},\s*([0-9.]+)\)$',
         lambda m: f"df['{m.group(1)}'].clip(upper={m.group(2)})"),
    ]

    # Excel function mapping
    EXCEL_FUNCTIONS = {
        'SUM': 'xl_sum',
        'AVERAGE': 'xl_average',
        'COUNT': 'xl_count',
        'MIN': 'min',
        'MAX': 'max',
        'ABS': 'abs',
        'ROUND': 'round',
        'IF': None,  # Special handling
        'AND': 'all',
        'OR': 'any',
        'NOT': 'not',
    }

    def __init__(self):
        """Initialize FormulaTranslator."""
        pass

    def translate_formula(
        self,
        formula: str,
        sheet: str,
        is_vectorized: bool = False,
        row_var: Optional[str] = None
    ) -> str:
        """
        Translate Excel formula to Python expression.

        Args:
            formula: Excel formula (may start with =)
            sheet: Sheet name for this formula
            is_vectorized: If True, use pandas DataFrame syntax
            row_var: Variable name for row iteration (e.g., "_r")

        Returns:
            Python expression string

        Examples:
            translate_formula("=A1+B1", "Sheet1", False)
            → "c.get(('Sheet1', 'A', 1), 0) + c.get(('Sheet1', 'B', 1), 0)"

            translate_formula("=A{row}+B{row}", "Sheet1", True)
            → "df['A'] + df['B']"
        """
        if not formula:
            return "None"

        # Remove leading apostrophe and equals sign
        formula = formula.lstrip("'").lstrip("=")

        if is_vectorized:
            return self._translate_vectorized(formula, sheet)
        else:
            return self._translate_cell_by_cell(formula, sheet, row_var)

    def _translate_vectorized(self, formula: str, sheet: str) -> str:
        """
        Translate pattern formula to pandas DataFrame operation.

        Args:
            formula: Pattern formula with {row} or {col} placeholders
            sheet: Sheet name

        Returns:
            Pandas expression
        """
        # Try each vectorizable pattern
        for pattern, translator in self.VECTORIZABLE_PATTERNS:
            match = re.match(pattern, formula)
            if match:
                try:
                    result = translator(match)
                    logger.debug(f"Vectorized: {formula} → {result}")
                    return result
                except Exception as e:
                    logger.warning(f"Pattern match failed for {formula}: {e}")
                    continue

        # Check for IF function (special handling with np.where)
        if_match = re.match(r'^IF\(([^,]+),([^,]+),([^)]+)\)$', formula)
        if if_match:
            cond = self._translate_vectorized(if_match.group(1).strip(), sheet)
            true_val = self._translate_vectorized(if_match.group(2).strip(), sheet)
            false_val = self._translate_vectorized(if_match.group(3).strip(), sheet)
            return f"np.where({cond}, {true_val}, {false_val})"

        # Fallback: can't vectorize - return as-is with note
        logger.warning(f"Cannot vectorize formula: {formula}")
        return f"# TODO: Cannot vectorize: {formula}"

    def _translate_cell_by_cell(self, formula: str, sheet: str, row_var: Optional[str] = None) -> str:
        """
        Translate formula for cell-by-cell evaluation.

        Args:
            formula: Excel formula
            sheet: Sheet name
            row_var: Optional row variable for loops

        Returns:
            Python expression using c.get() for cell access
        """
        # Replace cell references with c.get() calls
        result = formula

        # Handle cross-sheet references: Sheet!A1 or 'Sheet Name'!A1
        result = re.sub(
            r"'([^']+)'!([A-Z]+)(\d+)",
            lambda m: f"c.get(('{m.group(1)}', '{m.group(2)}', {m.group(3)}), 0)",
            result
        )

        result = re.sub(
            r"([A-Za-z0-9_]+)!([A-Z]+)(\d+)",
            lambda m: f"c.get(('{m.group(1)}', '{m.group(2)}', {m.group(3)}), 0)",
            result
        )

        # Handle same-sheet references: A1, B2, etc.
        result = re.sub(
            r"\b([A-Z]+)(\d+)\b",
            lambda m: f"c.get(('{sheet}', '{m.group(1)}', {m.group(2)}), 0)",
            result
        )

        # Replace operators
        result = result.replace('^', '**')  # Power operator

        # Replace Excel functions
        for excel_func, python_func in self.EXCEL_FUNCTIONS.items():
            if python_func and excel_func in result:
                result = result.replace(f'{excel_func}(', f'{python_func}(')

        # Handle IF function specially
        if 'IF(' in result:
            result = self._translate_if_function(result)

        return result

    def _translate_if_function(self, formula: str) -> str:
        """Translate IF function to Python ternary or xl_if()."""
        # Simple IF: IF(cond, true_val, false_val) → (true_val if cond else false_val)
        # For complex cases, use helper function

        # Try to extract IF arguments
        if_pattern = r'IF\(([^,]+),([^,]+),([^)]+)\)'
        match = re.search(if_pattern, formula)

        if match:
            cond = match.group(1).strip()
            true_val = match.group(2).strip()
            false_val = match.group(3).strip()

            # Use ternary if simple, otherwise use helper
            if len(cond) < 50:  # Simple condition
                replacement = f"({true_val} if {cond} else {false_val})"
                formula = formula.replace(match.group(0), replacement)

        return formula

    def detect_vectorization_complexity(self, formula: str) -> str:
        """
        Analyze formula complexity for vectorization decision.

        Returns:
            "SIMPLE", "MODERATE", or "COMPLEX"
        """
        if not formula:
            return "SIMPLE"

        # Strip leading apostrophe and equals sign
        formula = formula.lstrip("'").lstrip("=")

        # Check for nested functions
        open_parens = formula.count('(')
        if open_parens > 3:
            return "COMPLEX"

        # Check for multiple IF statements
        if formula.count('IF(') > 2:
            return "COMPLEX"

        # Check for lookup functions
        if any(func in formula for func in ['VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'OFFSET', 'INDIRECT']):
            return "MODERATE"

        # Check for simple arithmetic
        if re.match(r'^[A-Z0-9+\-*/(). ]+$', formula.replace('{row}', '').replace('{col}', '')):
            return "SIMPLE"

        return "MODERATE"

    def extract_dependencies(self, formula: str, sheet: str) -> List[Tuple[str, str, int]]:
        """
        Extract cell references from formula.

        Args:
            formula: Excel formula
            sheet: Current sheet name

        Returns:
            List of (sheet, column, row) tuples
        """
        if not formula:
            return []

        dependencies = []

        # Cross-sheet references: 'Sheet Name'!A1 or Sheet!A1
        cross_sheet_refs = re.findall(r"'([^']+)'!([A-Z]+)(\d+)", formula)
        for ref_sheet, col, row in cross_sheet_refs:
            dependencies.append((ref_sheet, col, int(row)))

        cross_sheet_refs = re.findall(r"([A-Za-z0-9_]+)!([A-Z]+)(\d+)", formula)
        for ref_sheet, col, row in cross_sheet_refs:
            if not ref_sheet.startswith("'"):  # Avoid duplicates
                dependencies.append((ref_sheet, col, int(row)))

        # Same-sheet references: A1, B2, etc.
        same_sheet_refs = re.findall(r"\b([A-Z]+)(\d+)\b", formula)
        for col, row in same_sheet_refs:
            # Exclude if already captured as cross-sheet
            if not any(dep[1] == col and dep[2] == int(row) for dep in dependencies):
                dependencies.append((sheet, col, int(row)))

        return dependencies


# Helper functions for generated code
HELPER_FUNCTIONS_CODE = '''
def xl_sum(*args):
    """Excel SUM function."""
    flat = []
    for a in args:
        if isinstance(a, (list, tuple)):
            flat.extend(a)
        else:
            flat.append(a)
    return sum(x for x in flat if isinstance(x, (int, float)) and x is not None)

def xl_average(*args):
    """Excel AVERAGE function."""
    flat = []
    for a in args:
        if isinstance(a, (list, tuple)):
            flat.extend(a)
        else:
            flat.append(a)
    nums = [x for x in flat if isinstance(x, (int, float)) and x is not None]
    return sum(nums) / len(nums) if nums else 0

def xl_count(*args):
    """Excel COUNT function."""
    flat = []
    for a in args:
        if isinstance(a, (list, tuple)):
            flat.extend(a)
        else:
            flat.append(a)
    return len([x for x in flat if isinstance(x, (int, float)) and x is not None])

def xl_if(condition, true_value, false_value):
    """Excel IF function."""
    return true_value if condition else false_value
'''
