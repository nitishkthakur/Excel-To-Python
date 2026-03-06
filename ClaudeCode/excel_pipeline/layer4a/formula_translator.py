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
        # Date functions
        'EOMONTH': 'xl_eomonth',
        'YEAR': 'xl_year',
        'MONTH': 'xl_month',
        'DAY': 'xl_day',
        'DATE': 'xl_date',
        'TODAY': 'xl_today',
        'DAYS': 'xl_days',
        'EDATE': 'xl_edate',
        # Text / other
        'LEN': 'len',
        'IFERROR': 'xl_iferror',
        'SUMIF': 'xl_sumif',
        'SUMPRODUCT': 'xl_sumproduct',
        'COUNTA': 'xl_counta',
        'ISNUMBER': 'xl_isnumber',
        'ISBLANK': 'xl_isblank',
        'CONCATENATE': 'xl_concatenate',
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

    def _expand_ranges(self, formula: str, sheet: str) -> str:
        """
        Pre-expand range references before individual cell substitution.

        Replaces A1:B3 with comma-separated cell references so that
        xl_sum(D7:D8) becomes xl_sum(D7, D8) → xl_sum(c.get(...), c.get(...)).
        Must be called before the cell-reference substitution pass.
        """
        from openpyxl.utils import column_index_from_string, get_column_letter

        def expand_range(match):
            start_col = match.group(1)
            start_row = int(match.group(2))
            end_col   = match.group(3)
            end_row   = int(match.group(4))
            start_idx = column_index_from_string(start_col)
            end_idx   = column_index_from_string(end_col)
            cells = []
            for ci in range(start_idx, end_idx + 1):
                col_letter = get_column_letter(ci)
                for ri in range(start_row, end_row + 1):
                    cells.append(f"{col_letter}{ri}")
            return ", ".join(cells)

        return re.sub(r'\b([A-Z]+)(\d+):([A-Z]+)(\d+)\b', expand_range, formula)

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
        # Pre-expand range references (e.g. D7:D8 → D7, D8) before cell substitution
        result = self._expand_ranges(formula, sheet)

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

    @staticmethod
    def _split_args(content: str) -> list:
        """
        Split a comma-separated argument string respecting nested parentheses.

        e.g. "c.get(('S', 'A', 1), 0)>0, -c.get(('S', 'A', 2), 0), 0"
             → ["c.get(('S', 'A', 1), 0)>0", "-c.get(('S', 'A', 2), 0)", "0"]
        """
        args, depth, buf = [], 0, []
        for ch in content:
            if ch == '(':
                depth += 1
                buf.append(ch)
            elif ch == ')':
                depth -= 1
                buf.append(ch)
            elif ch == ',' and depth == 0:
                args.append(''.join(buf).strip())
                buf = []
            else:
                buf.append(ch)
        if buf:
            args.append(''.join(buf).strip())
        return args

    def _translate_if_function(self, formula: str) -> str:
        """Translate IF( ) to Python ternary using a parenthesis-aware splitter."""
        result = formula
        # Iteratively replace all IF( ) calls from innermost outward
        # by scanning for 'IF(' and extracting the balanced-paren content.
        while 'IF(' in result:
            idx = result.find('IF(')
            if idx == -1:
                break
            # Find the matching closing paren
            depth, end = 0, idx + 3  # start after 'IF('
            for i, ch in enumerate(result[idx + 3:], start=idx + 3):
                if ch == '(':
                    depth += 1
                elif ch == ')':
                    if depth == 0:
                        end = i
                        break
                    depth -= 1
            inner = result[idx + 3: end]
            args = self._split_args(inner)
            if len(args) == 3:
                cond, true_val, false_val = args
                replacement = f"({true_val} if {cond} else {false_val})"
            elif len(args) == 2:
                cond, true_val = args
                replacement = f"({true_val} if {cond} else None)"
            else:
                # Can't parse — leave as xl_if call
                replacement = f"xl_if({inner})"
            result = result[:idx] + replacement + result[end + 1:]
        return result

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

def xl_iferror(value, value_if_error):
    """Excel IFERROR function."""
    try:
        return value
    except Exception:
        return value_if_error

def xl_sumif(rng, criteria, sum_range=None):
    """Excel SUMIF — simplified: sum values in rng equal to criteria."""
    if sum_range is None:
        sum_range = rng
    if not isinstance(rng, (list, tuple)):
        rng = [rng]
    if not isinstance(sum_range, (list, tuple)):
        sum_range = [sum_range]
    return sum(v for c2, v in zip(rng, sum_range)
               if isinstance(v, (int, float)) and c2 == criteria)

def xl_sumproduct(*arrays):
    """Excel SUMPRODUCT."""
    if not arrays:
        return 0
    arrays = [a if isinstance(a, (list, tuple)) else [a] for a in arrays]
    length = min(len(a) for a in arrays)
    total = 0
    for i in range(length):
        product = 1
        for a in arrays:
            v = a[i]
            product *= (v if isinstance(v, (int, float)) else 0)
        total += product
    return total

def xl_counta(*args):
    """Excel COUNTA — count non-empty values."""
    flat = []
    for a in args:
        if isinstance(a, (list, tuple)):
            flat.extend(a)
        else:
            flat.append(a)
    return sum(1 for x in flat if x is not None and x != "")

def xl_isnumber(value):
    """Excel ISNUMBER."""
    return isinstance(value, (int, float))

def xl_isblank(value):
    """Excel ISBLANK."""
    return value is None or value == ""

def xl_concatenate(*args):
    """Excel CONCATENATE."""
    return "".join(str(a) if a is not None else "" for a in args)

# ---- Date helpers ----
import calendar as _calendar
from datetime import date as _date, datetime as _datetime, timedelta as _timedelta

def _to_date(v):
    """Convert Excel serial number or date object to Python date."""
    if isinstance(v, (_date, _datetime)):
        return v if isinstance(v, _date) else v.date()
    if isinstance(v, (int, float)) and v > 0:
        return (_date(1899, 12, 30) + _timedelta(days=int(v)))
    return None

def xl_eomonth(start_date, months):
    """Excel EOMONTH — last day of month N months from start_date."""
    d = _to_date(start_date)
    if d is None:
        return None
    month = d.month - 1 + int(months)
    year = d.year + month // 12
    month = month % 12 + 1
    last_day = _calendar.monthrange(year, month)[1]
    return _date(year, month, last_day)

def xl_edate(start_date, months):
    """Excel EDATE — same day N months from start_date."""
    d = _to_date(start_date)
    if d is None:
        return None
    month = d.month - 1 + int(months)
    year = d.year + month // 12
    month = month % 12 + 1
    last_day = _calendar.monthrange(year, month)[1]
    return _date(year, month, min(d.day, last_day))

def xl_year(value):
    """Excel YEAR."""
    d = _to_date(value)
    return d.year if d else None

def xl_month(value):
    """Excel MONTH."""
    d = _to_date(value)
    return d.month if d else None

def xl_day(value):
    """Excel DAY."""
    d = _to_date(value)
    return d.day if d else None

def xl_date(year, month, day):
    """Excel DATE."""
    try:
        return _date(int(year), int(month), int(day))
    except Exception:
        return None

def xl_today():
    """Excel TODAY."""
    return _date.today()

def xl_days(end_date, start_date):
    """Excel DAYS — number of days between two dates."""
    e = _to_date(end_date)
    s = _to_date(start_date)
    if e is None or s is None:
        return None
    return (e - s).days
'''
