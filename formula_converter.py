"""
Converts Excel formulas to equivalent Python expressions.

Handles:
- Basic arithmetic (+, -, *, /, ^)
- Cell references (A1, $A$1, Sheet1!A1)
- Range references (A1:A10)
- Common Excel functions (SUM, AVERAGE, IF, VLOOKUP, etc.)
- Table structured references (Table1[Column])
- String concatenation (&)
"""

import re


# Maps Excel function names to Python helper function names
EXCEL_FUNCTION_MAP = {
    "SUM": "xl_sum",
    "AVERAGE": "xl_average",
    "COUNT": "xl_count",
    "COUNTA": "xl_counta",
    "MIN": "xl_min",
    "MAX": "xl_max",
    "IF": "xl_if",
    "AND": "xl_and",
    "OR": "xl_or",
    "NOT": "xl_not",
    "ABS": "abs",
    "ROUND": "round",
    "ROUNDUP": "xl_roundup",
    "ROUNDDOWN": "xl_rounddown",
    "INT": "int",
    "MOD": "xl_mod",
    "POWER": "xl_power",
    "SQRT": "xl_sqrt",
    "LEN": "xl_len",
    "LEFT": "xl_left",
    "RIGHT": "xl_right",
    "MID": "xl_mid",
    "UPPER": "xl_upper",
    "LOWER": "xl_lower",
    "TRIM": "xl_trim",
    "CONCATENATE": "xl_concatenate",
    "TEXT": "xl_text",
    "VALUE": "xl_value",
    "VLOOKUP": "xl_vlookup",
    "HLOOKUP": "xl_hlookup",
    "INDEX": "xl_index",
    "MATCH": "xl_match",
    "IFERROR": "xl_iferror",
    "ISBLANK": "xl_isblank",
    "SUMIF": "xl_sumif",
    "SUMIFS": "xl_sumifs",
    "COUNTIF": "xl_countif",
    "COUNTIFS": "xl_countifs",
    "AVERAGEIF": "xl_averageif",
    "SUMPRODUCT": "xl_sumproduct",
    "OFFSET": "xl_offset",
    "INDIRECT": "xl_indirect",
    "ROW": "xl_row",
    "COLUMN": "xl_column",
    "ROWS": "xl_rows",
    "COLUMNS": "xl_columns",
    "TODAY": "xl_today",
    "NOW": "xl_now",
    "YEAR": "xl_year",
    "MONTH": "xl_month",
    "DAY": "xl_day",
    "DATE": "xl_date",
    "EOMONTH": "xl_eomonth",
    "EDATE": "xl_edate",
    "DATEDIF": "xl_datedif",
    "PI": "xl_pi",
    "TRUE": "True",
    "FALSE": "False",
}

# Pattern for cell references: optional sheet prefix, column letters, row numbers
# Handles: A1, $A$1, Sheet1!A1, 'Sheet Name'!A1
CELL_REF_PATTERN = re.compile(
    r"(?:(?:'([^']+)'|(\w+))!)?"  # optional sheet reference
    r"\$?([A-Z]{1,3})\$?(\d+)"    # column and row
)

# Pattern for range references: A1:B10 or Sheet1!A1:B10
RANGE_REF_PATTERN = re.compile(
    r"(?:(?:'([^']+)'|(\w+))!)?"
    r"\$?([A-Z]{1,3})\$?(\d+)"
    r":"
    r"\$?([A-Z]{1,3})\$?(\d+)"
)

# Pattern for table structured references: TableName[ColumnName] or TableName[[#Headers],[Column]]
TABLE_REF_PATTERN = re.compile(
    r"(\w+)\[([^\]]*)\]"
)


def col_letter_to_index(col_str):
    """Convert column letter(s) to 1-based index. A=1, B=2, ..., Z=26, AA=27."""
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def index_to_col_letter(index):
    """Convert 1-based column index to letter(s). 1=A, 2=B, ..., 26=Z, 27=AA."""
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_to_var_name(sheet_name, col, row):
    """Convert a cell reference to a Python variable name."""
    safe_sheet = re.sub(r'[^a-zA-Z0-9]', '_', sheet_name)
    return f"s_{safe_sheet}_{col}{row}"


def range_to_var_name(sheet_name, col1, row1, col2, row2):
    """Convert a range reference to a Python variable name for a range helper."""
    safe_sheet = re.sub(r'[^a-zA-Z0-9]', '_', sheet_name)
    return f"rng_{safe_sheet}_{col1}{row1}_{col2}{row2}"


def table_ref_to_var_name(table_name, column_name):
    """Convert a table structured reference to a Python variable name."""
    safe_table = re.sub(r'[^a-zA-Z0-9]', '_', table_name)
    safe_col = re.sub(r'[^a-zA-Z0-9]', '_', column_name)
    return f"tbl_{safe_table}_{safe_col}"


class FormulaConverter:
    """Converts an Excel formula string to a Python expression."""

    def __init__(self, current_sheet, tables=None):
        """
        Args:
            current_sheet: Name of the sheet where the formula resides.
            tables: Dict mapping table names to their info
                    {table_name: {"sheet": str, "ref": str, "columns": [str], "header_row": int, "data_start_row": int, "data_end_row": int}}
        """
        self.current_sheet = current_sheet
        self.tables = tables or {}
        self.referenced_cells = set()   # (sheet, col, row) tuples
        self.referenced_ranges = set()  # (sheet, col1, row1, col2, row2) tuples
        self.referenced_tables = set()  # (table_name, column_name) tuples

    def convert(self, formula):
        """Convert an Excel formula to a Python expression string.

        Args:
            formula: Excel formula string (without leading '=')

        Returns:
            Python expression string
        """
        if formula.startswith("="):
            formula = formula[1:]

        result = self._convert_expression(formula)
        return result

    def _convert_expression(self, expr):
        """Recursively convert an expression."""
        expr = expr.strip()
        if not expr:
            return "''"

        result = []
        i = 0
        while i < len(expr):
            c = expr[i]

            # String literal
            if c == '"':
                end = expr.index('"', i + 1)
                result.append(expr[i:end + 1])
                i = end + 1
                continue

            # Excel string concatenation operator
            if c == '&':
                result.append(' + str(')
                # We need to close the str() after the next token
                i += 1
                # Find the next token and wrap it
                remaining = expr[i:].strip()
                token, consumed = self._parse_next_token(remaining)
                result.append(token + ')')
                i = i + (len(expr[i:]) - len(expr[i:].strip())) + consumed
                continue

            # Comparison operators
            if c == '<' and i + 1 < len(expr) and expr[i + 1] == '>':
                result.append(' != ')
                i += 2
                continue
            if c == '<' and i + 1 < len(expr) and expr[i + 1] == '=':
                result.append(' <= ')
                i += 2
                continue
            if c == '>' and i + 1 < len(expr) and expr[i + 1] == '=':
                result.append(' >= ')
                i += 2
                continue
            if c in '<>':
                result.append(f' {c} ')
                i += 1
                continue

            # Equal sign as comparison (in formula context)
            if c == '=':
                result.append(' == ')
                i += 1
                continue

            # Power operator
            if c == '^':
                result.append(' ** ')
                i += 1
                continue

            # Basic arithmetic
            if c in '+-*/':
                result.append(f' {c} ')
                i += 1
                continue

            # Parenthesized expression
            if c == '(':
                matching = self._find_matching_paren(expr, i)
                inner = expr[i + 1:matching]
                result.append('(' + self._convert_expression(inner) + ')')
                i = matching + 1
                continue

            # Comma (function argument separator)
            if c == ',':
                result.append(', ')
                i += 1
                continue

            # Whitespace
            if c == ' ':
                i += 1
                continue

            # Percent
            if c == '%':
                result.append(' / 100')
                i += 1
                continue

            # Start of token (function call, cell ref, number, etc.)
            token, consumed = self._parse_next_token(expr[i:])
            result.append(token)
            i += consumed

        return ''.join(result)

    def _parse_next_token(self, expr):
        """Parse the next token from the expression.

        Returns:
            (python_expression, chars_consumed)
        """
        expr_stripped = expr.lstrip()
        skip = len(expr) - len(expr_stripped)
        expr = expr_stripped

        if not expr:
            return ('', skip)

        # Number
        m = re.match(r'^(\d+\.?\d*(?:[eE][+-]?\d+)?)', expr)
        if m:
            return (m.group(1), skip + m.end())

        # String literal
        if expr[0] == '"':
            end = expr.index('"', 1)
            return (expr[:end + 1], skip + end + 1)

        # Boolean
        if expr.upper().startswith('TRUE'):
            if len(expr) == 4 or not expr[4].isalpha():
                return ('True', skip + 4)
        if expr.upper().startswith('FALSE'):
            if len(expr) == 5 or not expr[5].isalpha():
                return ('False', skip + 5)

        # Quoted sheet reference with range: 'Sheet Name'!A1:B2
        m = re.match(r"^'([^']+)'!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            sheet, c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
            var = range_to_var_name(sheet, c1, r1, c2, r2)
            self.referenced_ranges.add((sheet, c1, int(r1), c2, int(r2)))
            return (var, skip + m.end())

        # Quoted sheet reference with cell: 'Sheet Name'!A1
        m = re.match(r"^'([^']+)'!\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            sheet, col, row = m.group(1), m.group(2), m.group(3)
            var = cell_to_var_name(sheet, col, row)
            self.referenced_cells.add((sheet, col, int(row)))
            return (var, skip + m.end())

        # Unquoted sheet reference with range: Sheet1!A1:B2
        m = re.match(r"^(\w+)!\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            sheet, c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
            var = range_to_var_name(sheet, c1, r1, c2, r2)
            self.referenced_ranges.add((sheet, c1, int(r1), c2, int(r2)))
            return (var, skip + m.end())

        # Unquoted sheet reference with cell: Sheet1!A1
        m = re.match(r"^(\w+)!\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            sheet, col, row = m.group(1), m.group(2), m.group(3)
            var = cell_to_var_name(sheet, col, row)
            self.referenced_cells.add((sheet, col, int(row)))
            return (var, skip + m.end())

        # Table structured reference: TableName[Column]
        m = re.match(r"^(\w+)\[([^\]]*)\]", expr)
        if m:
            table_name, col_ref = m.group(1), m.group(2)
            if table_name in self.tables:
                var = table_ref_to_var_name(table_name, col_ref)
                self.referenced_tables.add((table_name, col_ref))
                return (var, skip + m.end())

        # Function call: FUNCNAME(...)
        m = re.match(r"^([A-Z][A-Z0-9_.]*)\(", expr, re.IGNORECASE)
        if m:
            func_name = m.group(1).upper()
            paren_start = m.end() - 1
            paren_end = self._find_matching_paren(expr, paren_start)
            args_str = expr[paren_start + 1:paren_end]

            # Convert arguments
            args = self._split_function_args(args_str)
            converted_args = [self._convert_expression(a) for a in args]

            py_func = EXCEL_FUNCTION_MAP.get(func_name, f"xl_{func_name.lower()}")

            # Special cases
            if func_name in ("TRUE",):
                return ("True", skip + paren_end + 1)
            if func_name in ("FALSE",):
                return ("False", skip + paren_end + 1)
            if func_name in ("PI",):
                return ("xl_pi()", skip + paren_end + 1)
            if func_name in ("TODAY", "NOW"):
                return (f"{py_func}()", skip + paren_end + 1)

            result = f"{py_func}({', '.join(converted_args)})"
            return (result, skip + paren_end + 1)

        # Range reference: A1:B2 (same sheet)
        m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            c1, r1, c2, r2 = m.group(1), m.group(2), m.group(3), m.group(4)
            var = range_to_var_name(self.current_sheet, c1, r1, c2, r2)
            self.referenced_ranges.add((self.current_sheet, c1, int(r1), c2, int(r2)))
            return (var, skip + m.end())

        # Cell reference: A1 (same sheet)
        m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)", expr)
        if m:
            col, row = m.group(1), m.group(2)
            var = cell_to_var_name(self.current_sheet, col, row)
            self.referenced_cells.add((self.current_sheet, col, int(row)))
            return (var, skip + m.end())

        # Fallback: consume one character
        return (expr[0], skip + 1)

    def _find_matching_paren(self, expr, start):
        """Find the matching closing parenthesis."""
        depth = 1
        i = start + 1
        in_string = False
        while i < len(expr):
            if expr[i] == '"' and not in_string:
                in_string = True
            elif expr[i] == '"' and in_string:
                in_string = False
            elif not in_string:
                if expr[i] == '(':
                    depth += 1
                elif expr[i] == ')':
                    depth -= 1
                    if depth == 0:
                        return i
            i += 1
        return len(expr) - 1

    def _split_function_args(self, args_str):
        """Split function arguments, respecting nested parentheses and strings."""
        args = []
        current = []
        depth = 0
        in_string = False

        for c in args_str:
            if c == '"':
                in_string = not in_string
                current.append(c)
            elif in_string:
                current.append(c)
            elif c == '(':
                depth += 1
                current.append(c)
            elif c == ')':
                depth -= 1
                current.append(c)
            elif c == ',' and depth == 0:
                args.append(''.join(current).strip())
                current = []
            else:
                current.append(c)

        if current:
            args.append(''.join(current).strip())

        return [a for a in args if a]


# The helper functions library that gets included in generated code
HELPER_FUNCTIONS_CODE = '''
import math
import datetime
from collections import defaultdict


def _to_num(val):
    """Convert a value to a number, returning 0 for non-numeric."""
    if val is None:
        return 0
    if isinstance(val, bool):
        return 1 if val else 0
    if isinstance(val, (int, float)):
        return val
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0


def _flatten(args):
    """Flatten nested lists/tuples into a single list."""
    result = []
    for a in args:
        if isinstance(a, (list, tuple)):
            result.extend(_flatten(a))
        else:
            result.append(a)
    return result


def xl_sum(*args):
    vals = _flatten(args)
    return sum(_to_num(v) for v in vals if v is not None)


def xl_average(*args):
    vals = [_to_num(v) for v in _flatten(args) if v is not None]
    return sum(vals) / len(vals) if vals else 0


def xl_count(*args):
    vals = _flatten(args)
    return sum(1 for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool))


def xl_counta(*args):
    vals = _flatten(args)
    return sum(1 for v in vals if v is not None and v != "")


def xl_min(*args):
    vals = [_to_num(v) for v in _flatten(args) if v is not None]
    return min(vals) if vals else 0


def xl_max(*args):
    vals = [_to_num(v) for v in _flatten(args) if v is not None]
    return max(vals) if vals else 0


def xl_if(condition, true_val, false_val=False):
    return true_val if condition else false_val


def xl_and(*args):
    return all(bool(a) for a in _flatten(args))


def xl_or(*args):
    return any(bool(a) for a in _flatten(args))


def xl_not(val):
    return not bool(val)


def xl_roundup(number, digits):
    import math
    factor = 10 ** int(digits)
    return math.ceil(number * factor) / factor


def xl_rounddown(number, digits):
    import math
    factor = 10 ** int(digits)
    return math.floor(number * factor) / factor


def xl_mod(number, divisor):
    return number % divisor


def xl_power(base, exp):
    return base ** exp


def xl_sqrt(val):
    return math.sqrt(_to_num(val))


def xl_len(text):
    return len(str(text)) if text is not None else 0


def xl_left(text, n=1):
    return str(text)[:int(n)]


def xl_right(text, n=1):
    return str(text)[-int(n):]


def xl_mid(text, start, length):
    s = str(text)
    return s[int(start) - 1:int(start) - 1 + int(length)]


def xl_upper(text):
    return str(text).upper()


def xl_lower(text):
    return str(text).lower()


def xl_trim(text):
    return " ".join(str(text).split())


def xl_concatenate(*args):
    return "".join(str(a) if a is not None else "" for a in args)


def xl_text(value, fmt):
    if isinstance(value, (int, float)):
        if "%" in str(fmt):
            return f"{value:.2%}"
        if "." in str(fmt):
            decimals = len(str(fmt).split(".")[-1].rstrip("0#"))
            return f"{value:.{decimals}f}"
        return str(value)
    return str(value)


def xl_value(text):
    try:
        return float(str(text).replace(",", ""))
    except ValueError:
        return 0


def xl_vlookup(lookup_val, table_range, col_index, approx=True):
    """VLOOKUP: search first column of table_range for lookup_val, return value in col_index."""
    col_idx = int(col_index) - 1
    if approx:
        last_match = None
        for row in table_range:
            if row[0] is not None and row[0] <= lookup_val:
                last_match = row
            elif row[0] is not None and row[0] > lookup_val:
                break
        if last_match and col_idx < len(last_match):
            return last_match[col_idx]
    else:
        for row in table_range:
            if row[0] == lookup_val:
                if col_idx < len(row):
                    return row[col_idx]
                return None
    return None


def xl_hlookup(lookup_val, table_range, row_index, approx=True):
    """HLOOKUP: search first row of table_range for lookup_val, return value in row_index."""
    row_idx = int(row_index) - 1
    if not table_range:
        return None
    header = table_range[0]
    if not approx:
        for ci, val in enumerate(header):
            if val == lookup_val:
                if row_idx < len(table_range):
                    return table_range[row_idx][ci]
    return None


def xl_index(array, row_num, col_num=None):
    """INDEX: return value at row_num, col_num in array."""
    r = int(row_num) - 1
    if col_num is not None:
        c = int(col_num) - 1
        if isinstance(array, list) and r < len(array):
            row = array[r]
            if isinstance(row, (list, tuple)) and c < len(row):
                return row[c]
            elif c == 0:
                return row
        return None
    if isinstance(array, list) and r < len(array):
        return array[r]
    return None


def xl_match(lookup_val, lookup_range, match_type=1):
    """MATCH: return position of lookup_val in lookup_range."""
    flat = _flatten(lookup_range) if isinstance(lookup_range, (list, tuple)) else lookup_range
    if match_type == 0:
        for i, v in enumerate(flat):
            if v == lookup_val:
                return i + 1
    elif match_type == 1:
        last_pos = None
        for i, v in enumerate(flat):
            if v is not None and v <= lookup_val:
                last_pos = i + 1
        return last_pos
    elif match_type == -1:
        last_pos = None
        for i, v in enumerate(flat):
            if v is not None and v >= lookup_val:
                last_pos = i + 1
        return last_pos
    return None


def xl_iferror(value, error_val):
    try:
        if value is None:
            return error_val
        return value
    except Exception:
        return error_val


def xl_isblank(val):
    return val is None or val == ""


def xl_sumif(criteria_range, criteria, sum_range=None):
    if sum_range is None:
        sum_range = criteria_range
    cr = _flatten(criteria_range)
    sr = _flatten(sum_range)
    total = 0
    for i, v in enumerate(cr):
        if _match_criteria(v, criteria) and i < len(sr):
            total += _to_num(sr[i])
    return total


def xl_sumifs(sum_range, *args):
    sr = _flatten(sum_range)
    pairs = list(zip(args[::2], args[1::2]))
    total = 0
    for i in range(len(sr)):
        match = True
        for criteria_range, criteria in pairs:
            cr = _flatten(criteria_range)
            if i >= len(cr) or not _match_criteria(cr[i], criteria):
                match = False
                break
        if match:
            total += _to_num(sr[i])
    return total


def xl_countif(criteria_range, criteria):
    cr = _flatten(criteria_range)
    return sum(1 for v in cr if _match_criteria(v, criteria))


def xl_countifs(*args):
    pairs = list(zip(args[::2], args[1::2]))
    if not pairs:
        return 0
    first_range = _flatten(pairs[0][0])
    count = 0
    for i in range(len(first_range)):
        match = True
        for criteria_range, criteria in pairs:
            cr = _flatten(criteria_range)
            if i >= len(cr) or not _match_criteria(cr[i], criteria):
                match = False
                break
        if match:
            count += 1
    return count


def xl_averageif(criteria_range, criteria, avg_range=None):
    if avg_range is None:
        avg_range = criteria_range
    cr = _flatten(criteria_range)
    ar = _flatten(avg_range)
    vals = []
    for i, v in enumerate(cr):
        if _match_criteria(v, criteria) and i < len(ar):
            vals.append(_to_num(ar[i]))
    return sum(vals) / len(vals) if vals else 0


def xl_sumproduct(*arrays):
    flat_arrays = [_flatten(a) for a in arrays]
    if not flat_arrays:
        return 0
    length = min(len(a) for a in flat_arrays)
    total = 0
    for i in range(length):
        product = 1
        for arr in flat_arrays:
            product *= _to_num(arr[i])
        total += product
    return total


def xl_offset(base_range, rows, cols, height=None, width=None):
    # Simplified: return base_range offset (not fully dynamic in static code)
    return base_range


def xl_indirect(ref_str):
    # Cannot be fully supported in static code generation
    return ref_str


def xl_row(ref=None):
    return 1  # Placeholder


def xl_column(ref=None):
    return 1  # Placeholder


def xl_rows(ref):
    if isinstance(ref, list):
        return len(ref)
    return 1


def xl_columns(ref):
    if isinstance(ref, list) and ref:
        if isinstance(ref[0], (list, tuple)):
            return len(ref[0])
    return 1


def xl_today():
    return datetime.date.today()


def xl_now():
    return datetime.datetime.now()


def xl_year(date_val):
    if isinstance(date_val, (datetime.date, datetime.datetime)):
        return date_val.year
    return 0


def xl_month(date_val):
    if isinstance(date_val, (datetime.date, datetime.datetime)):
        return date_val.month
    return 0


def xl_day(date_val):
    if isinstance(date_val, (datetime.date, datetime.datetime)):
        return date_val.day
    return 0


def xl_date(year, month, day):
    return datetime.date(int(year), int(month), int(day))


def xl_eomonth(start_date, months):
    import calendar
    if isinstance(start_date, (datetime.date, datetime.datetime)):
        m = start_date.month + int(months)
        y = start_date.year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        d = calendar.monthrange(y, m)[1]
        return datetime.date(y, m, d)
    return start_date


def xl_edate(start_date, months):
    if isinstance(start_date, (datetime.date, datetime.datetime)):
        m = start_date.month + int(months)
        y = start_date.year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        d = min(start_date.day, __import__("calendar").monthrange(y, m)[1])
        return datetime.date(y, m, d)
    return start_date


def xl_datedif(start, end, unit):
    if isinstance(start, (datetime.date, datetime.datetime)) and isinstance(end, (datetime.date, datetime.datetime)):
        if unit.upper() == "D":
            return (end - start).days
        if unit.upper() == "M":
            return (end.year - start.year) * 12 + end.month - start.month
        if unit.upper() == "Y":
            return end.year - start.year
    return 0


def xl_pi():
    return math.pi


def _match_criteria(value, criteria):
    """Match a value against an Excel-style criteria string."""
    if isinstance(criteria, str):
        if criteria.startswith(">="):
            return _to_num(value) >= _to_num(criteria[2:])
        if criteria.startswith("<="):
            return _to_num(value) <= _to_num(criteria[2:])
        if criteria.startswith("<>"):
            return value != criteria[2:] and str(value) != criteria[2:]
        if criteria.startswith(">"):
            return _to_num(value) > _to_num(criteria[1:])
        if criteria.startswith("<"):
            return _to_num(value) < _to_num(criteria[1:])
        if criteria.startswith("="):
            crit_val = criteria[1:]
            return value == crit_val or str(value) == crit_val
        # Wildcard matching
        if "*" in criteria or "?" in criteria:
            import fnmatch
            return fnmatch.fnmatch(str(value).lower(), criteria.lower())
        return value == criteria or str(value) == criteria
    return value == criteria


def _get_range(ws_data, col1, row1, col2, row2):
    """Extract a range of values from worksheet data dict.
    ws_data: dict mapping (col, row) -> value
    """
    result = []
    col1_idx = _col_to_idx(col1)
    col2_idx = _col_to_idx(col2)
    for r in range(row1, row2 + 1):
        row_vals = []
        for c in range(col1_idx, col2_idx + 1):
            col_letter = _idx_to_col(c)
            row_vals.append(ws_data.get((col_letter, r), None))
        result.append(row_vals)
    return result


def _col_to_idx(col_str):
    result = 0
    for char in col_str.upper():
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def _idx_to_col(index):
    result = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _get_table_column(table_data, col_name):
    """Get values from a table column. table_data is list of dicts."""
    return [row.get(col_name, None) for row in table_data]
'''
