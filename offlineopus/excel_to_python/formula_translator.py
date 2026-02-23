"""
Formula Translator Module
=========================
Translates Excel formulas into Python expressions.
Handles cell references, ranges, cross-sheet references, and common Excel functions.
"""

import re
import logging
from openpyxl.utils import get_column_letter, column_index_from_string

logger = logging.getLogger(__name__)

# Excel functions mapped to Python equivalents
EXCEL_FUNC_MAP = {
    'SUM': 'xl_sum',
    'AVERAGE': 'xl_average',
    'COUNT': 'xl_count',
    'COUNTA': 'xl_counta',
    'COUNTIF': 'xl_countif',
    'COUNTIFS': 'xl_countifs',
    'MIN': 'xl_min',
    'MAX': 'xl_max',
    'ABS': 'abs',
    'ROUND': 'xl_round',
    'ROUNDUP': 'xl_roundup',
    'ROUNDDOWN': 'xl_rounddown',
    'INT': 'int',
    'MOD': 'xl_mod',
    'POWER': 'xl_power',
    'SQRT': 'xl_sqrt',
    'LN': 'xl_ln',
    'LOG': 'xl_log',
    'LOG10': 'xl_log10',
    'EXP': 'xl_exp',
    'IF': 'xl_if',
    'AND': 'xl_and',
    'OR': 'xl_or',
    'NOT': 'xl_not',
    'IFERROR': 'xl_iferror',
    'IFNA': 'xl_ifna',
    'ISERROR': 'xl_iserror',
    'ISNA': 'xl_isna',
    'ISBLANK': 'xl_isblank',
    'ISNUMBER': 'xl_isnumber',
    'VLOOKUP': 'xl_vlookup',
    'HLOOKUP': 'xl_hlookup',
    'INDEX': 'xl_index',
    'MATCH': 'xl_match',
    'OFFSET': 'xl_offset',
    'INDIRECT': 'xl_indirect',
    'ROW': 'xl_row',
    'COLUMN': 'xl_column',
    'ROWS': 'xl_rows',
    'COLUMNS': 'xl_columns',
    'LEFT': 'xl_left',
    'RIGHT': 'xl_right',
    'MID': 'xl_mid',
    'LEN': 'xl_len',
    'TRIM': 'xl_trim',
    'UPPER': 'xl_upper',
    'LOWER': 'xl_lower',
    'CONCATENATE': 'xl_concatenate',
    'TEXT': 'xl_text',
    'VALUE': 'xl_value',
    'FIND': 'xl_find',
    'SEARCH': 'xl_search',
    'SUBSTITUTE': 'xl_substitute',
    'REPLACE': 'xl_replace',
    'EOMONTH': 'xl_eomonth',
    'EDATE': 'xl_edate',
    'DATE': 'xl_date',
    'YEAR': 'xl_year',
    'MONTH': 'xl_month',
    'DAY': 'xl_day',
    'TODAY': 'xl_today',
    'NOW': 'xl_now',
    'DAYS': 'xl_days',
    'SUMIF': 'xl_sumif',
    'SUMIFS': 'xl_sumifs',
    'SUMPRODUCT': 'xl_sumproduct',
    'AVERAGEIF': 'xl_averageif',
    'AVERAGEIFS': 'xl_averageifs',
    'MAXIFS': 'xl_maxifs',
    'MINIFS': 'xl_minifs',
    'LARGE': 'xl_large',
    'SMALL': 'xl_small',
    'NPV': 'xl_npv',
    'IRR': 'xl_irr',
    'XNPV': 'xl_xnpv',
    'XIRR': 'xl_xirr',
    'PMT': 'xl_pmt',
    'PPMT': 'xl_ppmt',
    'IPMT': 'xl_ipmt',
    'PV': 'xl_pv',
    'FV': 'xl_fv',
    'NPER': 'xl_nper',
    'RATE': 'xl_rate',
    'CEILING': 'xl_ceiling',
    'FLOOR': 'xl_floor',
    'MEDIAN': 'xl_median',
    'STDEV': 'xl_stdev',
    'VAR': 'xl_var',
    'CHOOSE': 'xl_choose',
    'TRANSPOSE': 'xl_transpose',
    'LOOKUP': 'xl_lookup',
    'NA': 'xl_na',
    'PI': 'xl_pi',
    'TRUE': 'xl_true',
    'FALSE': 'xl_false',
}

# Cell reference regex patterns
# Matches: A1, $A$1, A$1, $A1, AA100, etc.
CELL_REF_PATTERN = r"\$?([A-Za-z]{1,3})\$?(\d{1,7})"
# Matches: Sheet!A1 or 'Sheet Name'!A1
SHEET_REF_PATTERN = r"(?:'([^']+)'|([A-Za-z_]\w*))!" + CELL_REF_PATTERN
# Range: A1:B10
RANGE_PATTERN = CELL_REF_PATTERN + r":" + CELL_REF_PATTERN
# Sheet range: Sheet!A1:B10
SHEET_RANGE_PATTERN = r"(?:'([^']+)'|([A-Za-z_]\w*))!" + CELL_REF_PATTERN + r":" + CELL_REF_PATTERN


def cell_to_key(col_str: str, row_str: str, sheet: str = None) -> str:
    """Convert a cell reference to a cells dict key string."""
    col_num = column_index_from_string(col_str.upper().replace('$', ''))
    row_num = int(row_str.replace('$', ''))
    return f"cells[({repr(sheet)}, {row_num}, {col_num})]"


def range_to_python(col1: str, row1: str, col2: str, row2: str, sheet: str = None) -> str:
    """Convert a range reference to a Python expression that yields a list of cell values."""
    c1 = column_index_from_string(col1.upper().replace('$', ''))
    r1 = int(row1.replace('$', ''))
    c2 = column_index_from_string(col2.upper().replace('$', ''))
    r2 = int(row2.replace('$', ''))
    return f"cell_range(cells, {repr(sheet)}, {r1}, {c1}, {r2}, {c2})"


def extract_cell_references(formula: str, current_sheet: str):
    """
    Extract all cell references from a formula.
    Returns a list of (sheet, row, col) tuples.
    """
    refs = []
    
    # Match sheet-qualified range references: 'Sheet'!A1:B10
    for m in re.finditer(
        r"(?:'([^']+)'|([A-Za-z_][\w ]*))\!" 
        r"\$?([A-Za-z]{1,3})\$?(\d{1,7})"
        r"(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?",
        formula
    ):
        sheet = m.group(1) or m.group(2)
        col1 = m.group(3).upper()
        row1 = int(m.group(4))
        c1 = column_index_from_string(col1)
        
        if m.group(5) and m.group(6):  # Range
            col2 = m.group(5).upper()
            row2 = int(m.group(6))
            c2 = column_index_from_string(col2)
            for r in range(row1, row2 + 1):
                for c in range(c1, c2 + 1):
                    refs.append((sheet, r, c))
        else:  # Single cell
            refs.append((sheet, row1, c1))

    # Now match non-sheet-qualified references
    # Remove parts already matched (sheet-qualified)
    remaining = re.sub(
        r"(?:'[^']+'|[A-Za-z_][\w ]*)\!\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?",
        "###",
        formula
    )
    
    # Match range references: A1:B10
    for m in re.finditer(
        r"\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})",
        remaining
    ):
        col1 = m.group(1).upper()
        row1 = int(m.group(2))
        col2 = m.group(3).upper()
        row2 = int(m.group(4))
        c1 = column_index_from_string(col1)
        c2 = column_index_from_string(col2)
        for r in range(row1, row2 + 1):
            for c in range(c1, c2 + 1):
                refs.append((current_sheet, r, c))

    # Remove ranges from remaining
    remaining2 = re.sub(
        r"\$?[A-Za-z]{1,3}\$?\d{1,7}:\$?[A-Za-z]{1,3}\$?\d{1,7}",
        "###",
        remaining
    )

    # Match single cell references: A1
    for m in re.finditer(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})", remaining2):
        col_str = m.group(1).upper()
        row_num = int(m.group(2))
        # Filter out things that look like cell refs but aren't (e.g., function names)
        # Functions are followed by ( so we can exclude those
        start = m.start()
        # Check if this is preceded by alpha chars (part of a function name)
        if start > 0 and remaining2[start-1:start].isalpha():
            continue
        try:
            col_num = column_index_from_string(col_str)
            refs.append((current_sheet, row_num, col_num))
        except ValueError:
            pass

    return refs


class FormulaTranslator:
    """Translates an Excel formula to a Python expression."""

    def __init__(self, current_sheet: str):
        self.current_sheet = current_sheet
        self.referenced_cells = []  # (sheet, row, col) list

    def translate(self, formula: str) -> str:
        """
        Translate an Excel formula to a Python expression.

        Args:
            formula: Excel formula string (without leading '=')

        Returns:
            Python expression string
        """
        self.referenced_cells = extract_cell_references(formula, self.current_sheet)
        result = self._translate_expr(formula)
        return result

    def _translate_expr(self, formula: str) -> str:
        """Main translation logic."""
        # Handle percentage literals like 1.426%
        formula = re.sub(r'(\d+\.?\d*)\%', lambda m: str(float(m.group(1)) / 100), formula)

        # Tokenize and translate
        tokens = self._tokenize(formula)
        python_tokens = self._translate_tokens(tokens)
        return ''.join(python_tokens)

    def _tokenize(self, formula: str) -> list:
        """
        Tokenize an Excel formula into a list of tokens.
        Each token is a tuple: (type, value)
        Types: 'SHEET_RANGE', 'SHEET_REF', 'RANGE', 'CELL', 'FUNC', 'NUMBER',
               'STRING', 'BOOL', 'OP', 'COMMA', 'PAREN', 'SEMICOLON', 'UNKNOWN'
        """
        tokens = []
        i = 0
        s = formula

        while i < len(s):
            # Skip whitespace
            if s[i] == ' ':
                i += 1
                continue

            # String literal
            if s[i] == '"':
                j = i + 1
                while j < len(s) and s[j] != '"':
                    if s[j] == '\\':
                        j += 1
                    j += 1
                tokens.append(('STRING', s[i:j+1]))
                i = j + 1
                continue

            # Sheet-qualified reference (with quotes): 'Sheet Name'!A1:B10 or 'Sheet Name'!A1
            m = re.match(
                r"'([^']+)'\!\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::(\$?[A-Za-z]{1,3})\$?(\d{1,7}))?",
                s[i:]
            )
            if m:
                if m.group(4) and m.group(5):
                    tokens.append(('SHEET_RANGE', m.group(0)))
                else:
                    tokens.append(('SHEET_REF', m.group(0)))
                i += m.end()
                continue

            # Sheet-qualified reference (without quotes): Sheet!A1:B10 or Sheet!A1
            m = re.match(
                r"([A-Za-z_]\w*)\!\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::(\$?[A-Za-z]{1,3})\$?(\d{1,7}))?",
                s[i:]
            )
            if m:
                # Make sure it's not a function name
                name = m.group(1).upper()
                if name not in EXCEL_FUNC_MAP:
                    if m.group(4) and m.group(5):
                        tokens.append(('SHEET_RANGE', m.group(0)))
                    else:
                        tokens.append(('SHEET_REF', m.group(0)))
                    i += m.end()
                    continue

            # Range reference: A1:B10
            m = re.match(
                r"\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})",
                s[i:]
            )
            if m:
                # Check it's not preceded by alpha (part of function name)
                if i == 0 or not s[i-1].isalpha():
                    tokens.append(('RANGE', m.group(0)))
                    i += m.end()
                    continue

            # Cell reference: A1, $A$1, etc.
            m = re.match(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})", s[i:])
            if m:
                # Check it's not preceded by alpha (part of function name)
                if i == 0 or not s[i-1].isalpha():
                    # Also check it's not followed by ( which would make it a function
                    end_pos = i + m.end()
                    if end_pos < len(s) and s[end_pos] == '(':
                        # This is a function call - treat as FUNC
                        pass
                    else:
                        tokens.append(('CELL', m.group(0)))
                        i += m.end()
                        continue

            # Function name
            m = re.match(r"([A-Za-z_][\w.]*)\s*\(", s[i:])
            if m:
                tokens.append(('FUNC', m.group(1)))
                i += len(m.group(1))
                # Don't consume the paren - it'll be caught below
                continue

            # Number
            m = re.match(r"(\d+\.?\d*(?:[eE][+-]?\d+)?)", s[i:])
            if m:
                tokens.append(('NUMBER', m.group(0)))
                i += m.end()
                continue

            # Boolean
            m = re.match(r"(TRUE|FALSE)\b", s[i:], re.IGNORECASE)
            if m:
                tokens.append(('BOOL', m.group(0)))
                i += m.end()
                continue

            # Operators
            if s[i:i+2] in ('<>', '<=', '>='):
                tokens.append(('OP', s[i:i+2]))
                i += 2
                continue
            if s[i] in '+-*/^&=<>':
                tokens.append(('OP', s[i]))
                i += 1
                continue

            # Comma
            if s[i] == ',':
                tokens.append(('COMMA', ','))
                i += 1
                continue

            # Semicolon (used like comma in some locales)
            if s[i] == ';':
                tokens.append(('COMMA', ','))
                i += 1
                continue

            # Parentheses
            if s[i] in '()':
                tokens.append(('PAREN', s[i]))
                i += 1
                continue

            # Colon (standalone - e.g., in structured references)
            if s[i] == ':':
                tokens.append(('OP', ':'))
                i += 1
                continue

            # Anything else
            tokens.append(('UNKNOWN', s[i]))
            i += 1

        return tokens

    def _translate_tokens(self, tokens: list) -> list:
        """Translate a list of tokens to Python expression parts."""
        result = []
        i = 0

        # Track if previous token suggests unary minus
        prev_type = None

        while i < len(tokens):
            ttype, tval = tokens[i]

            if ttype == 'SHEET_RANGE':
                result.append(self._translate_sheet_range(tval))
            elif ttype == 'SHEET_REF':
                result.append(self._translate_sheet_ref(tval))
            elif ttype == 'RANGE':
                result.append(self._translate_range(tval))
            elif ttype == 'CELL':
                result.append(self._translate_cell(tval))
            elif ttype == 'FUNC':
                result.append(self._translate_func(tval))
            elif ttype == 'NUMBER':
                result.append(tval)
            elif ttype == 'STRING':
                result.append(tval)
            elif ttype == 'BOOL':
                result.append('True' if tval.upper() == 'TRUE' else 'False')
            elif ttype == 'OP':
                result.append(self._translate_op(tval))
            elif ttype == 'COMMA':
                result.append(', ')
            elif ttype == 'PAREN':
                result.append(tval)
            elif ttype == 'UNKNOWN':
                result.append(tval)

            prev_type = ttype
            i += 1

        return result

    def _translate_op(self, op: str) -> str:
        """Translate an Excel operator to Python."""
        op_map = {
            '=': ' == ',
            '<>': ' != ',
            '<=': ' <= ',
            '>=': ' >= ',
            '<': ' < ',
            '>': ' > ',
            '+': ' + ',
            '-': ' - ',
            '*': ' * ',
            '/': ' / ',
            '^': ' ** ',
            '&': ' + str_concat ',  # String concatenation - handled specially
        }
        if op == '&':
            return ' + '  # Simplify: assume string concat works with +
        return op_map.get(op, f' {op} ')

    def _translate_cell(self, ref: str) -> str:
        """Translate a cell reference like A1, $A$1 to cells[...] lookup."""
        m = re.match(r"\$?([A-Za-z]{1,3})\$?(\d{1,7})", ref)
        if m:
            return cell_to_key(m.group(1), m.group(2), self.current_sheet)
        return ref

    def _translate_sheet_ref(self, ref: str) -> str:
        """Translate a sheet-qualified cell reference like 'Sheet'!A1."""
        m = re.match(
            r"(?:'([^']+)'|([A-Za-z_]\w*))\!\$?([A-Za-z]{1,3})\$?(\d{1,7})",
            ref
        )
        if m:
            sheet = m.group(1) or m.group(2)
            return cell_to_key(m.group(3), m.group(4), sheet)
        return ref

    def _translate_range(self, ref: str) -> str:
        """Translate a range reference like A1:B10."""
        m = re.match(
            r"\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})",
            ref
        )
        if m:
            return range_to_python(m.group(1), m.group(2), m.group(3), m.group(4), self.current_sheet)
        return ref

    def _translate_sheet_range(self, ref: str) -> str:
        """Translate a sheet-qualified range like 'Sheet'!A1:B10."""
        m = re.match(
            r"(?:'([^']+)'|([A-Za-z_]\w*))\!\$?([A-Za-z]{1,3})\$?(\d{1,7}):\$?([A-Za-z]{1,3})\$?(\d{1,7})",
            ref
        )
        if m:
            sheet = m.group(1) or m.group(2)
            return range_to_python(m.group(3), m.group(4), m.group(5), m.group(6), sheet)
        return ref

    def _translate_func(self, func_name: str) -> str:
        """Translate an Excel function name to Python."""
        upper_name = func_name.upper()
        if upper_name in EXCEL_FUNC_MAP:
            return EXCEL_FUNC_MAP[upper_name]
        # Unknown function - keep as-is with xl_ prefix
        logger.warning(f"Unknown Excel function: {func_name}")
        return f"xl_{func_name.lower()}"


def translate_formula(formula: str, current_sheet: str) -> tuple:
    """
    Convenience function to translate a formula.

    Returns:
        (python_expression, list_of_referenced_cells)
    """
    translator = FormulaTranslator(current_sheet)
    python_expr = translator.translate(formula)
    return python_expr, translator.referenced_cells
