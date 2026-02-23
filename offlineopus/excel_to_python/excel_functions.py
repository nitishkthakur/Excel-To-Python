"""
Excel Functions Module
======================
Python implementations of common Excel functions.
This file is copied to the output directory and used by the generated calculator.
"""

import math
import datetime
from calendar import monthrange
from typing import Any, List, Optional, Union

# Sentinel for empty/missing cells
EMPTY = None


def _num(val):
    """Convert a value to a number, treating None/empty as 0."""
    if val is None:
        return 0
    if isinstance(val, bool):
        return 1 if val else 0
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, str):
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0
    if isinstance(val, datetime.datetime):
        # Excel serial date
        delta = val - datetime.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    if isinstance(val, datetime.date):
        delta = val - datetime.date(1899, 12, 30)
        return delta.days
    return 0


def _nums(vals):
    """Flatten and convert an iterable of values to numbers, ignoring non-numeric."""
    result = []
    if isinstance(vals, (list, tuple)):
        for v in vals:
            if isinstance(v, (list, tuple)):
                result.extend(_nums(v))
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                result.append(v)
            elif v is not None:
                try:
                    result.append(float(v))
                except (ValueError, TypeError):
                    pass
    elif isinstance(vals, (int, float)) and not isinstance(vals, bool):
        result.append(vals)
    return result


def _flatten(*args):
    """Flatten nested lists/tuples into a single list."""
    result = []
    for a in args:
        if isinstance(a, (list, tuple)):
            result.extend(_flatten(*a))
        else:
            result.append(a)
    return result


def _flatten_nums(*args):
    """Flatten and keep only numeric values."""
    flat = _flatten(*args)
    return [_num(v) for v in flat if v is not None and not isinstance(v, str)]


def cell_range(cells, sheet, r1, c1, r2, c2):
    """Get a list of cell values from a range."""
    values = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            values.append(cells.get((sheet, r, c), None))
    return values


# ============================================================
# Math Functions
# ============================================================

def xl_sum(*args):
    """SUM function - sum all numeric values."""
    return sum(_flatten_nums(*args))


def xl_average(*args):
    """AVERAGE function."""
    nums = _flatten_nums(*args)
    return sum(nums) / len(nums) if nums else 0


def xl_count(*args):
    """COUNT function - count numeric values."""
    return len(_flatten_nums(*args))


def xl_counta(*args):
    """COUNTA function - count non-empty values."""
    flat = _flatten(*args)
    return sum(1 for v in flat if v is not None and v != '')


def xl_min(*args):
    """MIN function."""
    nums = _flatten_nums(*args)
    return min(nums) if nums else 0


def xl_max(*args):
    """MAX function."""
    nums = _flatten_nums(*args)
    return max(nums) if nums else 0


def xl_abs(val):
    """ABS function."""
    return abs(_num(val))


def xl_round(val, digits=0):
    """ROUND function."""
    return round(_num(val), int(_num(digits)))


def xl_roundup(val, digits=0):
    """ROUNDUP function."""
    d = int(_num(digits))
    n = _num(val)
    factor = 10 ** d
    if n >= 0:
        return math.ceil(n * factor) / factor
    else:
        return math.floor(n * factor) / factor


def xl_rounddown(val, digits=0):
    """ROUNDDOWN function."""
    d = int(_num(digits))
    n = _num(val)
    factor = 10 ** d
    if n >= 0:
        return math.floor(n * factor) / factor
    else:
        return math.ceil(n * factor) / factor


def xl_mod(n, d):
    """MOD function."""
    return _num(n) % _num(d) if _num(d) != 0 else 0


def xl_power(base, exp):
    """POWER function."""
    return _num(base) ** _num(exp)


def xl_sqrt(val):
    """SQRT function."""
    return math.sqrt(_num(val))


def xl_ln(val):
    """LN function - natural log."""
    n = _num(val)
    return math.log(n) if n > 0 else 0


def xl_log(val, base=10):
    """LOG function."""
    n = _num(val)
    b = _num(base)
    return math.log(n, b) if n > 0 and b > 0 else 0


def xl_log10(val):
    """LOG10 function."""
    n = _num(val)
    return math.log10(n) if n > 0 else 0


def xl_exp(val):
    """EXP function."""
    return math.exp(_num(val))


def xl_ceiling(val, significance=1):
    """CEILING function."""
    n = _num(val)
    s = _num(significance)
    if s == 0:
        return 0
    return math.ceil(n / s) * s


def xl_floor(val, significance=1):
    """FLOOR function."""
    n = _num(val)
    s = _num(significance)
    if s == 0:
        return 0
    return math.floor(n / s) * s


def xl_median(*args):
    """MEDIAN function."""
    nums = sorted(_flatten_nums(*args))
    n = len(nums)
    if n == 0:
        return 0
    mid = n // 2
    if n % 2 == 0:
        return (nums[mid - 1] + nums[mid]) / 2
    return nums[mid]


def xl_stdev(*args):
    """STDEV function (sample standard deviation)."""
    nums = _flatten_nums(*args)
    n = len(nums)
    if n < 2:
        return 0
    mean = sum(nums) / n
    return math.sqrt(sum((x - mean) ** 2 for x in nums) / (n - 1))


def xl_var(*args):
    """VAR function (sample variance)."""
    nums = _flatten_nums(*args)
    n = len(nums)
    if n < 2:
        return 0
    mean = sum(nums) / n
    return sum((x - mean) ** 2 for x in nums) / (n - 1)


def xl_pi():
    """PI function."""
    return math.pi


# ============================================================
# Logical Functions
# ============================================================

def xl_if(condition, true_val, false_val=False):
    """IF function."""
    return true_val if condition else false_val


def xl_and(*args):
    """AND function."""
    flat = _flatten(*args)
    return all(bool(v) for v in flat if v is not None)


def xl_or(*args):
    """OR function."""
    flat = _flatten(*args)
    return any(bool(v) for v in flat if v is not None)


def xl_not(val):
    """NOT function."""
    return not bool(val)


def xl_iferror(val, error_val):
    """IFERROR function - return error_val if val is an error."""
    try:
        if val is None or (isinstance(val, float) and math.isnan(val)):
            return error_val
        return val
    except Exception:
        return error_val


def xl_ifna(val, na_val):
    """IFNA function."""
    if val is None:
        return na_val
    return val


def xl_iserror(val):
    """ISERROR function."""
    try:
        if val is None or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
            return True
        return False
    except Exception:
        return True


def xl_isna(val):
    """ISNA function."""
    return val is None


def xl_isblank(val):
    """ISBLANK function."""
    return val is None or val == ''


def xl_isnumber(val):
    """ISNUMBER function."""
    return isinstance(val, (int, float)) and not isinstance(val, bool)


def xl_true():
    return True


def xl_false():
    return False


def xl_na():
    return None


# ============================================================
# Lookup Functions
# ============================================================

def xl_vlookup(lookup_val, table_array, col_index, range_lookup=True):
    """VLOOKUP function."""
    lookup_val = _num(lookup_val) if isinstance(lookup_val, (int, float)) else lookup_val

    if not isinstance(table_array, list) or len(table_array) == 0:
        return None

    # table_array should be a flat list; figure out dimensions
    # We assume it's from cell_range which returns a flat list
    # Need to know the number of columns
    col_idx = int(_num(col_index))

    # Try to detect dimensions - this is tricky with flat lists
    # If range_lookup is False (exact match)
    if not range_lookup:
        for i, val in enumerate(table_array):
            if val == lookup_val or (isinstance(val, (int, float)) and isinstance(lookup_val, (int, float)) and abs(_num(val) - _num(lookup_val)) < 1e-10):
                # Return the value at col_index offset
                # This approximation works for single-column lookups
                return table_array[i + col_idx - 1] if i + col_idx - 1 < len(table_array) else None
    else:
        # Approximate match (sorted ascending)
        best_idx = None
        for i, val in enumerate(table_array):
            if val is not None and _num(val) <= _num(lookup_val):
                best_idx = i
        if best_idx is not None:
            return table_array[best_idx + col_idx - 1] if best_idx + col_idx - 1 < len(table_array) else None

    return None


def xl_hlookup(lookup_val, table_array, row_index, range_lookup=True):
    """HLOOKUP function."""
    return xl_vlookup(lookup_val, table_array, row_index, range_lookup)


def xl_index(array, row_num, col_num=None):
    """INDEX function."""
    if isinstance(array, (list, tuple)):
        idx = int(_num(row_num)) - 1
        if 0 <= idx < len(array):
            result = array[idx]
            if col_num is not None and isinstance(result, (list, tuple)):
                cidx = int(_num(col_num)) - 1
                if 0 <= cidx < len(result):
                    return result[cidx]
            return result
    return None


def xl_match(lookup_val, lookup_array, match_type=1):
    """MATCH function."""
    if not isinstance(lookup_array, (list, tuple)):
        return None

    mt = int(_num(match_type))
    lookup_val = _num(lookup_val) if isinstance(lookup_val, (int, float)) else lookup_val

    if mt == 0:  # Exact match
        for i, val in enumerate(lookup_array):
            if val == lookup_val or (_num(val) == _num(lookup_val) if isinstance(val, (int, float)) and isinstance(lookup_val, (int, float)) else False):
                return i + 1
        return None
    elif mt == 1:  # Largest value <= lookup_val
        best_idx = None
        for i, val in enumerate(lookup_array):
            if val is not None and _num(val) <= _num(lookup_val):
                best_idx = i
        return best_idx + 1 if best_idx is not None else None
    elif mt == -1:  # Smallest value >= lookup_val
        best_idx = None
        for i, val in enumerate(lookup_array):
            if val is not None and _num(val) >= _num(lookup_val):
                best_idx = i
        return best_idx + 1 if best_idx is not None else None

    return None


def xl_offset(cells, sheet, base_row, base_col, row_offset, col_offset, height=1, width=1):
    """OFFSET function - returns a cell value or range."""
    r = int(_num(base_row)) + int(_num(row_offset))
    c = int(_num(base_col)) + int(_num(col_offset))
    h = int(_num(height))
    w = int(_num(width))

    if h == 1 and w == 1:
        return cells.get((sheet, r, c), None)
    else:
        return cell_range(cells, sheet, r, c, r + h - 1, c + w - 1)


def xl_indirect(ref_text, cells=None, current_sheet=None):
    """
    INDIRECT function - evaluates a cell reference from a string.
    This is a simplified implementation.
    """
    if cells is None or ref_text is None:
        return None

    import re
    from openpyxl.utils import column_index_from_string

    ref_text = str(ref_text).strip().strip("'").strip('"')

    # Handle sheet!cell format
    if '!' in ref_text:
        parts = ref_text.split('!', 1)
        sheet = parts[0].strip("'")
        cell_ref = parts[1]
    else:
        sheet = current_sheet
        cell_ref = ref_text

    m = re.match(r'\$?([A-Za-z]{1,3})\$?(\d{1,7})', cell_ref)
    if m:
        col = column_index_from_string(m.group(1).upper())
        row = int(m.group(2))
        return cells.get((sheet, row, col), None)

    return None


def xl_row(ref=None):
    """ROW function - simplified."""
    if ref is None:
        return 1
    return ref


def xl_column(ref=None):
    """COLUMN function - simplified."""
    if ref is None:
        return 1
    return ref


def xl_rows(array):
    """ROWS function."""
    if isinstance(array, (list, tuple)):
        return len(array)
    return 1


def xl_columns(array):
    """COLUMNS function."""
    if isinstance(array, (list, tuple)) and len(array) > 0:
        if isinstance(array[0], (list, tuple)):
            return len(array[0])
    return 1


def xl_choose(index, *choices):
    """CHOOSE function."""
    idx = int(_num(index))
    if 1 <= idx <= len(choices):
        return choices[idx - 1]
    return None


def xl_lookup(lookup_val, lookup_vector, result_vector=None):
    """LOOKUP function."""
    if result_vector is None:
        result_vector = lookup_vector
    if not isinstance(lookup_vector, (list, tuple)):
        return None
    best_idx = None
    for i, val in enumerate(lookup_vector):
        if val is not None and _num(val) <= _num(lookup_val):
            best_idx = i
    if best_idx is not None and best_idx < len(result_vector):
        return result_vector[best_idx]
    return None


# ============================================================
# Text Functions
# ============================================================

def xl_left(text, num_chars=1):
    """LEFT function."""
    return str(text)[:int(_num(num_chars))]


def xl_right(text, num_chars=1):
    """RIGHT function."""
    return str(text)[-int(_num(num_chars)):]


def xl_mid(text, start_num, num_chars):
    """MID function."""
    s = int(_num(start_num)) - 1
    n = int(_num(num_chars))
    return str(text)[s:s + n]


def xl_len(text):
    """LEN function."""
    return len(str(text)) if text is not None else 0


def xl_trim(text):
    """TRIM function."""
    return str(text).strip() if text is not None else ''


def xl_upper(text):
    """UPPER function."""
    return str(text).upper() if text is not None else ''


def xl_lower(text):
    """LOWER function."""
    return str(text).lower() if text is not None else ''


def xl_concatenate(*args):
    """CONCATENATE function."""
    return ''.join(str(a) if a is not None else '' for a in args)


def xl_text(value, format_text):
    """TEXT function - simplified formatting."""
    try:
        v = _num(value)
        fmt = str(format_text)
        if '#' in fmt or '0' in fmt:
            # Numeric formatting - approximate
            decimals = 0
            if '.' in fmt:
                decimals = len(fmt.split('.')[-1].replace('#', '0').rstrip('0')) or len(fmt.split('.')[-1])
            return f"{v:.{decimals}f}"
        if '%' in fmt:
            return f"{v * 100:.1f}%"
        return str(value)
    except Exception:
        return str(value)


def xl_value(text):
    """VALUE function."""
    try:
        return float(str(text).replace(',', ''))
    except (ValueError, TypeError):
        return 0


def xl_find(find_text, within_text, start_num=1):
    """FIND function (case-sensitive)."""
    try:
        s = int(_num(start_num)) - 1
        idx = str(within_text).index(str(find_text), s)
        return idx + 1
    except (ValueError, AttributeError):
        return None


def xl_search(find_text, within_text, start_num=1):
    """SEARCH function (case-insensitive)."""
    try:
        s = int(_num(start_num)) - 1
        idx = str(within_text).lower().index(str(find_text).lower(), s)
        return idx + 1
    except (ValueError, AttributeError):
        return None


def xl_substitute(text, old_text, new_text, instance_num=None):
    """SUBSTITUTE function."""
    t = str(text) if text is not None else ''
    old = str(old_text)
    new = str(new_text)
    if instance_num is not None:
        n = int(_num(instance_num))
        count = 0
        result = ''
        i = 0
        while i < len(t):
            if t[i:i + len(old)] == old:
                count += 1
                if count == n:
                    result += new
                else:
                    result += old
                i += len(old)
            else:
                result += t[i]
                i += 1
        return result
    return t.replace(old, new)


def xl_replace(old_text, start_num, num_chars, new_text):
    """REPLACE function."""
    t = str(old_text)
    s = int(_num(start_num)) - 1
    n = int(_num(num_chars))
    return t[:s] + str(new_text) + t[s + n:]


# ============================================================
# Date Functions
# ============================================================

def _excel_date_to_python(serial):
    """Convert Excel serial date to Python datetime."""
    if isinstance(serial, (datetime.datetime, datetime.date)):
        return serial
    n = _num(serial)
    if n == 0:
        return datetime.datetime(1899, 12, 30)
    return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=n)


def _python_to_excel_date(dt):
    """Convert Python datetime to Excel serial date."""
    if isinstance(dt, (int, float)):
        return dt
    if isinstance(dt, datetime.datetime):
        delta = dt - datetime.datetime(1899, 12, 30)
    elif isinstance(dt, datetime.date):
        delta = dt - datetime.date(1899, 12, 30)
    else:
        return 0
    return delta.days + (delta.seconds / 86400 if hasattr(delta, 'seconds') else 0)


def xl_eomonth(start_date, months):
    """EOMONTH function - end of month date."""
    dt = _excel_date_to_python(start_date)
    m = int(_num(months))

    # Calculate target month/year
    month = dt.month + m
    year = dt.year
    while month > 12:
        month -= 12
        year += 1
    while month < 1:
        month += 12
        year -= 1

    # Last day of target month
    last_day = monthrange(year, month)[1]
    result = datetime.datetime(year, month, last_day)
    return _python_to_excel_date(result)


def xl_edate(start_date, months):
    """EDATE function."""
    dt = _excel_date_to_python(start_date)
    m = int(_num(months))
    month = dt.month + m
    year = dt.year
    while month > 12:
        month -= 12
        year += 1
    while month < 1:
        month += 12
        year -= 1
    day = min(dt.day, monthrange(year, month)[1])
    result = datetime.datetime(year, month, day)
    return _python_to_excel_date(result)


def xl_date(year, month, day):
    """DATE function."""
    y = int(_num(year))
    m = int(_num(month))
    d = int(_num(day))
    # Handle month overflow
    while m > 12:
        m -= 12
        y += 1
    while m < 1:
        m += 12
        y -= 1
    d = min(d, monthrange(y, m)[1])
    return _python_to_excel_date(datetime.datetime(y, m, d))


def xl_year(serial):
    """YEAR function."""
    dt = _excel_date_to_python(serial)
    return dt.year


def xl_month(serial):
    """MONTH function."""
    dt = _excel_date_to_python(serial)
    return dt.month


def xl_day(serial):
    """DAY function."""
    dt = _excel_date_to_python(serial)
    return dt.day


def xl_today():
    """TODAY function."""
    return _python_to_excel_date(datetime.date.today())


def xl_now():
    """NOW function."""
    return _python_to_excel_date(datetime.datetime.now())


def xl_days(end_date, start_date):
    """DAYS function."""
    return _num(end_date) - _num(start_date)


# ============================================================
# Conditional Aggregation Functions
# ============================================================

def xl_sumif(range_vals, criteria, sum_range=None):
    """SUMIF function."""
    if sum_range is None:
        sum_range = range_vals
    if not isinstance(range_vals, (list, tuple)):
        range_vals = [range_vals]
    if not isinstance(sum_range, (list, tuple)):
        sum_range = [sum_range]

    crit = _parse_criteria(criteria)
    total = 0
    for i, val in enumerate(range_vals):
        if crit(val):
            if i < len(sum_range):
                total += _num(sum_range[i])
    return total


def xl_sumifs(sum_range, *args):
    """SUMIFS function."""
    if not isinstance(sum_range, (list, tuple)):
        sum_range = [sum_range]

    # args come in pairs: criteria_range, criteria
    criteria_pairs = []
    for j in range(0, len(args), 2):
        if j + 1 < len(args):
            cr = args[j] if isinstance(args[j], (list, tuple)) else [args[j]]
            criteria_pairs.append((cr, _parse_criteria(args[j + 1])))

    total = 0
    for i in range(len(sum_range)):
        match = True
        for cr, crit_fn in criteria_pairs:
            if i < len(cr):
                if not crit_fn(cr[i]):
                    match = False
                    break
            else:
                match = False
                break
        if match:
            total += _num(sum_range[i])
    return total


def xl_countif(range_vals, criteria):
    """COUNTIF function."""
    if not isinstance(range_vals, (list, tuple)):
        range_vals = [range_vals]
    crit = _parse_criteria(criteria)
    return sum(1 for v in range_vals if crit(v))


def xl_countifs(*args):
    """COUNTIFS function."""
    if len(args) < 2:
        return 0
    first_range = args[0] if isinstance(args[0], (list, tuple)) else [args[0]]
    criteria_pairs = []
    for j in range(0, len(args), 2):
        if j + 1 < len(args):
            cr = args[j] if isinstance(args[j], (list, tuple)) else [args[j]]
            criteria_pairs.append((cr, _parse_criteria(args[j + 1])))
    count = 0
    for i in range(len(first_range)):
        match = True
        for cr, crit_fn in criteria_pairs:
            if i < len(cr):
                if not crit_fn(cr[i]):
                    match = False
                    break
            else:
                match = False
                break
        if match:
            count += 1
    return count


def xl_averageif(range_vals, criteria, avg_range=None):
    """AVERAGEIF function."""
    if avg_range is None:
        avg_range = range_vals
    if not isinstance(range_vals, (list, tuple)):
        range_vals = [range_vals]
    if not isinstance(avg_range, (list, tuple)):
        avg_range = [avg_range]

    crit = _parse_criteria(criteria)
    vals = []
    for i, val in enumerate(range_vals):
        if crit(val) and i < len(avg_range):
            vals.append(_num(avg_range[i]))
    return sum(vals) / len(vals) if vals else 0


def xl_averageifs(avg_range, *args):
    """AVERAGEIFS function."""
    if not isinstance(avg_range, (list, tuple)):
        avg_range = [avg_range]
    criteria_pairs = []
    for j in range(0, len(args), 2):
        if j + 1 < len(args):
            cr = args[j] if isinstance(args[j], (list, tuple)) else [args[j]]
            criteria_pairs.append((cr, _parse_criteria(args[j + 1])))
    vals = []
    for i in range(len(avg_range)):
        match = True
        for cr, crit_fn in criteria_pairs:
            if i < len(cr):
                if not crit_fn(cr[i]):
                    match = False
                    break
            else:
                match = False
                break
        if match:
            vals.append(_num(avg_range[i]))
    return sum(vals) / len(vals) if vals else 0


def xl_maxifs(max_range, *args):
    """MAXIFS function."""
    if not isinstance(max_range, (list, tuple)):
        max_range = [max_range]
    criteria_pairs = []
    for j in range(0, len(args), 2):
        if j + 1 < len(args):
            cr = args[j] if isinstance(args[j], (list, tuple)) else [args[j]]
            criteria_pairs.append((cr, _parse_criteria(args[j + 1])))
    vals = []
    for i in range(len(max_range)):
        match = True
        for cr, crit_fn in criteria_pairs:
            if i < len(cr):
                if not crit_fn(cr[i]):
                    match = False
                    break
            else:
                match = False
                break
        if match:
            vals.append(_num(max_range[i]))
    return max(vals) if vals else 0


def xl_minifs(min_range, *args):
    """MINIFS function."""
    if not isinstance(min_range, (list, tuple)):
        min_range = [min_range]
    criteria_pairs = []
    for j in range(0, len(args), 2):
        if j + 1 < len(args):
            cr = args[j] if isinstance(args[j], (list, tuple)) else [args[j]]
            criteria_pairs.append((cr, _parse_criteria(args[j + 1])))
    vals = []
    for i in range(len(min_range)):
        match = True
        for cr, crit_fn in criteria_pairs:
            if i < len(cr):
                if not crit_fn(cr[i]):
                    match = False
                    break
            else:
                match = False
                break
        if match:
            vals.append(_num(min_range[i]))
    return min(vals) if vals else 0


def xl_sumproduct(*args):
    """SUMPRODUCT function."""
    arrays = []
    for a in args:
        if isinstance(a, (list, tuple)):
            arrays.append([_num(v) for v in a])
        else:
            arrays.append([_num(a)])

    if not arrays:
        return 0

    min_len = min(len(a) for a in arrays)
    total = 0
    for i in range(min_len):
        product = 1
        for arr in arrays:
            product *= arr[i]
        total += product
    return total


def xl_large(array, k):
    """LARGE function - k-th largest value."""
    if not isinstance(array, (list, tuple)):
        return _num(array)
    nums = sorted(_flatten_nums(array), reverse=True)
    idx = int(_num(k)) - 1
    if 0 <= idx < len(nums):
        return nums[idx]
    return 0


def xl_small(array, k):
    """SMALL function - k-th smallest value."""
    if not isinstance(array, (list, tuple)):
        return _num(array)
    nums = sorted(_flatten_nums(array))
    idx = int(_num(k)) - 1
    if 0 <= idx < len(nums):
        return nums[idx]
    return 0


# ============================================================
# Financial Functions
# ============================================================

def xl_npv(rate, *cashflows):
    """NPV function."""
    r = _num(rate)
    flat = _flatten_nums(*cashflows)
    total = 0
    for i, cf in enumerate(flat):
        total += cf / ((1 + r) ** (i + 1))
    return total


def xl_irr(cashflows, guess=0.1):
    """IRR function using Newton's method."""
    if not isinstance(cashflows, (list, tuple)):
        return 0

    cfs = [_num(v) for v in cashflows]
    rate = _num(guess)

    for _ in range(1000):
        npv = sum(cf / (1 + rate) ** i for i, cf in enumerate(cfs))
        dnpv = sum(-i * cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cfs))
        if abs(dnpv) < 1e-14:
            break
        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate

    return rate


def xl_xnpv(rate, cashflows, dates):
    """XNPV function."""
    r = _num(rate)
    if not isinstance(cashflows, (list, tuple)):
        cashflows = [cashflows]
    if not isinstance(dates, (list, tuple)):
        dates = [dates]

    cfs = [_num(v) for v in cashflows]
    ds = [_num(v) for v in dates]

    if not ds:
        return 0

    d0 = ds[0]
    total = 0
    for cf, d in zip(cfs, ds):
        years = (d - d0) / 365.0
        total += cf / ((1 + r) ** years)
    return total


def xl_xirr(cashflows, dates, guess=0.1):
    """XIRR function."""
    if not isinstance(cashflows, (list, tuple)):
        return 0
    if not isinstance(dates, (list, tuple)):
        return 0

    cfs = [_num(v) for v in cashflows]
    ds = [_num(v) for v in dates]
    rate = _num(guess)

    if not ds:
        return 0

    d0 = ds[0]

    for _ in range(1000):
        npv = 0
        dnpv = 0
        for cf, d in zip(cfs, ds):
            years = (d - d0) / 365.0
            denom = (1 + rate) ** years
            if abs(denom) < 1e-14:
                break
            npv += cf / denom
            dnpv += -years * cf / ((1 + rate) ** (years + 1))

        if abs(dnpv) < 1e-14:
            break
        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate

    return rate


def xl_pmt(rate, nper, pv, fv=0, payment_type=0):
    """PMT function."""
    r = _num(rate)
    n = _num(nper)
    p = _num(pv)
    f = _num(fv)
    pt = int(_num(payment_type))

    if r == 0:
        return -(p + f) / n if n != 0 else 0

    pmt = -(p * (r * (1 + r) ** n) + f * r) / (((1 + r) ** n - 1) * (1 + r * pt))
    return pmt


def xl_ppmt(rate, per, nper, pv, fv=0, payment_type=0):
    """PPMT function."""
    pmt_val = xl_pmt(rate, nper, pv, fv, payment_type)
    ipmt_val = xl_ipmt(rate, per, nper, pv, fv, payment_type)
    return pmt_val - ipmt_val


def xl_ipmt(rate, per, nper, pv, fv=0, payment_type=0):
    """IPMT function."""
    r = _num(rate)
    p = int(_num(per))
    n = _num(nper)
    pval = _num(pv)

    pmt_val = xl_pmt(rate, nper, pv, fv, payment_type)

    if p == 1:
        return -pval * r if payment_type == 0 else 0

    # Balance after p-1 payments
    balance = pval * (1 + r) ** (p - 1) + pmt_val * ((1 + r) ** (p - 1) - 1) / r
    return -balance * r


def xl_pv(rate, nper, pmt, fv=0, payment_type=0):
    """PV function."""
    r = _num(rate)
    n = _num(nper)
    p = _num(pmt)
    f = _num(fv)

    if r == 0:
        return -(p * n + f)

    return -(p * (1 - (1 + r) ** -n) / r + f * (1 + r) ** -n)


def xl_fv(rate, nper, pmt, pv=0, payment_type=0):
    """FV function."""
    r = _num(rate)
    n = _num(nper)
    p = _num(pmt)
    pval = _num(pv)

    if r == 0:
        return -(pval + p * n)

    return -(pval * (1 + r) ** n + p * ((1 + r) ** n - 1) / r)


def xl_nper(rate, pmt, pv, fv=0, payment_type=0):
    """NPER function."""
    r = _num(rate)
    p = _num(pmt)
    pval = _num(pv)
    f = _num(fv)

    if r == 0:
        return -(pval + f) / p if p != 0 else 0

    return math.log((-f * r + p) / (pval * r + p)) / math.log(1 + r)


def xl_rate(nper, pmt, pv, fv=0, payment_type=0, guess=0.1):
    """RATE function using Newton's method."""
    n = _num(nper)
    p = _num(pmt)
    pval = _num(pv)
    f = _num(fv)
    rate = _num(guess)

    for _ in range(1000):
        t = (1 + rate) ** n
        npv = pval * t + p * (t - 1) / rate + f
        dnpv = pval * n * (1 + rate) ** (n - 1) + p * (n * rate * (1 + rate) ** (n - 1) - t + 1) / (rate ** 2)

        if abs(dnpv) < 1e-14:
            break
        new_rate = rate - npv / dnpv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate

    return rate


def xl_transpose(array):
    """TRANSPOSE function."""
    return array  # Simplified


# ============================================================
# Criteria Parser Helper
# ============================================================

def _parse_criteria(criteria):
    """
    Parse an Excel criteria string into a callable.
    Supports: ">5", "<10", ">=3", "<=8", "<>0", "=text", etc.
    """
    import re as _re

    if callable(criteria):
        return criteria

    s = str(criteria).strip()

    m = _re.match(r'^(<>|>=|<=|>|<|=)\s*(.+)$', s)
    if m:
        op = m.group(1)
        val_str = m.group(2).strip()
        try:
            val = float(val_str)
        except ValueError:
            val = val_str

        if op == '>':
            return lambda x: _num(x) > val if isinstance(val, (int, float)) else str(x) > val
        elif op == '<':
            return lambda x: _num(x) < val if isinstance(val, (int, float)) else str(x) < val
        elif op == '>=':
            return lambda x: _num(x) >= val if isinstance(val, (int, float)) else str(x) >= val
        elif op == '<=':
            return lambda x: _num(x) <= val if isinstance(val, (int, float)) else str(x) <= val
        elif op == '<>':
            return lambda x: _num(x) != val if isinstance(val, (int, float)) else str(x) != val
        elif op == '=':
            return lambda x: _num(x) == val if isinstance(val, (int, float)) else str(x) == val

    # Plain value - exact match
    try:
        val = float(s)
        return lambda x: _num(x) == val
    except ValueError:
        return lambda x: str(x) == s
