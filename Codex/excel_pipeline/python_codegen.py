from __future__ import annotations

import re
from collections import defaultdict, deque
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl.formula.tokenizer import Tokenizer

from .mapping_io import read_mapping_report
from .utils import is_ref_token, parse_cell_ref, split_sheet_ref

CELL_REF_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$")
CELL_RANGE_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$")
COL_RANGE_RE = re.compile(r"^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}$")
ROW_RANGE_RE = re.compile(r"^\$?\d+:\$?\d+$")


class PythonCodegenError(Exception):
    """Raised when a formula cannot be deterministically compiled to Python."""


@dataclass
class ExprNode:
    pass


@dataclass
class NumberNode(ExprNode):
    value: str


@dataclass
class TextNode(ExprNode):
    value: str


@dataclass
class BoolNode(ExprNode):
    value: bool


@dataclass
class EmptyNode(ExprNode):
    pass


@dataclass
class ErrorNode(ExprNode):
    value: str


@dataclass
class RefNode(ExprNode):
    sheet: str
    ref: str


@dataclass
class UnaryNode(ExprNode):
    op: str
    operand: ExprNode


@dataclass
class BinaryNode(ExprNode):
    op: str
    left: ExprNode
    right: ExprNode


@dataclass
class FuncNode(ExprNode):
    name: str
    args: list[ExprNode]


@dataclass
class CompiledFormula:
    key: tuple[str, str]
    function_name: str
    expression: str


_SUPPORTED_FUNC_TO_HELPER = {
    "SUM": "xl_sum",
    "NPV": "xl_npv",
    "IF": "xl_if",
    "IFERROR": "xl_iferror",
    "OR": "xl_or",
    "AND": "xl_and",
    "AVERAGE": "xl_average",
    "MAX": "xl_max",
    "MIN": "xl_min",
    "MOD": "xl_mod",
    "ROUNDUP": "xl_roundup",
    "ROUNDDOWN": "xl_rounddown",
    "YEAR": "xl_year",
    "MONTH": "xl_month",
    "TODAY": "xl_today",
    "NOW": "xl_now",
    "SUMPRODUCT": "xl_sumproduct",
}

_BINARY_OP_TO_HELPER = {
    "+": "xl_add",
    "-": "xl_sub",
    "*": "xl_mul",
    "/": "xl_div",
    "^": "xl_pow",
    "&": "xl_concat",
    "=": "xl_eq",
    "<>": "xl_ne",
    ">": "xl_gt",
    ">=": "xl_ge",
    "<": "xl_lt",
    "<=": "xl_le",
}

_PREFIX_OP_TO_HELPER = {
    "+": "xl_uplus",
    "-": "xl_uminus",
}

_POSTFIX_OP_TO_HELPER = {
    "%": "xl_percent",
}


class _FormulaParser:
    def __init__(self, formula: str, sheet: str, cell: str) -> None:
        self.sheet = sheet
        self.cell = cell
        self.tokens = [tok for tok in Tokenizer(formula).items if tok.type != "WSPACE"]
        self.pos = 0

    def parse(self) -> ExprNode:
        if not self.tokens:
            raise PythonCodegenError(f"{self.sheet}!{self.cell}: empty formula")
        expr = self._parse_expression(0)
        if self._peek() is not None:
            token = self._peek()
            raise PythonCodegenError(
                f"{self.sheet}!{self.cell}: unexpected token {token.value!r}"
            )
        return expr

    def _peek(self) -> Any:
        if self.pos >= len(self.tokens):
            return None
        return self.tokens[self.pos]

    def _advance(self) -> Any:
        token = self._peek()
        if token is None:
            raise PythonCodegenError(f"{self.sheet}!{self.cell}: unexpected end of formula")
        self.pos += 1
        return token

    def _parse_expression(self, min_bp: int) -> ExprNode:
        lhs = self._parse_prefix_or_primary()

        while True:
            token = self._peek()
            if token is None:
                break

            if token.type in {"SEP"}:
                break
            if token.type == "FUNC" and token.subtype == "CLOSE":
                break
            if token.type == "PAREN" and token.subtype == "CLOSE":
                break

            if token.type == "OPERATOR-POSTFIX":
                helper = _POSTFIX_OP_TO_HELPER.get(token.value)
                if helper is None:
                    raise PythonCodegenError(
                        f"{self.sheet}!{self.cell}: unsupported postfix operator {token.value!r}"
                    )
                lbp = 80
                if lbp < min_bp:
                    break
                self._advance()
                lhs = UnaryNode(op=token.value, operand=lhs)
                continue

            if token.type != "OPERATOR-INFIX":
                break

            op = token.value
            if op in {"=", "<>", ">", ">=", "<", "<="}:
                lbp, rbp = 10, 11
            elif op == "&":
                lbp, rbp = 20, 21
            elif op in {"+", "-"}:
                lbp, rbp = 30, 31
            elif op in {"*", "/"}:
                lbp, rbp = 40, 41
            elif op == "^":
                lbp, rbp = 50, 50
            else:
                raise PythonCodegenError(
                    f"{self.sheet}!{self.cell}: unsupported infix operator {op!r}"
                )

            if lbp < min_bp:
                break

            self._advance()
            rhs = self._parse_expression(rbp)
            lhs = BinaryNode(op=op, left=lhs, right=rhs)

        return lhs

    def _parse_prefix_or_primary(self) -> ExprNode:
        token = self._peek()
        if token is None:
            raise PythonCodegenError(f"{self.sheet}!{self.cell}: unexpected end of formula")

        if token.type == "OPERATOR-PREFIX":
            op = token.value
            helper = _PREFIX_OP_TO_HELPER.get(op)
            if helper is None:
                raise PythonCodegenError(
                    f"{self.sheet}!{self.cell}: unsupported prefix operator {op!r}"
                )
            self._advance()
            operand = self._parse_expression(70)
            return UnaryNode(op=op, operand=operand)

        if token.type == "PAREN" and token.subtype == "OPEN":
            self._advance()
            expr = self._parse_expression(0)
            close = self._advance()
            if close.type != "PAREN" or close.subtype != "CLOSE":
                raise PythonCodegenError(
                    f"{self.sheet}!{self.cell}: expected ')'"
                )
            return expr

        if token.type == "FUNC" and token.subtype == "OPEN":
            return self._parse_function()

        if token.type == "OPERAND":
            self._advance()
            return self._parse_operand(token)

        raise PythonCodegenError(
            f"{self.sheet}!{self.cell}: unsupported token {token.type}/{token.subtype} {token.value!r}"
        )

    def _parse_function(self) -> ExprNode:
        open_token = self._advance()
        raw_name = open_token.value[:-1] if open_token.value.endswith("(") else open_token.value
        name = _normalize_function_name(raw_name)

        args: list[ExprNode] = []

        next_token = self._peek()
        if next_token is not None and next_token.type == "FUNC" and next_token.subtype == "CLOSE":
            self._advance()
            return FuncNode(name=name, args=args)

        while True:
            args.append(self._parse_expression(0))
            separator = self._peek()
            if separator is not None and separator.type == "SEP":
                self._advance()
                continue
            break

        close = self._advance()
        if close.type != "FUNC" or close.subtype != "CLOSE":
            raise PythonCodegenError(
                f"{self.sheet}!{self.cell}: function {name} missing closing ')'"
            )

        return FuncNode(name=name, args=args)

    def _parse_operand(self, token: Any) -> ExprNode:
        subtype = token.subtype
        value = token.value

        if subtype == "NUMBER":
            return NumberNode(value=value)

        if subtype == "TEXT":
            text = value
            if text.startswith('"') and text.endswith('"') and len(text) >= 2:
                text = text[1:-1].replace('""', '"')
            return TextNode(value=text)

        if subtype == "LOGICAL":
            return BoolNode(value=str(value).strip().upper() == "TRUE")

        if subtype == "ERROR":
            return ErrorNode(value=str(value))

        if subtype == "RANGE":
            token_value = str(value).strip()
            ref_sheet, ref = split_sheet_ref(token_value)
            if not is_ref_token(ref):
                raise PythonCodegenError(
                    f"{self.sheet}!{self.cell}: unsupported reference token {token_value!r}"
                )
            return RefNode(sheet=ref_sheet or self.sheet, ref=ref.replace("$", ""))

        if subtype == "":
            return EmptyNode()

        raise PythonCodegenError(
            f"{self.sheet}!{self.cell}: unsupported operand {subtype} {value!r}"
        )


def _normalize_function_name(raw_name: str) -> str:
    name = raw_name.strip().upper()
    while name.startswith("@"):  # dynamic array implicit intersection marker
        name = name[1:]
    if name.startswith("_XLFN."):
        name = name.split(".", 1)[1]
    return name


def _emit_expression(node: ExprNode, sheet: str, cell: str) -> str:
    if isinstance(node, NumberNode):
        return node.value
    if isinstance(node, TextNode):
        return repr(node.value)
    if isinstance(node, BoolNode):
        return "True" if node.value else "False"
    if isinstance(node, EmptyNode):
        return "None"
    if isinstance(node, ErrorNode):
        return f"xl_error({node.value!r})"
    if isinstance(node, RefNode):
        if _is_cell_ref(node.ref):
            return f"xl_ref(ctx.cell({node.sheet!r}, {node.ref!r}))"
        return f"ctx.range({node.sheet!r}, {node.ref!r})"

    if isinstance(node, UnaryNode):
        helper = _POSTFIX_OP_TO_HELPER.get(node.op) or _PREFIX_OP_TO_HELPER.get(node.op)
        if helper is None:
            raise PythonCodegenError(f"{sheet}!{cell}: unsupported unary operator {node.op!r}")
        operand = _emit_expression(node.operand, sheet, cell)
        return f"{helper}({operand})"

    if isinstance(node, BinaryNode):
        helper = _BINARY_OP_TO_HELPER.get(node.op)
        if helper is None:
            raise PythonCodegenError(f"{sheet}!{cell}: unsupported binary operator {node.op!r}")
        left = _emit_expression(node.left, sheet, cell)
        right = _emit_expression(node.right, sheet, cell)
        return f"{helper}({left}, {right})"

    if isinstance(node, FuncNode):
        helper = _SUPPORTED_FUNC_TO_HELPER.get(node.name)
        if helper is None:
            raise PythonCodegenError(f"{sheet}!{cell}: unsupported function {node.name}")
        args = ", ".join(_emit_expression(arg, sheet, cell) for arg in node.args)
        return f"{helper}({args})" if args else f"{helper}()"

    raise PythonCodegenError(f"{sheet}!{cell}: unsupported AST node {type(node).__name__}")


def _is_cell_ref(ref: str) -> bool:
    clean = ref.replace("$", "")
    return bool(CELL_REF_RE.fullmatch(clean))


def _token_to_bounds(ref: str) -> tuple[int | None, int | None, int | None, int | None]:
    clean = ref.replace("$", "")

    if CELL_REF_RE.fullmatch(clean):
        row, col, _, _ = parse_cell_ref(clean)
        return (col, row, col, row)

    if ROW_RANGE_RE.fullmatch(clean):
        start, end = clean.split(":", 1)
        return (None, int(start), None, int(end))

    if COL_RANGE_RE.fullmatch(clean):
        left, right = clean.split(":", 1)
        left_row, left_col, _, _ = parse_cell_ref(f"{left}1")
        right_row, right_col, _, _ = parse_cell_ref(f"{right}1")
        _ = left_row
        _ = right_row
        return (left_col, None, right_col, None)

    if CELL_RANGE_RE.fullmatch(clean):
        left, right = clean.split(":", 1)
        left_row, left_col, _, _ = parse_cell_ref(left)
        right_row, right_col, _, _ = parse_cell_ref(right)
        min_col, max_col = sorted((left_col, right_col))
        min_row, max_row = sorted((left_row, right_row))
        return (min_col, min_row, max_col, max_row)

    raise PythonCodegenError(f"Unsupported reference for dependency analysis: {ref}")


def _is_in_bounds(
    row: int,
    col: int,
    min_col: int | None,
    min_row: int | None,
    max_col: int | None,
    max_row: int | None,
) -> bool:
    if min_col is not None and col < min_col:
        return False
    if max_col is not None and col > max_col:
        return False
    if min_row is not None and row < min_row:
        return False
    if max_row is not None and row > max_row:
        return False
    return True


def _extract_dependencies(
    formula: str,
    current_sheet: str,
    formula_cells_by_sheet: dict[str, set[tuple[int, int]]],
) -> set[tuple[str, int, int]]:
    dependencies: set[tuple[str, int, int]] = set()
    tokens = [tok for tok in Tokenizer(formula).items if tok.type != "WSPACE"]

    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue

        raw = str(token.value).strip()
        ref_sheet, ref = split_sheet_ref(raw)
        target_sheet = ref_sheet or current_sheet

        if not is_ref_token(ref):
            raise PythonCodegenError(
                f"{current_sheet}: unsupported reference token in dependency graph: {raw!r}"
            )

        candidates = formula_cells_by_sheet.get(target_sheet, set())
        if not candidates:
            continue

        min_col, min_row, max_col, max_row = _token_to_bounds(ref)
        for row, col in candidates:
            if _is_in_bounds(
                row=row,
                col=col,
                min_col=min_col,
                min_row=min_row,
                max_col=max_col,
                max_row=max_row,
            ):
                dependencies.add((target_sheet, row, col))

    return dependencies


def _safe_ident(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]+", "_", value)
    cleaned = cleaned.strip("_")
    if not cleaned:
        cleaned = "x"
    if cleaned[0].isdigit():
        cleaned = f"n_{cleaned}"
    return cleaned


def _topological_order(
    keys_in_order: list[tuple[str, str]],
    deps: dict[tuple[str, str], set[tuple[str, str]]],
) -> list[tuple[str, str]]:
    adjacency: dict[tuple[str, str], set[tuple[str, str]]] = defaultdict(set)
    indegree: dict[tuple[str, str], int] = {key: 0 for key in keys_in_order}

    for key, dep_keys in deps.items():
        for dep in dep_keys:
            adjacency[dep].add(key)
            indegree[key] += 1

    order_index = {key: idx for idx, key in enumerate(keys_in_order)}
    queue = deque(sorted((k for k, d in indegree.items() if d == 0), key=order_index.get))

    result: list[tuple[str, str]] = []
    while queue:
        key = queue.popleft()
        result.append(key)
        downstream = sorted(adjacency.get(key, set()), key=order_index.get)
        for neighbor in downstream:
            indegree[neighbor] -= 1
            if indegree[neighbor] == 0:
                queue.append(neighbor)

    if len(result) != len(keys_in_order):
        remaining = [key for key in keys_in_order if key not in set(result)]
        refs = ", ".join(f"{sheet}!{cell}" for sheet, cell in remaining[:20])
        raise PythonCodegenError(f"Circular or unresolved formula dependencies detected: {refs}")

    return result


def _build_script(
    mapping_report_path: Path,
    default_input_path: Path,
    default_output_path: Path,
    compiled: dict[tuple[str, str], CompiledFormula],
    order: list[tuple[str, str]],
) -> str:
    lines: list[str] = []
    lines.append("#!/usr/bin/env python3")
    lines.append("from __future__ import annotations")
    lines.append("")
    lines.append("import argparse")
    lines.append("from pathlib import Path")
    lines.append("")
    lines.append("from excel_pipeline.python_engine_support import (")
    lines.append("    execute_unstructured_python_engine,")
    lines.append("    xl_add,")
    lines.append("    xl_and,")
    lines.append("    xl_average,")
    lines.append("    xl_concat,")
    lines.append("    xl_div,")
    lines.append("    xl_eq,")
    lines.append("    xl_error,")
    lines.append("    xl_ge,")
    lines.append("    xl_gt,")
    lines.append("    xl_if,")
    lines.append("    xl_iferror,")
    lines.append("    xl_le,")
    lines.append("    xl_lt,")
    lines.append("    xl_max,")
    lines.append("    xl_min,")
    lines.append("    xl_mod,")
    lines.append("    xl_mul,")
    lines.append("    xl_ne,")
    lines.append("    xl_now,")
    lines.append("    xl_npv,")
    lines.append("    xl_or,")
    lines.append("    xl_percent,")
    lines.append("    xl_pow,")
    lines.append("    xl_ref,")
    lines.append("    xl_rounddown,")
    lines.append("    xl_roundup,")
    lines.append("    xl_sub,")
    lines.append("    xl_sum,")
    lines.append("    xl_sumproduct,")
    lines.append("    xl_today,")
    lines.append("    xl_uminus,")
    lines.append("    xl_uplus,")
    lines.append("    xl_month,")
    lines.append("    xl_year,")
    lines.append(")")
    lines.append("")
    lines.append(f"DEFAULT_MAPPING_REPORT = Path(r\"{str(mapping_report_path.resolve())}\")")
    lines.append(f"DEFAULT_INPUT = Path(r\"{str(default_input_path.resolve())}\")")
    lines.append(f"DEFAULT_OUTPUT = Path(r\"{str(default_output_path.resolve())}\")")
    lines.append("")

    for key in order:
        formula = compiled[key]
        lines.append(f"def {formula.function_name}(ctx):")
        lines.append(f"    return {formula.expression}")
        lines.append("")

    lines.append("FORMULA_FUNCS = {")
    for key in order:
        formula = compiled[key]
        sheet, cell = key
        lines.append(f"    ({sheet!r}, {cell!r}): {formula.function_name},")
    lines.append("}")
    lines.append("")

    lines.append("FORMULA_ORDER = [")
    for sheet, cell in order:
        lines.append(f"    ({sheet!r}, {cell!r}),")
    lines.append("]")
    lines.append("")

    lines.append(
        "def run(mapping_report_path: Path = DEFAULT_MAPPING_REPORT, "
        "inputs_path: Path = DEFAULT_INPUT, "
        "output_path: Path = DEFAULT_OUTPUT) -> Path:"
    )
    lines.append("    return execute_unstructured_python_engine(")
    lines.append("        mapping_report_path=mapping_report_path,")
    lines.append("        unstructured_input_path=inputs_path,")
    lines.append("        output_path=output_path,")
    lines.append("        formula_order=FORMULA_ORDER,")
    lines.append("        formula_funcs=FORMULA_FUNCS,")
    lines.append("    )")
    lines.append("")

    lines.append("def main() -> None:")
    lines.append("    parser = argparse.ArgumentParser(")
    lines.append("        description=\"Run generated Python formula engine on unstructured inputs\"")
    lines.append("    )")
    lines.append("    parser.add_argument('--mapping', type=Path, default=DEFAULT_MAPPING_REPORT)")
    lines.append("    parser.add_argument('--inputs', type=Path, default=DEFAULT_INPUT)")
    lines.append("    parser.add_argument('--output', type=Path, default=DEFAULT_OUTPUT)")
    lines.append("    args = parser.parse_args()")
    lines.append("    run(mapping_report_path=args.mapping, inputs_path=args.inputs, output_path=args.output)")
    lines.append("")
    lines.append("if __name__ == '__main__':")
    lines.append("    main()")

    return "\n".join(lines) + "\n"


def generate_unstructured_python_engine(
    mapping_report_path: Path,
    output_script_path: Path,
    default_input_path: Path,
    default_output_path: Path,
) -> Path:
    model = read_mapping_report(mapping_report_path)

    formula_records: list[Any] = []
    for sheet in model.sheet_order:
        records = sorted(model.cells_by_sheet.get(sheet, []), key=lambda r: (r.row, r.column))
        for record in records:
            if not record.formula:
                continue
            formula_records.append(record)

    if not formula_records:
        raise PythonCodegenError("No formula records found in mapping report")

    formula_cells_by_sheet: dict[str, set[tuple[int, int]]] = defaultdict(set)

    for record in formula_records:
        if record.formula == "__SPECIAL_FORMULA__":
            raise PythonCodegenError(
                f"{record.sheet}!{record.cell}: special formula type is unsupported in Python codegen"
            )
        if not isinstance(record.formula, str) or not record.formula.startswith("="):
            raise PythonCodegenError(
                f"{record.sheet}!{record.cell}: formula text is missing or invalid"
            )
        formula_cells_by_sheet[record.sheet].add((record.row, record.column))

    compile_errors: list[str] = []
    compiled: dict[tuple[str, str], CompiledFormula] = {}
    dependencies: dict[tuple[str, str], set[tuple[str, str]]] = {}
    ordered_keys: list[tuple[str, str]] = []

    for record in formula_records:
        key = (record.sheet, record.cell)
        ordered_keys.append(key)

        try:
            parser = _FormulaParser(record.formula, record.sheet, record.cell)
            ast = parser.parse()
            expression = _emit_expression(ast, record.sheet, record.cell)
            dep_row_col = _extract_dependencies(
                formula=record.formula,
                current_sheet=record.sheet,
                formula_cells_by_sheet=formula_cells_by_sheet,
            )
            dep_keys = {
                (sheet, _cell_from_row_col(row, col))
                for sheet, row, col in dep_row_col
            }
            dep_keys.discard(key)

            function_name = f"calc_{_safe_ident(record.sheet)}_{_safe_ident(record.cell)}"
            compiled[key] = CompiledFormula(
                key=key,
                function_name=function_name,
                expression=expression,
            )
            dependencies[key] = dep_keys
        except Exception as exc:
            compile_errors.append(f"{record.sheet}!{record.cell}: {exc}")

    if compile_errors:
        sample = "\n".join(compile_errors[:50])
        if len(compile_errors) > 50:
            sample += f"\n... and {len(compile_errors) - 50} more"
        raise PythonCodegenError(
            "Failed to generate deterministic Python formula engine:\n" + sample
        )

    order = _topological_order(
        keys_in_order=ordered_keys,
        deps=dependencies,
    )

    script = _build_script(
        mapping_report_path=mapping_report_path,
        default_input_path=default_input_path,
        default_output_path=default_output_path,
        compiled=compiled,
        order=order,
    )

    output_script_path.parent.mkdir(parents=True, exist_ok=True)
    output_script_path.write_text(script, encoding="utf-8")
    return output_script_path


def _cell_from_row_col(row: int, col: int) -> str:
    letters = ""
    current = col
    while current > 0:
        current, rem = divmod(current - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"
