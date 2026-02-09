from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Set, Tuple

from .model import normalize_cell

CELL_REF_RE = re.compile(r'(?<![A-Z0-9_"])\$?[A-Z]{1,3}\$?\d+(?![A-Z0-9_"])')
RANGE_RE = re.compile(r'(\$?[A-Z]{1,3}\$?\d+)\s*:\s*(\$?[A-Z]{1,3}\$?\d+)')


@dataclass(frozen=True)
class ParsedFormula:
    python_expr: str
    dependencies: Set[str]


FUNCTION_MAP = {
    "IF": "if_func",
    "OR": "or_func",
    "AND": "and_func",
    "MAX": "max",
    "MIN": "min",
    "ABS": "abs",
    "SQRT": "sqrt",
    "EXP": "exp",
    "COS": "cos",
    "SIN": "sin",
    "TAN": "tan",
    "ASIN": "asin",
    "ACOS": "acos",
    "ATAN": "atan",
    "PI": "pi_func",
    "SUM": "sum",
}

def normalize_comparisons(expr: str) -> str:
    # Excel: <>  -> Python: !=
    expr = expr.replace("<>", "!=")
    # Excel: =   -> Python: ==  (только одиночный "=", не трогая <= >= ==)
    expr = re.sub(r'(?<![<>=])=(?![<>=])', '==', expr)
    return expr

def parse_formula(formula: str) -> ParsedFormula:
    expr = formula.lstrip("=")
    expr = expr.replace(";", ",")
    expr = expr.replace("^", "**")
    expr = normalize_comparisons(expr)

    dependencies: Set[str] = set()

    def range_replacer(match: re.Match[str]) -> str:
        start = normalize_cell(match.group(1))
        end = normalize_cell(match.group(2))
        dependencies.update({start, end})
        return f'range_("{start}", "{end}")'

    expr = RANGE_RE.sub(range_replacer, expr)

    def cell_replacer(match: re.Match[str]) -> str:
        cell = normalize_cell(match.group(0))
        dependencies.add(cell)
        return f'cell_("{cell}")'

    expr = CELL_REF_RE.sub(cell_replacer, expr)

    for excel_name, py_name in FUNCTION_MAP.items():
        expr = re.sub(rf"\b{excel_name}\b", py_name, expr, flags=re.IGNORECASE)

    return ParsedFormula(python_expr=expr, dependencies=dependencies)


def extract_dependencies(formula: str) -> Set[str]:
    return parse_formula(formula).dependencies
