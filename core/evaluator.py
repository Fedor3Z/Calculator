from __future__ import annotations

import math
from typing import Callable, Dict, Iterable, List

from .formula_parser import ParsedFormula, parse_formula


class FormulaError(RuntimeError):
    pass


def if_func(condition: float, true_val: float, false_val: float) -> float:
    return true_val if condition else false_val


def or_func(*args: float) -> bool:
    return any(args)


def and_func(*args: float) -> bool:
    return all(args)


def pi_func() -> float:
    return math.pi


def build_eval_context(values: Dict[str, float]) -> Dict[str, Callable[..., float]]:
    def cell_(cell: str) -> float:
        if cell not in values:
            raise FormulaError(f"Отсутствует значение для ячейки {cell}")
        return float(values[cell])

    def range_(start: str, end: str) -> List[float]:
        start_col, start_row = split_cell(start)
        end_col, end_row = split_cell(end)
        cols = range(min(start_col, end_col), max(start_col, end_col) + 1)
        rows = range(min(start_row, end_row), max(start_row, end_row) + 1)
        result = []
        for row in rows:
            for col in cols:
                cell_name = f"{col_to_letters(col)}{row}"
                result.append(cell_(cell_name))
        return result

    return {
        "cell_": cell_,
        "range_": range_,
        "if_func": if_func,
        "or_func": or_func,
        "and_func": and_func,
        "sqrt": math.sqrt,
        "exp": math.exp,
        "cos": math.cos,
        "sin": math.sin,
        "tan": math.tan,
        "asin": math.asin,
        "acos": math.acos,
        "atan": math.atan,
        "abs": abs,
        "max": max,
        "min": min,
        "sum": sum,
        "pi_func": pi_func,
    }


def evaluate_formula(formula: str, values: Dict[str, float]) -> float:
    parsed = parse_formula(formula)
    context = build_eval_context(values)
    try:
        return float(eval(parsed.python_expr, {"__builtins__": {}}, context))
    except FormulaError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise FormulaError(f"Ошибка вычисления формулы '{formula}': {exc}") from exc


def split_cell(cell: str) -> tuple[int, int]:
    letters = "".join([c for c in cell if c.isalpha()])
    numbers = "".join([c for c in cell if c.isdigit()])
    return letters_to_col(letters), int(numbers)


def letters_to_col(letters: str) -> int:
    col = 0
    for char in letters.upper():
        col = col * 26 + (ord(char) - ord("A") + 1)
    return col


def col_to_letters(col: int) -> str:
    letters = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters
