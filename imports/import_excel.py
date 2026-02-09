from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

from openpyxl import load_workbook


@dataclass
class ImportResult:
    values: Dict[str, float]
    updated: Dict[str, Tuple[float, float]]
    errors: List[str]


def import_from_excel(path: str, input_cells: List[str], current: Dict[str, float]) -> ImportResult:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    errors: List[str] = []
    updated: Dict[str, Tuple[float, float]] = {}
    values = dict(current)

    for cell in input_cells:
        value = sheet[cell].value
        if value is None or not isinstance(value, (int, float)):
            errors.append(cell)
            continue
        prev = values.get(cell, 0.0)
        values[cell] = float(value)
        if prev != values[cell]:
            updated[cell] = (prev, values[cell])

    return ImportResult(values=values, updated=updated, errors=errors)
