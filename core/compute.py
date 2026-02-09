from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

from .dependency_graph import build_dependency_graph, topological_sort
from .evaluator import FormulaError, evaluate_formula
from .model import Schema, build_default_cell_values, normalize_cell, schema_formulas_by_cell


@dataclass
class ComputeResult:
    values: Dict[str, float]
    key_outputs: Dict[str, float]
    table_rows: List[Tuple[float, float, float, float]]


KEY_OUTPUTS = ["J122", "J132", "L141", "J169", "T168", "T178", "J204"]
TABLE_RANGE = {
    "k": "N106:N116",
    "a": "O106:O116",
    "b": "P106:P116",
    "s": "Q106:Q116",
}


class ComputeEngine:
    def __init__(self, schema: Schema) -> None:
        self.schema = schema
        self.formulas = schema_formulas_by_cell(schema)
        self.graph = build_dependency_graph({cell: f.formula for cell, f in self.formulas.items()})
        self.order = topological_sort(self.graph)

    def default_values(self) -> Dict[str, float]:
        return build_default_cell_values(self.schema)

    def compute(self, values: Dict[str, float]) -> ComputeResult:
        data = {normalize_cell(k): float(v) for k, v in values.items()}
        for cell in self.order:
            formula = self.formulas[cell].formula
            try:
                data[cell] = evaluate_formula(formula, data)
            except FormulaError as exc:
                raise FormulaError(f"Ошибка в ячейке {cell}: {exc}") from exc
        key_outputs = {cell: data.get(cell) for cell in KEY_OUTPUTS}
        table_rows = self._build_table(data)
        return ComputeResult(values=data, key_outputs=key_outputs, table_rows=table_rows)

    def _build_table(self, data: Dict[str, float]) -> List[Tuple[float, float, float, float]]:
        rows = []
        for row in range(106, 117):
            row_values = (
                data.get(f"N{row}"),
                data.get(f"O{row}"),
                data.get(f"P{row}"),
                data.get(f"Q{row}"),
            )
            rows.append(row_values)
        return rows
