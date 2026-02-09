from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional


@dataclass(frozen=True)
class InputCell:
    cell: str
    name: str
    default: float
    unit: str
    description: str


@dataclass(frozen=True)
class FormulaCell:
    cell: str
    formula: str


@dataclass(frozen=True)
class SolverConstraint:
    type: str
    lhs: str
    rhs: str


@dataclass(frozen=True)
class SolverSpec:
    objective: str
    variables: List[str]
    constraints: List[SolverConstraint]


@dataclass(frozen=True)
class Schema:
    inputs: List[InputCell]
    formulas: List[FormulaCell]
    solver: SolverSpec
    version: str


class SchemaLoader:
    def __init__(self, path: Path) -> None:
        self.path = path

    def load(self) -> Schema:
        data = json.loads(self.path.read_text(encoding="utf-8"))
        inputs = [
            InputCell(
                cell=item["cell"],
                name=item["name"],
                default=float(item["default"]),
                unit=item.get("unit", ""),
                description=item.get("description", ""),
            )
            for item in data["inputs"]
        ]
        formulas = [FormulaCell(cell=item["cell"], formula=item["formula"]) for item in data["formulas"]]
        solver_data = data["solver"]
        solver = SolverSpec(
            objective=solver_data["objective"],
            variables=solver_data["variables"],
            constraints=[
                SolverConstraint(type=item["type"], lhs=item["lhs"], rhs=item["rhs"])
                for item in solver_data["constraints"]
            ],
        )
        return Schema(inputs=inputs, formulas=formulas, solver=solver, version=data.get("version", ""))


def build_default_cell_values(schema: Schema) -> Dict[str, float]:
    values: Dict[str, float] = {}
    for item in schema.inputs:
        values[normalize_cell(item.cell)] = float(item.default)
    return values


def normalize_cell(cell: str) -> str:
    return cell.replace("$", "").upper()


def schema_inputs_by_cell(schema: Schema) -> Dict[str, InputCell]:
    return {normalize_cell(item.cell): item for item in schema.inputs}


def schema_formulas_by_cell(schema: Schema) -> Dict[str, FormulaCell]:
    return {normalize_cell(item.cell): item for item in schema.formulas}
