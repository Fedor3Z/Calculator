from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class IterationRecord:
    iteration: int
    variables: Dict[str, float]
    objective: float
    outputs: Dict[str, float]
    violations: Dict[str, float]


@dataclass
class OptimizationHistory:
    records: List[IterationRecord] = field(default_factory=list)

    def add(self, record: IterationRecord) -> None:
        self.records.append(record)

    def to_dict(self) -> List[Dict[str, object]]:
        return [
            {
                "iteration": rec.iteration,
                "variables": rec.variables,
                "objective": rec.objective,
                "outputs": rec.outputs,
                "violations": rec.violations,
            }
            for rec in self.records
        ]
