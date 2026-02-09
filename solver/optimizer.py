from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Tuple

import numpy as np
from scipy.optimize import minimize

from core.compute import ComputeEngine, KEY_OUTPUTS
from core.model import Schema
from solver.callbacks import IterationRecord, OptimizationHistory
from core.model import normalize_cell


@dataclass
class ConstraintStatus:
    name: str
    violation: float
    status: str


@dataclass
class OptimizationResult:
    success: bool
    message: str
    values: Dict[str, float]
    objective: float
    constraints: List[ConstraintStatus]
    history: OptimizationHistory
    stage: str


class Optimizer:
    def __init__(self, schema: Schema, engine: ComputeEngine) -> None:
        self.schema = schema
        self.engine = engine
        self.variables = schema.solver.variables

    def optimize(self, values: Dict[str, float]) -> OptimizationResult:
        initial = np.array([values[var] for var in self.variables], dtype=float)
        history = OptimizationHistory()
        strict_result = self._run_optimization(initial, values, history, soft=False)
        if strict_result.success:
            strict_result.stage = "strict"
            return strict_result
        soft_history = OptimizationHistory()
        soft_result = self._run_optimization(initial, values, soft_history, soft=True)
        soft_result.stage = "soft"
        return soft_result

    def _run_optimization(
        self,
        initial: np.ndarray,
        base_values: Dict[str, float],
        history: OptimizationHistory,
        soft: bool,
    ) -> OptimizationResult:
        bounds = self._build_bounds(base_values)
        constraints = self._build_constraints(base_values, soft=soft)

        def objective(x: np.ndarray) -> float:
            current = self._update_values(base_values, x)
            result = self.engine.compute(current)
            value = result.key_outputs["J204"]
            penalty = 0.0
            if soft:
                penalties, _ = self._constraint_violations(current)
                penalty = sum(1000.0 * v**2 for v in penalties.values())
            return value + penalty

        def callback(x: np.ndarray) -> None:
            current = self._update_values(base_values, x)
            result = self.engine.compute(current)
            violations, _ = self._constraint_violations(current)
            history.add(
                IterationRecord(
                    iteration=len(history.records) + 1,
                    variables={var: current[var] for var in self.variables},
                    objective=result.key_outputs["J204"],
                    outputs={
                        **{key: result.key_outputs[key] for key in KEY_OUTPUTS},
                        "J76": result.values["J76"],
                        "J77": result.values["J77"],
                    },
                    violations=violations,
                )
            )

        result = minimize(
            objective,
            initial,
            method="SLSQP",
            bounds=bounds,
            constraints=constraints,
            callback=callback,
        )
        final_values = self._update_values(base_values, result.x)
        violations, statuses = self._constraint_violations(final_values)
        acceptable = all(v <= 0.05 for v in violations.values())
        success = bool(result.success) and (acceptable if soft else result.success)
        return OptimizationResult(
            success=success,
            message=result.message,
            values=final_values,
            objective=self.engine.compute(final_values).key_outputs["J204"],
            constraints=statuses,
            history=history,
            stage="soft" if soft else "strict",
        )

    def _update_values(self, values: Dict[str, float], x: np.ndarray) -> Dict[str, float]:
        updated = dict(values)
        for idx, var in enumerate(self.variables):
            updated[var] = float(x[idx])
        return updated

    def _build_bounds(self, values: Dict[str, float]) -> List[Tuple[float, float]]:
        # values может содержать только вводы. Границы (E195/E197/J191/...) часто вычисляются.
        computed = self.engine.compute(values).values  # все ячейки после расчёта

        bounds_map = {
            "C7": (float(computed["E195"]), float(computed["E197"])),
            "J7": (float(computed["J191"]), float(computed["J193"])),
            "C9": (float(computed["E191"]), float(computed["E193"])),
            "J94": (float(computed["J195"]), float(computed["J197"])),
        }

    bounds: List[Tuple[float, float]] = []
    for var in self.variables:
        lo, hi = bounds_map[var]
        # на всякий случай: если перепутались местами из-за кривых данных
        if lo > hi:
            lo, hi = hi, lo
        bounds.append((lo, hi))
    return bounds
        
    def _build_constraints(self, values: Dict[str, float], soft: bool) -> List[Dict[str, object]]:
        if soft:
            return []
        constraints = []
        for constraint in self.schema.solver.constraints:
            lhs = constraint.lhs
            rhs = constraint.rhs

            def func(x: np.ndarray, lhs=lhs, rhs=rhs, ctype=constraint.type) -> float:
                current = self._update_values(values, x)
                left_val = self._resolve_value(lhs, current)
                right_val = self._resolve_value(rhs, current)
                if ctype == "le":
                    return right_val - left_val
                if ctype == "ge":
                    return left_val - right_val
                return left_val - right_val

            if constraint.type == "eq":
                constraints.append({"type": "eq", "fun": func})
            else:
                constraints.append({"type": "ineq", "fun": func})
        return constraints

    def _resolve_value(self, expr: str, values: Dict[str, float]) -> float:
        expr = str(expr)
        try:
            return float(expr)
        except ValueError:
            pass

        key = normalize_cell(expr)
        if key in values:
            return float(values[key])

        computed = self.engine.compute(values).values
        return float(computed[key])

    def _constraint_violations(self, values: Dict[str, float]) -> Tuple[Dict[str, float], List[ConstraintStatus]]:
        violations: Dict[str, float] = {}
        statuses: List[ConstraintStatus] = []
        computed = self.engine.compute(values).values
        eps = 1e-6
        for idx, constraint in enumerate(self.schema.solver.constraints, start=1):
            lhs_val = self._resolve_value(constraint.lhs, computed)
            rhs_val = self._resolve_value(constraint.rhs, computed)
            name = f"C{idx}: {constraint.lhs} {constraint.type} {constraint.rhs}"
            if constraint.type == "le":
                violation = max(0.0, (lhs_val - rhs_val) / max(abs(rhs_val), eps))
            elif constraint.type == "ge":
                violation = max(0.0, (rhs_val - lhs_val) / max(abs(lhs_val), eps))
            else:
                violation = abs(lhs_val - rhs_val) / max(abs(rhs_val), eps)
            violations[name] = violation
            status = "green" if violation <= 0.0001 else "yellow" if violation <= 0.05 else "red"
            statuses.append(ConstraintStatus(name=name, violation=violation, status=status))
        return violations, statuses
