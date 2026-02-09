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
    
        # Порог допустимости: strict = 1e-4, soft = 5%
        tol = 0.05 if soft else 0.0001
    
        best_values: Dict[str, float] | None = None
        best_obj: float = float("inf")
        best_iter: int | None = None
    
        def max_violation(violations: Dict[str, float]) -> float:
            return max(violations.values(), default=0.0)
    
        def objective(x: np.ndarray) -> float:
            current = self._update_values(base_values, x)
            result = self.engine.compute(current)
            value = float(result.key_outputs["J204"])
            penalty = 0.0
            if soft:
                penalties, _ = self._constraint_violations(current)
                # штрафуем нарушения квадратично
                penalty = sum(1000.0 * v**2 for v in penalties.values())
            return value + penalty
    
        def callback(x: np.ndarray) -> None:
            nonlocal best_values, best_obj, best_iter
    
            current = self._update_values(base_values, x)
            result = self.engine.compute(current)
    
            violations, _ = self._constraint_violations(current)
            mv = max_violation(violations)
    
            # пишем историю (как было)
            history.add(
                IterationRecord(
                    iteration=len(history.records) + 1,
                    variables={var: current[var] for var in self.variables},
                    objective=float(result.key_outputs["J204"]),
                    outputs={
                        **{key: float(result.key_outputs[key]) for key in KEY_OUTPUTS},
                        "J76": float(result.values["J76"]),
                        "J77": float(result.values["J77"]),
                    },
                    violations=violations,
                )
            )
    
            # сохраняем лучшее ДОПУСТИМОЕ решение (min J204)
            if mv <= tol:
                j204 = float(result.key_outputs["J204"])
                if j204 < best_obj:
                    best_obj = j204
                    best_values = dict(current)  # копия
                    best_iter = len(history.records)
    
        result = minimize(
            objective,
            initial,
            method="SLSQP",
            bounds=bounds,
            constraints=constraints,
            callback=callback,
        )
    
        # Финальная точка оптимизатора
        final_values = self._update_values(base_values, result.x)
        final_violations, final_statuses = self._constraint_violations(final_values)
        final_mv = max_violation(final_violations)
        final_obj = float(self.engine.compute(final_values).key_outputs["J204"])
    
        # Выбираем, что возвращать: best-feasible или final
        chosen_values = final_values
        chosen_statuses = final_statuses
        chosen_obj = final_obj
        chosen_mv = final_mv
        chosen_note = ""
    
        if best_values is not None:
            best_violations, best_statuses = self._constraint_violations(best_values)
            best_mv = max_violation(best_violations)
    
            # best должен быть реально допустимым по текущему режиму
            if best_mv <= tol:
                # выбираем меньший J204
                if best_obj < chosen_obj:
                    chosen_values = best_values
                    chosen_statuses = best_statuses
                    chosen_obj = best_obj
                    chosen_mv = best_mv
                    chosen_note = f" (лучшее допустимое: итерация {best_iter})"
    
        # Определяем успех
        # strict: считаем успехом, если есть допустимая точка (chosen_mv <= 1e-4)
        # soft: успех, если chosen_mv <= 5%
        if soft:
            success = chosen_mv <= 0.05
        else:
            success = chosen_mv <= 0.0001
    
        message = str(result.message) + chosen_note
    
        return OptimizationResult(
            success=success,
            message=message,
            values=chosen_values,
            objective=chosen_obj,
            constraints=chosen_statuses,
            history=history,
            stage="soft" if soft else "strict",
        )

    def _update_values(self, values: Dict[str, float], x: np.ndarray) -> Dict[str, float]:
        updated = dict(values)
        for idx, var in enumerate(self.variables):
            updated[var] = float(x[idx])
        return updated

    def _build_bounds(self, values: Dict[str, float]) -> List[Tuple[float, float]]:
        computed = self.engine.compute(values).values

        def get(cell: str) -> float:
            cell = normalize_cell(cell)
            if cell in values:
                return float(values[cell])
            if cell in computed:
                return float(computed[cell])
            defaults = self.engine.default_values()
            if cell in defaults:
                return float(defaults[cell])
            raise KeyError(cell)

        bounds_map = {
            "C7": (get("E195"), get("E197")),
            "J7": (get("J191"), get("J193")),
            "C9": (get("E191"), get("E193")),
            "J94": (get("J195"), get("J197")),
        }

        bounds: List[Tuple[float, float]] = []
        for var in self.variables:
            lo, hi = bounds_map[var]
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

    def _resolve_value(self, expr: str, values: Dict[str, float], computed: Dict[str, float] | None = None) -> float:
        expr = str(expr)
        try:
            return float(expr)
        except ValueError:
            pass

        key = normalize_cell(expr)
        if key in values:
            return float(values[key])

        if computed is None:
            computed = self.engine.compute(values).values

        if key in computed:
            return float(computed[key])

        defaults = self.engine.default_values()
        if key in defaults:
            return float(defaults[key])

        raise KeyError(key)

    def _constraint_violations(self, values: Dict[str, float]) -> Tuple[Dict[str, float], List[ConstraintStatus]]:
        violations: Dict[str, float] = {}
        statuses: List[ConstraintStatus] = []
        computed = self.engine.compute(values).values
        eps = 1e-6
        for idx, constraint in enumerate(self.schema.solver.constraints, start=1):
            lhs_val = self._resolve_value(constraint.lhs, values, computed)
            rhs_val = self._resolve_value(constraint.rhs, values, computed)
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
