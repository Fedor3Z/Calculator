from pathlib import Path

import pytest

from core.compute import ComputeEngine
from core.model import SchemaLoader

pytest.importorskip("numpy")
pytest.importorskip("scipy")

from solver.optimizer import Optimizer


def test_optimizer_runs():
    schema = SchemaLoader(Path("assets/kinematics_calc_extracted.json")).load()
    engine = ComputeEngine(schema)
    optimizer = Optimizer(schema, engine)
    values = engine.default_values()
    result = optimizer.optimize(values)

    assert result is not None
    assert len(result.constraints) == len(schema.solver.constraints)
    assert result.history.records is not None
