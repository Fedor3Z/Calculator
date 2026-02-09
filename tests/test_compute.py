from pathlib import Path

from core.compute import ComputeEngine
from core.model import SchemaLoader


def test_compute_defaults():
    schema = SchemaLoader(Path("assets/kinematics_calc_extracted.json")).load()
    engine = ComputeEngine(schema)
    values = engine.default_values()
    result = engine.compute(values)

    assert result.key_outputs["J122"] == 13.0
    assert result.key_outputs["J132"] == 12.0
    assert result.key_outputs["L141"] == 5.0
    assert result.key_outputs["J169"] == 17.0
    assert abs(result.key_outputs["T168"] - 14.3) < 1e-4
    assert abs(result.key_outputs["T178"] - 14.4) < 1e-4
    assert abs(result.key_outputs["J204"] - 5.0) < 1e-4
