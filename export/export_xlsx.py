from __future__ import annotations

from pathlib import Path
import sys
from typing import Dict, List, Optional

from openpyxl import load_workbook

from core.model import InputCell, normalize_cell
from solver.optimizer import ConstraintStatus


def _template_path() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "assets" / "excel_template.xlsx"


def export_report(
    path: str,
    inputs: List[InputCell],
    values: Dict[str, float],
    key_outputs: Optional[Dict[str, float]] = None,
    formulas: Optional[Dict[str, str]] = None,
    constraints: Optional[List[ConstraintStatus]] = None,
) -> None:
    """Экспорт в XLSX в виде исходного Excel-калькулятора.

    Мы берём Excel-шаблон (excel_template.xlsx), записываем в него входные ячейки,
    а формулы и оформление остаются как в оригинале.

    Параметры key_outputs/formulas/constraints оставлены для обратной совместимости
    с вызовами из интерфейса.
    """
    wb = load_workbook(_template_path())
    ws = wb.active

    for item in inputs:
        cell = item.cell
        key = normalize_cell(cell)
        if key in values and values[key] is not None:
            ws[cell].value = float(values[key])

    wb.save(path)
