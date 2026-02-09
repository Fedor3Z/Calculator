from __future__ import annotations

from datetime import datetime
from typing import Dict, List

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from core.model import InputCell
from solver.optimizer import ConstraintStatus


def export_pdf(
    path: str,
    project_name: str,
    version: str,
    inputs: List[InputCell],
    values: Dict[str, float],
    key_outputs: Dict[str, float],
    constraints: List[ConstraintStatus],
    notes: str,
) -> None:
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4
    y = height - 20 * mm

    def draw_line(text: str, y_pos: float) -> float:
        c.drawString(20 * mm, y_pos, text)
        return y_pos - 6 * mm

    y = draw_line(f"Паспорт расчёта: {project_name}", y)
    y = draw_line(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}", y)
    y = draw_line(f"Версия приложения: {version}", y)
    y = draw_line("", y)
    y = draw_line("Входные данные:", y)
    for item in inputs:
        y = draw_line(f"{item.name} ({item.cell}): {values.get(normalize_cell(item.cell))} {item.unit}", y)
        if y < 40 * mm:
            c.showPage()
            y = height - 20 * mm

    y = draw_line("", y)
    y = draw_line("Ключевые результаты:", y)
    for cell, value in key_outputs.items():
        y = draw_line(f"{cell}: {value}", y)

    y = draw_line("", y)
    y = draw_line("Сводка ограничений:", y)
    for item in constraints:
        y = draw_line(f"{item.name}: {item.status} ({item.violation:.2%})", y)

    y = draw_line("", y)
    y = draw_line("Примечания:", y)
    for line in notes.splitlines() or ["-"]:
        y = draw_line(line, y)

    c.showPage()
    c.save()
