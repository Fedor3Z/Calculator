from __future__ import annotations

import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from core.model import InputCell, normalize_cell
from solver.optimizer import ConstraintStatus

from export.export_xlsx import export_report


def _find_soffice() -> str | None:
    """Return path to LibreOffice/soffice executable if available."""
    for exe in ("soffice", "libreoffice", "soffice.exe", "libreoffice.exe"):
        p = shutil.which(exe)
        if p:
            return p
    return None


def _convert_xlsx_to_pdf(xlsx_path: Path, out_pdf_path: Path) -> None:
    """Convert XLSX to PDF via LibreOffice headless."""
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice (soffice) не найден. Установите LibreOffice или добавьте soffice в PATH."
        )

    out_dir = out_pdf_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice writes PDF into out_dir using the same basename.
    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(xlsx_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "Ошибка конвертации XLSX→PDF через LibreOffice.\n"
            + (proc.stderr or proc.stdout)
        )

    produced = out_dir / (xlsx_path.stem + ".pdf")
    if not produced.exists():
        raise RuntimeError("LibreOffice завершился без ошибки, но PDF не создан.")

    # Move/replace to the requested path
    if produced.resolve() != out_pdf_path.resolve():
        if out_pdf_path.exists():
            out_pdf_path.unlink()
        produced.replace(out_pdf_path)


def _fallback_reportlab_pdf(
    path: str,
    project_name: str,
    version: str,
    inputs: List[InputCell],
    values: Dict[str, float],
    key_outputs: Dict[str, float],
    constraints: List[ConstraintStatus],
    notes: str,
) -> None:
    """Fallback PDF export (text report)."""
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
        cell = normalize_cell(item.cell)
        y = draw_line(f"{item.name}: {values.get(cell)} {item.unit}", y)
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
    """Export PDF that visually matches the original Excel as close as possible.

    Preferred path:
    1) Create a temporary XLSX based on the bundled Excel template.
    2) Convert that XLSX to PDF via LibreOffice (headless).

    If LibreOffice is not available, falls back to a text PDF.
    """
    out_pdf = Path(path)

    with tempfile.TemporaryDirectory(prefix="kincalc_") as tmp:
        tmp_dir = Path(tmp)
        tmp_xlsx = tmp_dir / "export.xlsx"

        # Build XLSX that looks like original calculator.
        export_report(str(tmp_xlsx), inputs, values)

        try:
            _convert_xlsx_to_pdf(tmp_xlsx, out_pdf)
        except Exception:
            # Fallback to a simple report.
            _fallback_reportlab_pdf(path, project_name, version, inputs, values, key_outputs, constraints, notes)
