from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import re
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, QLocale
from PySide6.QtGui import QAction, QDoubleValidator, QPixmap, QFont, QFontMetrics
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStyle,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from app import VERSION
from core.compute import ComputeEngine, KEY_OUTPUTS
from core.model import InputCell, Schema, normalize_cell, schema_formulas_by_cell, schema_inputs_by_cell
from export.export_pdf import export_pdf
from export.export_xlsx import export_report
from imports.import_excel import import_from_excel
from solver.optimizer import Optimizer
from storage.projects import ProjectData, ProjectStorage
from storage.recents import RecentProjects


# Пояснения для параметров (показываются при наведении на '?' в кружочке).
# Ключи — ровно то, что выводится в лейбле параметра (например: "D1", "Q'", "Tогранич").
PARAM_TOOLTIPS: Dict[str, str] = {
    "D1": "Диаметр ведущего шкива.",
    "D2": "Диаметр ведомого шкива.",
    "n1": "Частота вращения ведущего шкива.",
    "P": "Мощность электродвигателя.",
    "Q'": "Результирующая нагрузка от ременной передачи при установке на комбайне.",
    "M'": "Крутящий момент на валу при работе в поле.",
    "Sэл": "Стоимость электроэнергии.",
    "L": "Длина ремня.",
    "b1": "Высота установки ведущего шкива от пола.",
    "b2": "Высота установки ведомого шкива от пола.",
    "μ": "Коэффициент трения между шкивом и ремнем.",
    "kрем": "Жесткость ремня.",
    "Fпров": "Сила проверки натяжения ремня.",
    "L1": "Длина меньшего плеча рычага.",
    "L2": "Длина большего плеча рычага.",
    "s2": "Расстояние между ведомым шкивом и подшипником нагружения.",
    "s3": "Расстояние между подшипником нагружения и подшипниковой опорой барабана.",
    "s4": "Расстояние между подшипниковыми опорами барабана.",
    "Se": "Предел выносливости при симметричном цикле нагружения.",
    "σв": "Временное сопротивление разрыву.",
    "ks": (
        "Коэффициент состояния поверхности (чем грубее поверхность, тем ниже выносливость):\n"
        "• шлифованная/полированная: 0,85…1,00\n"
        "• чистовая мехобработка: 0,75…0,90\n"
        "• грубая мехобработка: 0,65…0,80\n"
        "• прокат/ковка (без обработки): 0,45…0,70"
    ),
    "kr": (
        "Коэффициент размера (крупные диаметры/сечения дают меньшую выносливость):\n"
        "• малые диаметры (условно до ~10 мм): ~1,0\n"
        "• 20…50 мм: 0,8…0,95\n"
        "• 50…200+ мм: 0,6…0,8"
    ),
    "kt": (
        "Температурный коэффициент (учитывает снижение выносливости при повышенной температуре):\n"
        "• до ~100 °C: ~1,0\n"
        "• 150…250 °C: 0,9…1,0\n"
        "• 300…400 °C: 0,7…0,9"
    ),
    "krel": (
        "Коэффициент надёжности (вероятности безотказной работы по выносливости).\n"
        "Чем выше требуемая надёжность, тем меньше допустимая выносливость:\n"
        "• 50%: 1,00\n"
        "• 90%: 0,90\n"
        "• 95%: 0,87\n"
        "• 99%: 0,81\n"
        "• 99,9%: 0,75"
    ),
    "kf": (
        "Коэффициент усталостной концентрации напряжений (учитывает геометрический концентратор + "
        "чувствительность материала к надрезу):\n"
        "• гладкая деталь без надрезов: 1,0\n"
        "• плавные переходы/галтели: 1,1…1,6\n"
        "• шпоночные пазы/резкие уступы: 1,6…3,0\n"
        "• очень острые надрезы: до 4…5+ (реже в “нормальном” машиностроении)"
    ),
    "Fр": "Преднатяжение ветвей ремня.",
    "D1min": "Минимальный диаметр ведущего шкива.",
    "D1max": "Максимальный диаметр ведущего шкива.",
    "D2min": "Минимальный диаметр ведомого шкива.",
    "D2max": "Максимальный диаметр ведомого шкива.",
    "n1min": "Минимальная частота вращения ведущего шкива.",
    "n1max": "Максимальная частота вращения ведущего шкива.",
    "Fpmin": "Минимальное преднатяжение ветвей ремня.",
    "Fpmax": "Максимальное преднатяжение ветвей ремня.",
    "Tогранич": "Максимальное натяжение ветви ремня.",
    "Sнач": "Начальная стоимость испытаний (можно взять после 1-го расчета).",
    "Fнач": "Начальная требуемая внешняя радиальная сила (можно взять после 1-го расчета).",
    "wст": "Вес начальной стоимости испытаний (для оптимизации).",
    "wF": "Вес начальной требуемой внешней радиальной силы (для оптимизации).",
}


@dataclass
class UiState:
    project_path: Optional[Path] = None
    notes: str = ""
    theme: str = "light"


LIGHT_STYLESHEET = """
QWidget { }
QLineEdit, QTableWidget {
    background: #ffffff;
    color: #111111;
}
QGroupBox {
    font-weight: 600;
}
"""


DARK_STYLESHEET = """
QWidget {
    background-color: #2b2b2b;
    color: #f0f0f0;
}
QLineEdit, QTableWidget {
    background: #1f1f1f;
    color: #f0f0f0;
    border: 1px solid #3b3b3b;
}
QHeaderView::section {
    background-color: #1f1f1f;
    color: #f0f0f0;
    border: 1px solid #3b3b3b;
}
QGroupBox {
    font-weight: 600;
    border: 1px solid #3b3b3b;
    margin-top: 10px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px 0 3px;
}
"""


class MatplotlibCanvas(FigureCanvas):
    def __init__(self) -> None:
        self.figure = Figure(figsize=(5, 3))
        self.ax = self.figure.add_subplot(111)
        self._qlocale = QLocale("ru_RU")
        self._qfont: QFont | None = None
        super().__init__(self.figure)

    def set_qt_style(self, font: QFont, locale: QLocale) -> None:
        self._qfont = font
        self._qlocale = locale

    def _apply_qt_style(self) -> None:
        if self._qfont is None:
            return
        import matplotlib as mpl

        size = float(self._qfont.pointSizeF() or 10)
        mpl.rcParams["font.family"] = self._qfont.family()
        mpl.rcParams["font.size"] = size

    def plot(self, x: List[float], y: List[float], title: str, xlabel: str, ylabel: str, y_decimals: int = 3) -> None:
        from matplotlib.ticker import FuncFormatter

        self._apply_qt_style()
        self.ax.clear()
        self.ax.plot(x, y, marker="o")
        self.ax.set_title(title)
        self.ax.set_xlabel(xlabel)
        self.ax.set_ylabel(ylabel)

        self.ax.xaxis.set_major_formatter(
            FuncFormatter(
                lambda v, p: str(int(v)) if float(v).is_integer() else self._qlocale.toString(v, "f", 1)
            )
        )
        self.ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: self._qlocale.toString(v, "f", y_decimals)))

        if self._qfont is not None:
            fs = float(self._qfont.pointSizeF() or 10)
            self.ax.tick_params(labelsize=fs)

        self.figure.tight_layout()
        self.draw()

    def plot_multi(self, series: Dict[str, List[float]], x: List[float], x_decimals: int = 1, y_decimals: int = 3) -> None:
        from matplotlib.ticker import FuncFormatter

        self._apply_qt_style()
        self.ax.clear()
        for label, values in series.items():
            self.ax.plot(x, values, label=label)
        self.ax.legend()
        self.ax.set_title("Профили цели")
        self.ax.set_xlabel("Отклонение, %")
        self.ax.set_ylabel("J")

        self.ax.xaxis.set_major_formatter(FuncFormatter(lambda v, p: self._qlocale.toString(v, "f", x_decimals)))
        self.ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: self._qlocale.toString(v, "f", y_decimals)))

        if self._qfont is not None:
            fs = float(self._qfont.pointSizeF() or 10)
            self.ax.tick_params(labelsize=fs)

        self.figure.tight_layout()
        self.draw()


class MainWindow(QMainWindow):
    """Основное окно.

    Требования пользователя:
    - Только вкладки: Расчет / Оптимизация / Профили цели
    - Без режима "Инженерный" (оптимизация всегда доступна)
    - Десятичная запятая
    - Единый шрифт (в т.ч. графики)
    - Формулы без искусственного увеличения
    - Округления: 1 знак для всех, кроме J76/J77/J204 (3 знака)
    """

    SPECIAL_3DP = {"J76", "J77", "J204"}

    # Пояснения для параметров (для '?' в кружочке)
    HELP_TEXTS: Dict[str, str] = {
        'D1': 'Диаметр ведущего шкива',
        'D2': 'Диаметр ведомого шкива',
        'n1': 'Частота вращения ведущего шкива',
        'P': 'Мощность электродвигателя',
        "Q'": 'Результирующая нагрузка от ременной передачи при установке на комбайне',
        "M'": 'Крутящий момент на валу при работе в поле',
        'Sэл': 'Стоимость электроэнергии',
        'L': 'Длина ремня',
        'b1': 'Высота установки ведущего шкива от пола',
        'b2': 'Высота установки ведомого шкива от пола',
        'μ': 'Коэффициент трения между шкивом и ремнем',
        'kрем': 'Жесткость ремня',
        'Fпров': 'Сила проверки натяжения ремня',
        'L1': 'Длина меньшего плеча рычага',
        'L2': 'Длина большего плеча рычага',
        's2': 'Расстояние между ведомым шкивом и подшипником нагружения',
        's3': 'Расстояние между подшипником нагружения и подшипниковой опорой барабана',
        's4': 'Расстояние между подшипниковыми опорами барабана',
        'Se': 'Предел выносливости при симметричном цикле нагружения',
        'σв': 'Временное сопротивление разрыву',
        'ks': 'Коэффициент состояния поверхности (чем грубее поверхность, тем ниже выносливость):\n  шлифованная/полированная: 0,85…1,00\n  чистовая мехобработка: 0,75…0,90\n  грубая мехобработка: 0,65…0,80\n  прокат/ковка (без обработки): 0,45…0,70',
        'kr': 'Коэффициент размера (крупные диаметры/сечения дают меньшую выносливость):\n  малые диаметры (условно до ~10 мм): ~1,0\n  20…50 мм: 0,8…0,95\n  50…200+ мм: 0,6…0,8',
        'kt': 'Температурный коэффициент (учитывает снижение выносливости при повышенной температуре):\n  до ~100 °C: ~1,0\n  150…250 °C: 0,9…1,0\n  300…400 °C: 0,7…0,9',
        'krel': 'Коэффициент надёжности (вероятности безотказной работы по выносливости). Чем выше требуемая надёжность, тем меньше допустимая выносливость:\n  50%: 1,00\n  90%: 0,90\n  95%: 0,87\n  99%: 0,81\n  99,9%: 0,75',
        'kf': 'Коэффициент усталостной концентрации напряжений (учитывает геометрический концентратор + чувствительность материала к надрезу):\n  гладкая деталь без надрезов: 1,0\n  плавные переходы/галтели: 1,1…1,6\n  шпоночные пазы/резкие уступы: 1,6…3,0\n  очень острые надрезы: до 4…5+ (реже в “нормальном” машиностроении)',
        'Fр': 'Преднатяжение ветвей ремня',
        'D1min': 'Минимальный диаметр ведущего шкива',
        'D1max': 'Максимальный диаметр ведущего шкива',
        'D2min': 'Минимальный диаметр ведомого шкива',
        'D2max': 'Максимальный диаметр ведомого шкива',
        'n1min': 'Минимальная частота вращения ведущего шкива',
        'n1max': 'Максимальная частота вращения ведущего шкива',
        'Fpmin': 'Минимальное преднатяжение ветвей ремня',
        'Fpmax': 'Максимальное преднатяжение ветвей ремня',
        'Tогранич': 'Максимальное натяжение ветви ремня',
        'Sнач': 'Начальная стоимость испытаний (можно взять после 1-го расчета)',
        'Fнач': 'Начальная требуемая внешняя радиальная сила (можно взять после 1-го расчета)',
        'wст': 'Вес начальной стоимости испытаний (для оптимизации)',
        'wF': 'Вес начальной требуемой внешней радиальной силы (для оптимизации)',
    }

    @staticmethod
    def _normalize_help_key(name: str) -> str:
        k = (name or "").strip().replace("=", "").strip()
        # Разные варианты апострофов/штрихов приводим к одному
        k = k.replace("’", "'").replace("′", "'").replace("`", "'")
        # Убираем пробелы внутри (например, если кто-то добавит)
        k = re.sub(r"\s+", "", k)
        return k

    def _resolve_help_text(self, name: str, cell_norm: str, item: InputCell) -> str:
        desc = (item.description or "").strip()
        if desc:
            return desc
        key = self._normalize_help_key(name)
        return self.HELP_TEXTS.get(key, "Пояснение не задано.")


    def __init__(self, schema: Schema) -> None:
        super().__init__()
        self.schema = schema
        self.engine = ComputeEngine(schema)
        self.optimizer = Optimizer(schema, self.engine)
        self.inputs_by_cell = schema_inputs_by_cell(schema)
        self.formulas_by_cell = schema_formulas_by_cell(schema)

        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
        self.assets_dir = base / "assets"

        self.sections_map = self._load_json(self.assets_dir / "sections_map.json", default={"sections": []})
        self.key_outputs_map = self._load_json(self.assets_dir / "key_outputs_map.json", default={})

        self.state = UiState()
        self.values = self.engine.default_values()
        self.recent_store = RecentProjects(self._settings_path())
        self.project_store = ProjectStorage(self._projects_dir())

        self.setWindowTitle(
            "Программа для расчета кинематики и системы нагружения в стенде для испытаний на прочность "
            "рабочих органов измельчителя-разбрасывателя соломы зерноуборочного комбайна"
        )

        self.resize(1400, 900)

        self._build_toolbar()
        self._build_menu()
        self._build_layout()
        self._apply_theme()
        self._update_ui_from_values()

    # --------------------- UI helpers ---------------------
    def _load_json(self, path: Path, default: Any) -> Any:
        try:
            if path.exists():
                return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default
        return default

    def _decimals_for_cell(self, cell: str) -> int:
        return 3 if normalize_cell(cell) in self.SPECIAL_3DP else 1

    def _format_value(self, value: float | None, decimals: int) -> str:
        if value is None:
            return "-"
        locale = QLocale("ru_RU")
        return locale.toString(float(value), "f", int(decimals))

    def _formula_image_path(self, filename: str) -> Path:
        return self.assets_dir / "formulas" / filename

    def _scale_formula_pixmap(self, pix: QPixmap) -> QPixmap:
        """Не увеличиваем изображения формул. Только уменьшаем при необходимости."""
        if pix.isNull():
            return pix

        # Ограничение по ширине (только downscale)
        max_w = 680
        scaled = pix
        if pix.width() > max_w:
            scaled = pix.scaledToWidth(max_w, Qt.SmoothTransformation)

        # Ограничение по высоте относительно 14 pt
        ref_font = QFont(self.font().family(), 14)
        fm = QFontMetrics(ref_font)
        max_h = int(fm.height() * 4.0)  # формулы с дробями обычно ~3-4 высоты строки
        if scaled.height() > max_h:
            scaled = scaled.scaledToHeight(max_h, Qt.SmoothTransformation)

        return scaled

    # --------------------- Build UI ---------------------
    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Главная")
        self.addToolBar(toolbar)

        def add_action(title: str, handler) -> QAction:
            action = QAction(title, self)
            action.triggered.connect(handler)
            toolbar.addAction(action)
            return action

        self.action_calculate = add_action("Рассчитать", self.on_calculate)
        self.action_optimize = add_action("Оптимизировать", self.on_optimize)
        self.action_save = add_action("Сохранить", self.on_save)
        self.action_save_as = add_action("Сохранить как", self.on_save_as)
        self.action_open = add_action("Открыть проект", self.on_open_project)
        self.action_import = add_action("Импорт из Excel", self.on_import_excel)
        self.action_export_xlsx = add_action("Экспорт в Excel", self.on_export_excel)
        self.action_export_pdf = add_action("Экспорт в PDF", self.on_export_pdf)
        self.action_reset = add_action("Сброс к значениям по умолчанию", self.on_reset_defaults)

        # Оптимизация всегда доступна
        self.action_optimize.setEnabled(True)

        self.theme_toggle = QCheckBox("Тёмная тема")
        self.theme_toggle.stateChanged.connect(self.on_toggle_theme)
        toolbar.addWidget(self.theme_toggle)

    def _build_menu(self) -> None:
        menu = self.menuBar().addMenu("Файл")
        open_action = QAction("Открыть проект", self)
        open_action.triggered.connect(self.on_open_project)
        menu.addAction(open_action)
        self.recent_menu = menu.addMenu("Недавние проекты")
        self._refresh_recent_menu()

    def _refresh_recent_menu(self) -> None:
        self.recent_menu.clear()
        recents = self.recent_store.load()
        if not recents:
            empty = QAction("Нет проектов", self)
            empty.setEnabled(False)
            self.recent_menu.addAction(empty)
            return
        for item in recents:
            action = QAction(item.path, self)
            action.triggered.connect(lambda checked=False, path=item.path: self._open_recent(path))
            self.recent_menu.addAction(action)

    def _open_recent(self, path: str) -> None:
        project = self.project_store.load(Path(path))
        self.state.project_path = Path(path)
        self.values.update({normalize_cell(k): v for k, v in project.inputs.items()})
        self._update_ui_from_values()
        self.on_calculate()
        self.recent_store.add(Path(path))
        self._refresh_recent_menu()

    def _build_layout(self) -> None:
        splitter = QSplitter(Qt.Horizontal)

        # Левая панель (ввод)
        inputs_widget = self._build_inputs_panel()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(inputs_widget)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        splitter.addWidget(scroll)
        splitter.addWidget(self._build_outputs_panel())
        splitter.setStretchFactor(1, 2)
        self.setCentralWidget(splitter)

    def _build_inputs_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        self.input_widgets: Dict[str, QLineEdit] = {}
        locale = QLocale("ru_RU")

        input_sections = [s for s in self.sections_map.get("sections", []) if s.get("inputs")]

        for section in input_sections:
            title = section.get("title", "Ввод")
            group = QGroupBox(title)
            form = QFormLayout(group)

            for cell in section.get("inputs", []):
                cell_norm = normalize_cell(cell)
                item = self.inputs_by_cell.get(cell_norm)
                if not item:
                    continue

                name = (item.name or "").strip().replace("=", "")
                if cell_norm.startswith("N") and cell_norm[1:].isdigit():
                    row = int(cell_norm[1:])
                    if 106 <= row <= 116:
                        name = f"k{row - 105}"
                if not name:
                    name = cell_norm

                
                help_text = self._resolve_help_text(name, cell_norm, item)
                unit_text = (item.unit or "").strip()
                if unit_text.startswith("="):
                    unit_text = ""

                label_widget = QWidget()
                label_layout = QHBoxLayout(label_widget)
                label_layout.setContentsMargins(0, 0, 0, 0)

                label = QLabel(name)
                label.setToolTip(cell_norm)
                label_layout.addWidget(label)

                # Кнопка подсказки: '?' в кружочке, без меню (без стрелки)
                help_btn = QPushButton()
                help_btn.setFlat(True)
                help_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_MessageBoxQuestion))
                help_btn.setToolTip(help_text)
                help_btn.setFixedSize(22, 22)
                help_btn.clicked.connect(
                    lambda _=False, t=help_text, ttl=name: QMessageBox.information(self, ttl or "Пояснение", t)
                )
                label_layout.addWidget(help_btn)
                label_layout.addStretch()

                field = QLineEdit()
                validator = QDoubleValidator()
                validator.setLocale(locale)
                field.setValidator(validator)

                row_widget = QWidget()
                row_layout = QHBoxLayout(row_widget)
                row_layout.setContentsMargins(0, 0, 0, 0)
                row_layout.addWidget(field)

                unit_label = QLabel(unit_text)
                unit_label.setMinimumWidth(52)
                unit_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                row_layout.addWidget(unit_label)

                form.addRow(label_widget, row_widget)
                self.input_widgets[cell_norm] = field

            layout.addWidget(group)

        layout.addStretch()
        return container

    def _build_outputs_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        # Ключевые результаты
        self.output_labels: Dict[str, QLabel] = {}
        output_group = QGroupBox("Ключевые результаты")
        output_layout = QFormLayout(output_group)

        for cell in KEY_OUTPUTS:
            meta = self.key_outputs_map.get(cell, {})
            label_text = meta.get("label", cell)
            unit_text = meta.get("unit", "")
            left = QLabel(f"{label_text} =")
            left.setToolTip(cell)

            value_label = QLabel("-")
            value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

            unit_label = QLabel(unit_text)
            unit_label.setMinimumWidth(52)
            unit_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.addWidget(value_label)
            row_layout.addWidget(unit_label)

            output_layout.addRow(left, row)
            self.output_labels[cell] = value_label

        layout.addWidget(output_group)

        # Вкладки
        self.tabs = QTabWidget()

        # Таблица k,a,b,s
        self.table = QTableWidget(11, 4)
        self.table.setHorizontalHeaderLabels(["k", "a", "b", "s"])

        self.tabs.addTab(self._build_calc_tab(), "Расчет")

        # Оптимизация
        self.iter_table = QTableWidget(0, 8)
        self.iter_table.setHorizontalHeaderLabels(
            ["Итерация", "D1, мм", "D2, мм", "n1, об/мин", "Fр, Н", "J", "TR", "TRmax"]
        )

        self.constraint_table = QTableWidget(0, 3)
        self.constraint_table.setHorizontalHeaderLabels(["Ограничение", "Нарушение", "Статус"])

        iter_container = QWidget()
        iter_layout = QVBoxLayout(iter_container)
        iter_layout.addWidget(QLabel("Статус ограничений"))
        iter_layout.addWidget(self.constraint_table)
        iter_layout.addWidget(QLabel("История итераций"))
        iter_layout.addWidget(self.iter_table)

        self.iter_plot = MatplotlibCanvas()
        self.iter_plot.set_qt_style(self.font(), QLocale("ru_RU"))
        iter_layout.addWidget(self.iter_plot)
        self.tabs.addTab(iter_container, "Оптимизация")

        # Профили цели
        self.profile_plot = MatplotlibCanvas()
        self.profile_plot.set_qt_style(self.font(), QLocale("ru_RU"))
        profile_container = QWidget()
        profile_layout = QVBoxLayout(profile_container)
        profile_layout.addWidget(self.profile_plot)
        self.tabs.addTab(profile_container, "Профили цели")

        layout.addWidget(self.tabs)
        return container

    def _build_calc_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(8, 8, 8, 8)

        self.section_output_labels: Dict[str, QLabel] = {}

        for section in self.sections_map.get("sections", []):
            title = section.get("title", "")
            if title == "Исходные данные":
                continue

            group = QGroupBox(title)
            g_layout = QVBoxLayout(group)

            # Формулы
            for img in section.get("images", []) or []:
                img_path = self._formula_image_path(img)
                if not img_path.exists():
                    continue
                pix = QPixmap(str(img_path))
                img_lbl = QLabel()
                img_lbl.setAlignment(Qt.AlignCenter)
                img_lbl.setPixmap(self._scale_formula_pixmap(pix))
                g_layout.addWidget(img_lbl)

            # Раздел 12: пояснение + таблица
            if title.startswith("12."):
                g_layout.addWidget(
                    QLabel(
                        "k — доля пролёта от ведущего шкива; a — расстояние от ведущего шкива до ролика, мм; "
                        "b — расстояние от ролика до ведомого шкива, мм."
                    )
                )
                g_layout.addWidget(QLabel("Таблица k, a, b, s"))
                g_layout.addWidget(self.table)
                layout.addWidget(group)
                continue

            outputs = section.get("outputs", []) or []
            if outputs:
                form = QFormLayout()
                for out in outputs:
                    cell = normalize_cell(out.get("cell", ""))
                    label_text = (out.get("label") or "").strip() or cell
                    unit_text = (out.get("unit") or "").strip()
                    if unit_text.startswith("="):
                        unit_text = ""

                    left = QLabel(label_text)
                    value_label = QLabel("-")
                    value_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

                    unit_label = QLabel(unit_text)
                    unit_label.setMinimumWidth(52)
                    unit_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

                    row = QWidget()
                    row_layout = QHBoxLayout(row)
                    row_layout.setContentsMargins(0, 0, 0, 0)
                    row_layout.addWidget(value_label)
                    row_layout.addWidget(unit_label)
                    form.addRow(left, row)

                    if cell:
                        self.section_output_labels[cell] = value_label

                g_layout.addLayout(form)

            layout.addWidget(group)

        layout.addStretch()
        scroll.setWidget(container)

        wrapper = QWidget()
        w_layout = QVBoxLayout(wrapper)
        w_layout.setContentsMargins(0, 0, 0, 0)
        w_layout.addWidget(scroll)
        return wrapper

    # --------------------- Data sync ---------------------
    def _update_ui_from_values(self) -> None:
    locale = QLocale("ru_RU")
    for cell, field in self.input_widgets.items():
        val = self.values.get(cell, 0.0)
        # Показываем пусто, если 0 (по умолчанию все параметры = 0)
        if val is None or abs(float(val)) < 1e-12:
            field.clear()
        else:
            field.setText(locale.toString(float(val)))

    def _collect_inputs(self) -> Dict[str, float]:
        locale = QLocale("ru_RU")
        values = dict(self.values)
        for cell, field in self.input_widgets.items():
            text = field.text().replace(" ", "").strip()
            if text == "":
                values[cell] = 0.0
                continue

        value, ok = locale.toDouble(text)
        if not ok:
            raise ValueError(f"Некорректное значение в {cell}")
        values[cell] = value
        return values

    # --------------------- Actions ---------------------
    def on_calculate(self) -> None:
        try:
            self.values = self._collect_inputs()
            result = self.engine.compute(self.values)
            self._update_outputs(result)
            self._autosave(result)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Ошибка", str(exc))

    def on_optimize(self) -> None:
        try:
            self.values = self._collect_inputs()
            result = self.optimizer.optimize(self.values)
            self.values = result.values
            compute_result = self.engine.compute(self.values)
            self._update_outputs(compute_result)
            self._update_constraints(result.constraints)
            self._update_history(result.history)
            self._update_profile_plot(self.values)
            self._autosave(compute_result, optimization=result)
            if not result.success:
                QMessageBox.warning(self, "Оптимизация", f"Не найдено решение: {result.message}")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.warning(self, "Ошибка", str(exc))

    def on_save(self) -> None:
        if not self.state.project_path:
            self.on_save_as()
            return
        self._save_project(self.state.project_path)

    def on_save_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить проект", str(self._projects_dir()), "Project (*.kproj)")
        if not path:
            return
        self.state.project_path = Path(path)
        self._save_project(self.state.project_path)

    def on_open_project(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Открыть проект", str(self._projects_dir()), "Project (*.kproj)")
        if not path:
            return
        project = self.project_store.load(Path(path))
        self.state.project_path = Path(path)
        self.values.update({normalize_cell(k): v for k, v in project.inputs.items()})
        self._update_ui_from_values()
        self.on_calculate()
        self.recent_store.add(Path(path))
        self._refresh_recent_menu()

    def on_import_excel(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Импорт из Excel", str(Path.home()), "Excel (*.xlsx)")
        if not path:
            return
        input_cells = list(self.input_widgets.keys())
        result = import_from_excel(path, input_cells, self.values)
        if result.errors:
            QMessageBox.warning(self, "Импорт", f"Некорректные значения: {', '.join(result.errors)}")
        if result.updated:
            summary = "\n".join([f"{cell}: {old} → {new}" for cell, (old, new) in result.updated.items()])
            QMessageBox.information(self, "Импорт", f"Обновлены значения:\n{summary}")
        self.values = result.values
        self._update_ui_from_values()

    def on_export_excel(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Экспорт в Excel", str(self._projects_dir()), "Excel (*.xlsx)")
        if not path:
            return
        compute_result = self.engine.compute(self.values)
        constraints = self.optimizer._constraint_violations(self.values)[1]
        export_report(
            path,
            self.schema.inputs,
            self.values,
            compute_result.key_outputs,
            {cell: f.formula for cell, f in self.formulas_by_cell.items()},
            constraints,
        )

    def on_export_pdf(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Экспорт в PDF", str(self._projects_dir()), "PDF (*.pdf)")
        if not path:
            return
        compute_result = self.engine.compute(self.values)
        constraints = self.optimizer._constraint_violations(self.values)[1]
        export_pdf(
            path,
            self.state.project_path.stem if self.state.project_path else "Без имени",
            VERSION,
            self.schema.inputs,
            self.values,
            compute_result.key_outputs,
            constraints,
            self.state.notes,
        )

    def on_reset_defaults(self) -> None:
        self.values = self.engine.default_values()
        for cell in self.inputs_by_cell.keys():
            self.values[cell] = 0.0
        self._update_ui_from_values()

        # Сбросим отображение результатов, но расчёт не запускаем
        for lbl in self.output_labels.values():
            lbl.setText("-")
        for lbl in getattr(self, "section_output_labels", {}).values():
            lbl.setText("-")
        self.table.clearContents()
        self.constraint_table.setRowCount(0)
        self.iter_table.setRowCount(0)
        self.iter_plot.ax.clear()
        self.iter_plot.draw()
        self.profile_plot.ax.clear()
        self.profile_plot.draw()

    def on_toggle_theme(self) -> None:
        self.state.theme = "dark" if self.theme_toggle.isChecked() else "light"
        self._apply_theme()

    # --------------------- Update views ---------------------
    def _update_outputs(self, result) -> None:
        # ключевые результаты
        for cell, label in self.output_labels.items():
            value = result.key_outputs.get(cell)
            label.setText(self._format_value(value, self._decimals_for_cell(cell)))

        # таблица k,a,b,s (1 знак)
        for row_idx, row in enumerate(result.table_rows):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(self._format_value(value, 1) if value is not None else "-")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)

        # значения по разделам
        for cell, label in getattr(self, "section_output_labels", {}).items():
            value = result.values.get(cell)
            label.setText(self._format_value(value, self._decimals_for_cell(cell)))

    def _update_constraints(self, constraints) -> None:
        locale = QLocale("ru_RU")
        self.constraint_table.setRowCount(0)

        def pretty_expr(expr: str) -> str:
            key = normalize_cell(str(expr))

            # 1) входные параметры
            item = self.inputs_by_cell.get(key)
            if item and (item.name or "").strip():
                name = (item.name or "").replace("=", "").strip()
                unit = (item.unit or "").strip()
                return f"{name}, {unit}".rstrip(", ") if unit else name

            # 2) ключевые результаты
            meta = self.key_outputs_map.get(key, {})
            if meta.get("label"):
                unit = (meta.get("unit") or "").strip()
                return f"{meta['label']}, {unit}".rstrip(", ") if unit else meta["label"]

            # 3) выходы разделов
            for sec in self.sections_map.get("sections", []):
                for out in (sec.get("outputs") or []):
                    if normalize_cell(out.get("cell", "")) == key:
                        lbl = (out.get("label") or key).replace("=", "").strip()
                        unit = (out.get("unit") or "").replace("=", "").strip()
                        return f"{lbl}, {unit}".rstrip(", ") if unit else lbl

            return key

        for idx, item in enumerate(constraints):
            self.constraint_table.insertRow(idx)
            raw = item.name
            parsed = re.match(r"C(\d+):\s*(.+?)\s+(le|ge|eq)\s+(.+)$", raw)
            if parsed:
                lhs = pretty_expr(parsed.group(2))
                rhs = pretty_expr(parsed.group(4))
                op = {"le": "≤", "ge": "≥", "eq": "="}.get(parsed.group(3), parsed.group(3))
                pretty = f"{lhs} {op} {rhs}"
            else:
                pretty = raw

            self.constraint_table.setItem(idx, 0, QTableWidgetItem(pretty))
            self.constraint_table.setItem(
                idx,
                1,
                QTableWidgetItem(locale.toString(item.violation * 100.0, "f", 2) + " %"),
            )
            self.constraint_table.setItem(idx, 2, QTableWidgetItem(item.status))

    def _update_history(self, history) -> None:
        self.iter_table.setRowCount(0)
        xs: List[float] = []
        ys: List[float] = []

        for idx, record in enumerate(history.records):
            self.iter_table.insertRow(idx)
            self.iter_table.setItem(idx, 0, QTableWidgetItem(str(record.iteration)))

            self.iter_table.setItem(idx, 1, QTableWidgetItem(self._format_value(record.variables["C7"], 1)))
            self.iter_table.setItem(idx, 2, QTableWidgetItem(self._format_value(record.variables["J7"], 1)))
            self.iter_table.setItem(idx, 3, QTableWidgetItem(self._format_value(record.variables["C9"], 1)))
            self.iter_table.setItem(idx, 4, QTableWidgetItem(self._format_value(record.variables["J94"], 1)))

            # J204 / J76 / J77
            self.iter_table.setItem(idx, 5, QTableWidgetItem(self._format_value(record.objective, 3)))
            self.iter_table.setItem(idx, 6, QTableWidgetItem(self._format_value(record.outputs.get("J76"), 3)))
            self.iter_table.setItem(idx, 7, QTableWidgetItem(self._format_value(record.outputs.get("J77"), 3)))

            xs.append(float(record.iteration))
            ys.append(float(record.objective))

        if xs:
            self.iter_plot.plot(xs, ys, "J по итерациям", "Итерация", "J", y_decimals=3)

    def _update_profile_plot(self, values: Dict[str, float], percent: float = 10.0, points: int = 41) -> None:
        final_values = values
        variables = ["C7", "J7", "C9", "J94"]  # D1, D2, n1, Fр
        sweep = [((i - (points // 2)) / (points // 2)) * percent for i in range(points)]
        series: Dict[str, List[float]] = {}
        for var in variables:
            ys: List[float] = []
            for delta in sweep:
                current = dict(final_values)
                current[var] = final_values[var] * (1 + delta / 100.0)
                ys.append(float(self.engine.compute(current).key_outputs["J204"]))
            label = {"C7": "D1, мм", "J7": "D2, мм", "C9": "n1, об/мин", "J94": "Fр, Н"}.get(var, var)
            series[label] = ys
        self.profile_plot.plot_multi(series, sweep, x_decimals=1, y_decimals=3)

    # --------------------- Theme & persistence ---------------------
    def _apply_theme(self) -> None:
        app = QApplication.instance()
        if not app:
            return
        app.setStyleSheet(DARK_STYLESHEET if self.state.theme == "dark" else LIGHT_STYLESHEET)
        self.theme_toggle.blockSignals(True)
        self.theme_toggle.setChecked(self.state.theme == "dark")
        self.theme_toggle.blockSignals(False)

    def _save_project(self, path: Path, optimization=None) -> None:
        compute_result = self.engine.compute(self.values)
        project = ProjectData(
            name=path.stem,
            created_at=datetime.now().isoformat(),
            modified_at="",
            version=VERSION,
            inputs=self.values,
            ui_options={"theme": self.state.theme},
            results={"key_outputs": compute_result.key_outputs, "table": compute_result.table_rows},
            optimization=optimization.history.to_dict() if optimization else {},
            notes=self.state.notes,
        )
        self.project_store.save(path, project)
        self.recent_store.add(path)
        self._refresh_recent_menu()

    def _autosave(self, result, optimization=None) -> None:
        autosave_path = self._projects_dir() / "autosave.kproj"
        self._save_project(autosave_path, optimization=optimization)

    def _settings_path(self) -> Path:
        return self._documents_dir() / "settings.json"

    def _projects_dir(self) -> Path:
        path = self._documents_dir() / "Projects"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _documents_dir(self) -> Path:
        return Path.home() / "Documents" / "KinematicsCalc"
