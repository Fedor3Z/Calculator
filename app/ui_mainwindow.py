from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
from typing import Dict, List, Optional, Any

from PySide6.QtCore import Qt, QLocale
from PySide6.QtGui import QAction, QDoubleValidator, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QToolButton,
    QMenu,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QScrollArea,
)

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from app import VERSION
from core.compute import ComputeEngine, KEY_OUTPUTS
from core.dependency_graph import build_dependency_graph, dependency_chain
from core.model import InputCell, Schema, normalize_cell, schema_formulas_by_cell, schema_inputs_by_cell
from export.export_pdf import export_pdf
from export.export_xlsx import export_report
from imports.import_excel import import_from_excel
from solver.optimizer import Optimizer
from storage.projects import ProjectData, ProjectStorage, default_project_name
from storage.recents import RecentProjects


@dataclass
class UiState:
    project_path: Optional[Path] = None
    notes: str = ""
    theme: str = "light"
    mode: str = "user"


LIGHT_STYLESHEET = """
QWidget { }
QLineEdit, QTableWidget, QComboBox {
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
QLineEdit, QTableWidget, QComboBox {
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
        super().__init__(self.figure)

    def plot(self, x: List[float], y: List[float], title: str, xlabel: str, ylabel: str) -> None:
        self.ax.clear()
        self.ax.plot(x, y, marker="o")
        self.ax.set_title(title)
        self.ax.set_xlabel(xlabel)
        self.ax.set_ylabel(ylabel)
        self.figure.tight_layout()
        self.draw()

    def plot_multi(self, series: Dict[str, List[float]], x: List[float]) -> None:
        self.ax.clear()
        for label, values in series.items():
            self.ax.plot(x, values, label=label)
        self.ax.legend()
        self.ax.set_title("Профили цели")
        self.ax.set_xlabel("Отклонение, %")
        self.ax.set_ylabel("J204")
        self.figure.tight_layout()
        self.draw()


class MainWindow(QMainWindow):
    def __init__(self, schema: Schema) -> None:
        super().__init__()
        self.schema = schema
        self.engine = ComputeEngine(schema)
        self.optimizer = Optimizer(schema, self.engine)
        self.inputs_by_cell = schema_inputs_by_cell(schema)
        self.formulas_by_cell = schema_formulas_by_cell(schema)

        self.assets_dir = Path(__file__).resolve().parent.parent / "assets"
        self.sections_map = self._load_sections_map()
        self.key_outputs_map = self._load_key_outputs_map()

        self.state = UiState()
        self.values = self.engine.default_values()
        self.recent_store = RecentProjects(self._settings_path())
        self.project_store = ProjectStorage(self._projects_dir())

        self.setWindowTitle("Расчет кинематики и системы нагружения")
        self.resize(1400, 900)

        self._build_toolbar()
        self._build_menu()
        self._build_layout()
        self._apply_theme()
        self._apply_mode()
        self._update_ui_from_values()

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

        self.theme_toggle = QCheckBox("Тёмная тема")
        self.theme_toggle.stateChanged.connect(self.on_toggle_theme)
        toolbar.addWidget(self.theme_toggle)

        self.mode_toggle = QComboBox()
        self.mode_toggle.addItems(["Пользовательский", "Инженерный"])
        self.mode_toggle.currentIndexChanged.connect(self.on_toggle_mode)
        toolbar.addWidget(self.mode_toggle)

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

    def _load_sections_map(self) -> Dict[str, Any]:
        path = self.assets_dir / "sections_map.json"
        if not path.exists():
            return {"sections": []}
        return json.loads(path.read_text(encoding="utf-8"))

    def _load_key_outputs_map(self) -> Dict[str, Any]:
        path = self.assets_dir / "key_outputs_map.json"
        if not path.exists():
            return {}
        return json.loads(path.read_text(encoding="utf-8"))

    def _formula_image_path(self, filename: str) -> Path:
        return self.assets_dir / "formulas" / filename

    def _build_layout(self) -> None:
        splitter = QSplitter(Qt.Horizontal)

        # Левая панель (ввод) -> в ScrollArea
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
        self.input_groups_user: List[QGroupBox] = []
        self.input_groups_engineering: List[QGroupBox] = []

        locale = QLocale("ru_RU")
        essential_titles = {
            "Исходные данные",
            "10. Преднатяжение Fр, Н",
            "14. Измеренный прогиб δизм, мм",
        }

        input_sections = [
            s for s in self.sections_map.get("sections", [])
            if s.get("inputs")
        ]

        for section in input_sections:
            title = section.get("title", "Ввод")
            group = QGroupBox(title)
            form = QFormLayout(group)

            for cell in section.get("inputs", []):
                cell_norm = normalize_cell(cell)
                item = self.inputs_by_cell.get(cell_norm)
                if not item:
                    continue

                # Более читабельные подписи: без ссылок на ячейки
                name = (item.name or "").strip()
                if cell_norm.startswith("N") and cell_norm[1:].isdigit():
                    row = int(cell_norm[1:])
                    if 106 <= row <= 116:
                        name = f"k{row - 105}"
                if not name:
                    name = cell_norm
                if not name.endswith("="):
                    name = f"{name} ="

                help_text = (item.description or "").strip()
                unit_text = (item.unit or "").strip()
                if unit_text.startswith("="):
                    unit_text = ""
                if not help_text:
                    help_text = "Пояснение будет добавлено."

                label_widget = QWidget()
                label_layout = QHBoxLayout(label_widget)
                label_layout.setContentsMargins(0, 0, 0, 0)
                label = QLabel(name)
                label.setToolTip(cell_norm)
                label_layout.addWidget(label)
                help_btn = QToolButton()
                help_btn.setText("?")
                help_btn.setAutoRaise(True)
                menu = QMenu(help_btn)
                act = QAction(help_text, menu)
                act.setEnabled(False)
                menu.addAction(act)
                help_btn.setMenu(menu)
                help_btn.setPopupMode(QToolButton.InstantPopup)
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

            if title in essential_titles:
                self.input_groups_user.append(group)
            else:
                self.input_groups_engineering.append(group)

        layout.addStretch()
        return container

    def _build_outputs_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        # --- Ключевые результаты (показываем обозначения, не адреса ячеек)
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

        # --- Вкладки
        self.tabs = QTabWidget()

        # Таблица k,a,b,s (будет вставлена в раздел 12 внутри вкладки "Расчет")
        self.table = QTableWidget(11, 4)
        self.table.setHorizontalHeaderLabels(["k", "a", "b", "s"])

        self.calc_tab_index = 0
        self.tabs.addTab(self._build_calc_tab(), "Расчёт")

        # Оптимизация
        self.iter_table = QTableWidget(0, 8)
        self.iter_table.setHorizontalHeaderLabels([
            "Итерация",
            "C7",
            "J7",
            "C9",
            "J94",
            "J204",
            "J76",
            "J77",
        ])
        self.constraint_table = QTableWidget(0, 3)
        self.constraint_table.setHorizontalHeaderLabels(["Ограничение", "Нарушение", "Статус"])

        iter_container = QWidget()
        iter_layout = QVBoxLayout(iter_container)
        iter_layout.addWidget(QLabel("Статус ограничений"))
        iter_layout.addWidget(self.constraint_table)
        iter_layout.addWidget(QLabel("История итераций"))
        iter_layout.addWidget(self.iter_table)
        self.iter_plot = MatplotlibCanvas()
        iter_layout.addWidget(self.iter_plot)
        self.tab_index_optimization = self.tabs.addTab(iter_container, "Оптимизация")

        # Профили цели
        self.profile_plot = MatplotlibCanvas()
        profile_container = QWidget()
        profile_layout = QVBoxLayout(profile_container)
        profile_layout.addWidget(self.profile_plot)
        self.tab_index_profiles = self.tabs.addTab(profile_container, "Профили цели")

        # Инженерный
        self.all_cells_table = QTableWidget(0, 2)
        self.formula_table = QTableWidget(0, 2)
        self.dependency_table = QTableWidget(0, 1)
        self.dependency_selector = QComboBox()
        self.dependency_selector.addItems(sorted(self.formulas_by_cell.keys()))
        self.dependency_selector.currentTextChanged.connect(self._update_dependency_chain)

        engineering_container = QWidget()
        engineering_layout = QVBoxLayout(engineering_container)
        engineering_layout.addWidget(QLabel("Все ячейки"))
        engineering_layout.addWidget(self.all_cells_table)
        engineering_layout.addWidget(QLabel("Формулы"))
        engineering_layout.addWidget(self.formula_table)
        engineering_layout.addWidget(QLabel("Зависимости"))
        engineering_layout.addWidget(self.dependency_selector)
        engineering_layout.addWidget(self.dependency_table)
        self.tab_index_engineering = self.tabs.addTab(engineering_container, "Инженерный")

        layout.addWidget(self.tabs)
        return container

    def _build_calc_tab(self) -> QWidget:
        # Содержимое вкладки "Расчёт" (разделы 1..19.2)
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
                continue  # исходные данные остаются в левой панели

            group = QGroupBox(title)
            g_layout = QVBoxLayout(group)

            # картинки формул (как в Excel)
            for img in section.get("images", []) or []:
                img_path = self._formula_image_path(img)
                if not img_path.exists():
                    continue
                pix = QPixmap(str(img_path))
                img_lbl = QLabel()
                img_lbl.setAlignment(Qt.AlignCenter)
                img_lbl.setPixmap(pix.scaledToWidth(680, Qt.SmoothTransformation))
                g_layout.addWidget(img_lbl)

            # раздел 12 — показываем таблицу k,a,b,s (вместо длинного списка ячеек)
            if title.startswith("12."):
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

    def _update_ui_from_values(self) -> None:
        locale = QLocale("ru_RU")
        for cell, field in self.input_widgets.items():
            field.setText(locale.toString(self.values.get(cell, 0.0)))

    def _collect_inputs(self) -> Dict[str, float]:
        locale = QLocale("ru_RU")
        values = dict(self.values)
        for cell, field in self.input_widgets.items():
            text = field.text().replace(" ", "")
            value, ok = locale.toDouble(text)
            if not ok:
                raise ValueError(f"Некорректное значение в {cell}")
            values[cell] = value
        return values

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
        self._update_ui_from_values()
        self.on_calculate()

    def _apply_theme(self) -> None:
        app = QApplication.instance()
        if not app:
            return
        app.setStyleSheet(DARK_STYLESHEET if self.state.theme == "dark" else LIGHT_STYLESHEET)
        # синхронизация переключателя
        self.theme_toggle.blockSignals(True)
        self.theme_toggle.setChecked(self.state.theme == "dark")
        self.theme_toggle.blockSignals(False)

    def _apply_mode(self) -> None:
        is_eng = self.state.mode == "engineering"

        # вкладки
        if hasattr(self, "tabs"):
            self.tabs.setTabVisible(self.tab_index_optimization, is_eng)
            self.tabs.setTabVisible(self.tab_index_profiles, is_eng)
            self.tabs.setTabVisible(self.tab_index_engineering, is_eng)

        # панель ввода
        for g in getattr(self, "input_groups_engineering", []):
            g.setVisible(is_eng)
        for g in getattr(self, "input_groups_user", []):
            g.setVisible(True)

        # действия/кнопки
        self.action_optimize.setEnabled(is_eng)

        # синхронизация комбобокса
        self.mode_toggle.blockSignals(True)
        self.mode_toggle.setCurrentIndex(1 if is_eng else 0)
        self.mode_toggle.blockSignals(False)

    def on_toggle_theme(self) -> None:
        self.state.theme = "dark" if self.theme_toggle.isChecked() else "light"
        self._apply_theme()

    def on_toggle_mode(self) -> None:
        self.state.mode = "engineering" if self.mode_toggle.currentIndex() == 1 else "user"
        self._apply_mode()

    def _update_outputs(self, result) -> None:
        locale = QLocale("ru_RU")
        for cell, label in self.output_labels.items():
            value = result.key_outputs.get(cell)
            label.setText(locale.toString(value, 'f', 4) if value is not None else "-")
        for row_idx, row in enumerate(result.table_rows):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(locale.toString(value, 'f', 4) if value is not None else "-")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)

        # значения по разделам (1..19.2)
        for cell, label in getattr(self, "section_output_labels", {}).items():
            value = result.values.get(cell)
            label.setText(locale.toString(value, 'f', 4) if value is not None else "-")
        self._update_all_cells(result.values)

    def _update_constraints(self, constraints) -> None:
        self.constraint_table.setRowCount(0)
        for idx, item in enumerate(constraints):
            self.constraint_table.insertRow(idx)
            self.constraint_table.setItem(idx, 0, QTableWidgetItem(item.name))
            self.constraint_table.setItem(idx, 1, QTableWidgetItem(f"{item.violation:.2%}"))
            self.constraint_table.setItem(idx, 2, QTableWidgetItem(item.status))

    def _update_history(self, history) -> None:
        self.iter_table.setRowCount(0)
        xs = []
        ys = []
        for idx, record in enumerate(history.records):
            self.iter_table.insertRow(idx)
            self.iter_table.setItem(idx, 0, QTableWidgetItem(str(record.iteration)))
            self.iter_table.setItem(idx, 1, QTableWidgetItem(f"{record.variables['C7']:.4f}"))
            self.iter_table.setItem(idx, 2, QTableWidgetItem(f"{record.variables['J7']:.4f}"))
            self.iter_table.setItem(idx, 3, QTableWidgetItem(f"{record.variables['C9']:.4f}"))
            self.iter_table.setItem(idx, 4, QTableWidgetItem(f"{record.variables['J94']:.4f}"))
            self.iter_table.setItem(idx, 5, QTableWidgetItem(f"{record.objective:.4f}"))
            self.iter_table.setItem(idx, 6, QTableWidgetItem(f"{record.outputs['J76']:.4f}"))
            self.iter_table.setItem(idx, 7, QTableWidgetItem(f"{record.outputs['J77']:.4f}"))
            xs.append(record.iteration)
            ys.append(record.objective)
        if xs:
            self.iter_plot.plot(xs, ys, "J204 по итерациям", "Итерация", "J204")

    def _update_profile_plot(self, values: Dict[str, float], percent: float = 10.0, points: int = 41) -> None:
        final_values = values
        variables = ["C7", "J7", "C9", "J94"]
        sweep = [((i - (points // 2)) / (points // 2)) * percent for i in range(points)]
        series: Dict[str, List[float]] = {}
        for var in variables:
            ys = []
            for delta in sweep:
                current = dict(final_values)
                current[var] = final_values[var] * (1 + delta / 100.0)
                ys.append(self.engine.compute(current).key_outputs["J204"])
            series[var] = ys
        self.profile_plot.plot_multi(series, sweep)

    def _update_all_cells(self, values: Dict[str, float]) -> None:
        self.all_cells_table.setRowCount(0)
        for idx, (cell, value) in enumerate(sorted(values.items())):
            self.all_cells_table.insertRow(idx)
            self.all_cells_table.setItem(idx, 0, QTableWidgetItem(cell))
            self.all_cells_table.setItem(idx, 1, QTableWidgetItem(f"{value:.4f}"))
        self.formula_table.setRowCount(0)
        for idx, (cell, formula) in enumerate(sorted(self.formulas_by_cell.items())):
            self.formula_table.insertRow(idx)
            self.formula_table.setItem(idx, 0, QTableWidgetItem(cell))
            self.formula_table.setItem(idx, 1, QTableWidgetItem(formula.formula))

    def _update_dependency_chain(self, cell: str) -> None:
        graph = build_dependency_graph({c: f.formula for c, f in self.formulas_by_cell.items()})
        deps = dependency_chain(graph, cell)
        self.dependency_table.setRowCount(0)
        for idx, dep in enumerate(deps):
            self.dependency_table.insertRow(idx)
            self.dependency_table.setItem(idx, 0, QTableWidgetItem(dep))

    def _save_project(self, path: Path, optimization=None) -> None:
        compute_result = self.engine.compute(self.values)
        project = ProjectData(
            name=path.stem,
            created_at=datetime.now().isoformat(),
            modified_at="",
            version=VERSION,
            inputs=self.values,
            ui_options={"theme": self.state.theme, "mode": self.state.mode},
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
