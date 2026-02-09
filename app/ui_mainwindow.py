from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from PySide6.QtCore import Qt, QLocale
from PySide6.QtGui import QAction, QDoubleValidator
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
    QVBoxLayout,
    QWidget,
    QFileDialog,
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
        self.state = UiState()
        self.values = self.engine.default_values()
        self.recent_store = RecentProjects(self._settings_path())
        self.project_store = ProjectStorage(self._projects_dir())

        self.setWindowTitle("Расчет кинематики и системы нагружения")
        self.resize(1400, 900)

        self._build_toolbar()
        self._build_menu()
        self._build_layout()
        self.tabs.setTabVisible(2, False)
        self._update_ui_from_values()

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Главная")
        self.addToolBar(toolbar)

        def add_action(title: str, handler) -> QAction:
            action = QAction(title, self)
            action.triggered.connect(handler)
            toolbar.addAction(action)
            return action

        add_action("Рассчитать", self.on_calculate)
        add_action("Оптимизировать", self.on_optimize)
        add_action("Сохранить", self.on_save)
        add_action("Сохранить как", self.on_save_as)
        add_action("Открыть проект", self.on_open_project)
        add_action("Импорт из Excel", self.on_import_excel)
        add_action("Экспорт в Excel", self.on_export_excel)
        add_action("Экспорт в PDF", self.on_export_pdf)
        add_action("Сброс к значениям по умолчанию", self.on_reset_defaults)

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

    def _build_layout(self) -> None:
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_inputs_panel())
        splitter.addWidget(self._build_outputs_panel())
        splitter.setStretchFactor(1, 2)
        self.setCentralWidget(splitter)

    def _build_inputs_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        self.input_widgets: Dict[str, QLineEdit] = {}
        locale = QLocale("ru_RU")

        sections = {
            "Геометрия": [],
            "Режим": [],
            "Нагрузка": [],
            "Коэффициенты": [],
            "Ограничения": [],
            "Экономика": [],
        }
        for idx, input_cell in enumerate(self.schema.inputs):
            if input_cell.cell.startswith(("E", "J")):
                sections["Ограничения"].append(input_cell)
            elif idx % 5 == 0:
                sections["Геометрия"].append(input_cell)
            elif idx % 5 == 1:
                sections["Режим"].append(input_cell)
            elif idx % 5 == 2:
                sections["Нагрузка"].append(input_cell)
            elif idx % 5 == 3:
                sections["Коэффициенты"].append(input_cell)
            else:
                sections["Экономика"].append(input_cell)

        for title, items in sections.items():
            group = QGroupBox(title)
            form = QFormLayout(group)
            for item in items:
                field = QLineEdit()
                validator = QDoubleValidator()
                validator.setLocale(locale)
                field.setValidator(validator)
                field.setToolTip(item.description or item.unit)
                unit_label = QLabel(item.unit)
                row_widget = QWidget()
                row_layout = QHBoxLayout(row_widget)
                row_layout.setContentsMargins(0, 0, 0, 0)
                row_layout.addWidget(field)
                row_layout.addWidget(unit_label)
                form.addRow(QLabel(f"{item.name} ({item.cell})"), row_widget)
                self.input_widgets[normalize_cell(item.cell)] = field
            layout.addWidget(group)
        layout.addStretch()
        return container

    def _build_outputs_panel(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)

        self.output_labels: Dict[str, QLabel] = {}
        output_group = QGroupBox("Ключевые результаты")
        output_layout = QFormLayout(output_group)
        for cell in KEY_OUTPUTS:
            label = QLabel("-")
            label.setAlignment(Qt.AlignRight)
            output_layout.addRow(QLabel(cell), label)
            self.output_labels[cell] = label
        layout.addWidget(output_group)

        self.table = QTableWidget(11, 4)
        self.table.setHorizontalHeaderLabels(["k", "a", "b", "s"])
        layout.addWidget(QLabel("Таблица k, a, b, s"))
        layout.addWidget(self.table)

        self.constraint_table = QTableWidget(0, 3)
        self.constraint_table.setHorizontalHeaderLabels(["Ограничение", "Нарушение", "Статус"])
        layout.addWidget(QLabel("Статус ограничений"))
        layout.addWidget(self.constraint_table)

        self.tabs = QTabWidget()
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
        iter_container = QWidget()
        iter_layout = QVBoxLayout(iter_container)
        iter_layout.addWidget(self.iter_table)
        self.iter_plot = MatplotlibCanvas()
        iter_layout.addWidget(self.iter_plot)
        self.tabs.addTab(iter_container, "Оптимизация")

        self.profile_plot = MatplotlibCanvas()
        profile_container = QWidget()
        profile_layout = QVBoxLayout(profile_container)
        profile_layout.addWidget(self.profile_plot)
        self.tabs.addTab(profile_container, "Профили цели")

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
        self.tabs.addTab(engineering_container, "Инженерный")

        layout.addWidget(self.tabs)
        return container

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

    def on_toggle_theme(self) -> None:
        self.state.theme = "dark" if self.theme_toggle.isChecked() else "light"
        if self.state.theme == "dark":
            self.setStyleSheet("QWidget { background-color: #2b2b2b; color: #f0f0f0; }")
        else:
            self.setStyleSheet("")

    def on_toggle_mode(self) -> None:
        self.state.mode = "engineering" if self.mode_toggle.currentIndex() == 1 else "user"
        self.tabs.setTabVisible(2, self.state.mode == "engineering")

    def _update_outputs(self, result) -> None:
        for cell, label in self.output_labels.items():
            value = result.key_outputs.get(cell)
            label.setText(f"{value:.4f}" if value is not None else "-")
        for row_idx, row in enumerate(result.table_rows):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(f"{value:.4f}" if value is not None else "-")
                item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row_idx, col_idx, item)
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
