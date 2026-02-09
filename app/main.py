from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication

from app.ui_mainwindow import MainWindow
from core.model import SchemaLoader


def main() -> int:
    app = QApplication(sys.argv)
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    schema_path = base / "assets" / "kinematics_calc_extracted.json"
    schema = SchemaLoader(schema_path).load()
    window = MainWindow(schema)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
