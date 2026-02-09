from __future__ import annotations

import json
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional


@dataclass
class ProjectData:
    name: str
    created_at: str
    modified_at: str
    version: str
    inputs: Dict[str, float]
    ui_options: Dict[str, object]
    results: Dict[str, object]
    optimization: Dict[str, object]
    notes: str


class ProjectStorage:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir
        self.base_dir.mkdir(parents=True, exist_ok=True)

    def save(self, path: Path, data: ProjectData) -> None:
        data.modified_at = datetime.now().isoformat()
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("project.json", json.dumps({
                "name": data.name,
                "created_at": data.created_at,
                "modified_at": data.modified_at,
                "version": data.version,
                "inputs": data.inputs,
                "ui_options": data.ui_options,
            }, ensure_ascii=False, indent=2))
            zf.writestr("results.json", json.dumps(data.results, ensure_ascii=False, indent=2))
            zf.writestr("opt_history.json", json.dumps(data.optimization, ensure_ascii=False, indent=2))
            zf.writestr("notes.txt", data.notes)

    def load(self, path: Path) -> ProjectData:
        with zipfile.ZipFile(path, "r") as zf:
            project = json.loads(zf.read("project.json").decode("utf-8"))
            results = json.loads(zf.read("results.json").decode("utf-8"))
            optimization = json.loads(zf.read("opt_history.json").decode("utf-8"))
            notes = zf.read("notes.txt").decode("utf-8")
        return ProjectData(
            name=project["name"],
            created_at=project["created_at"],
            modified_at=project["modified_at"],
            version=project["version"],
            inputs=project.get("inputs", {}),
            ui_options=project.get("ui_options", {}),
            results=results,
            optimization=optimization,
            notes=notes,
        )


def default_project_name() -> str:
    return f"Проект {datetime.now().strftime('%Y-%m-%d %H-%M')}"
