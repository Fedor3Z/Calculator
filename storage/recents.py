from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List


@dataclass
class RecentItem:
    path: str
    last_opened: str


class RecentProjects:
    def __init__(self, settings_path: Path) -> None:
        self.settings_path = settings_path
        self.settings_path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> List[RecentItem]:
        if not self.settings_path.exists():
            return []
        data = json.loads(self.settings_path.read_text(encoding="utf-8"))
        return [RecentItem(**item) for item in data.get("recents", [])]

    def save(self, items: List[RecentItem]) -> None:
        payload = {"recents": [item.__dict__ for item in items]}
        self.settings_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def add(self, path: Path) -> None:
        items = self.load()
        now = datetime.now().isoformat()
        items = [item for item in items if item.path != str(path)]
        items.insert(0, RecentItem(path=str(path), last_opened=now))
        self.save(items[:10])
