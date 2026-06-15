from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from jy_toolbox.core.constants import CONFIG_DIR, CONFIG_FILE

class ConfigStore:
    """把全局设置、分组设置和每个工具自己的表单状态保存到本地 JSON。"""

    def __init__(self, path: Path = CONFIG_FILE) -> None:
        self.path = path
        self.data: dict[str, Any] = {
            "passwords": [],
            "categories": [],
            "tool_categories": {},
            "tool_orders": {},
            "tool_states": {},
            "tool_descriptions": {},
        }
        self.load()

    def load(self) -> None:
        if not self.path.exists():
            return
        try:
            loaded = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        if isinstance(loaded, dict):
            self.data.update(loaded)

    def save(self) -> None:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")

    def get_tool_state(self, key: str) -> dict[str, Any]:
        states = self.data.setdefault("tool_states", {})
        state = states.setdefault(key, {})
        return state if isinstance(state, dict) else {}

    def set_tool_state(self, key: str, state: dict[str, Any]) -> None:
        self.data.setdefault("tool_states", {})[key] = state
        self.save()
