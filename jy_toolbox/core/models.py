from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, TYPE_CHECKING
import tkinter as tk

if TYPE_CHECKING:
    from jy_toolbox.ui.base import BaseToolFrame
    from jy_toolbox.ui.app import ToolboxApp

class ToolDefinition:
    key: str
    name: str
    default_category: str
    description: str
    factory: Callable[[tk.Widget, ToolboxApp, dict[str, Any]], BaseToolFrame]
