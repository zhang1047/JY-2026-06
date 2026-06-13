from __future__ import annotations

import threading
from typing import Any, Callable
import tkinter as tk
from tkinter import messagebox, ttk

from jy_toolbox.core.constants import *
from jy_toolbox.ui.widgets import make_rounded_button

class BaseToolFrame(ttk.Frame):
    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, style="Surface.TFrame")
        self.app = app
        self.state = state
        self.description = description
        self.description_var = tk.StringVar(value=description)
        self._background_running = False
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar: ttk.Progressbar | None = None

        desc_row = ttk.Frame(self, style="Surface.TFrame")
        desc_row.pack(fill="x", padx=22, pady=(8, 6))
        ttk.Label(desc_row, text="说明：", style="Muted.TLabel").pack(side="left", anchor="n", pady=(3, 0))
        ttk.Label(
            desc_row,
            textvariable=self.description_var,
            style="InlineDescription.TLabel",
            wraplength=780,
            justify="left",
        ).pack(side="left", fill="x", expand=True, padx=(2, 10))
        make_rounded_button(desc_row, "编辑", self.open_description_editor, width=50, height=24).pack(side="right")

    def _description_value(self) -> str:
        return self.description_var.get().strip()

    def _save_description_shortcut(self, _event: tk.Event) -> str:
        self.save_description()
        return "break"

    def open_description_editor(self) -> None:
        DescriptionEditDialog(self)

    def update_description(self, description: str) -> None:
        tool_key = self.app.current_tool_key
        if not tool_key:
            return
        default_description = self.app.tools[tool_key].description
        self.description = description or default_description
        self.description_var.set(self.description)
        self.save_description()

    def reset_description(self) -> None:
        tool_key = self.app.current_tool_key
        if not tool_key:
            return
        self.update_description(self.app.tools[tool_key].description)

    def save_description(self) -> None:
        tool_key = self.app.current_tool_key
        if not tool_key:
            return
        description = self._description_value()
        default_description = self.app.tools[tool_key].description
        descriptions = self.app.config.data.setdefault("tool_descriptions", {})
        if description and description != default_description:
            descriptions[tool_key] = description
        else:
            descriptions.pop(tool_key, None)
        self.description = description or default_description
        self.description_var.set(self.description)
        self.app.config.save()

    def save_state(self) -> None:
        raise NotImplementedError

    def add_progress_bar(self, parent: tk.Widget) -> ttk.Progressbar:
        """为工具动作增加统一进度条；实际耗时任务必须通过后台线程执行。"""
        self.progress_bar = ttk.Progressbar(parent, variable=self.progress_var, maximum=100, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(8, 0))
        return self.progress_bar

    def set_progress(self, value: float, message: str | None = None) -> None:
        self.progress_var.set(max(0, min(100, value)))
        if message is not None and hasattr(self, "status_var"):
            self.status_var.set(message)

    def run_in_background(
        self,
        task: Callable[[Callable[[float, str | None], None]], Any],
        on_success: Callable[[Any], None],
        *,
        start_message: str = "正在后台处理，请稍候……",
        error_message: str = "处理失败，请检查文件、列名或密码本。",
    ) -> None:
        if self._background_running:
            messagebox.showinfo("提示", "当前工具正在后台执行，请等待完成。", parent=self)
            return
        self._background_running = True
        if self.progress_bar is not None:
            self.progress_bar.configure(mode="indeterminate")
            self.progress_bar.start(12)
        self.set_progress(0, start_message)

        def progress(value: float, message: str | None = None) -> None:
            self.after(0, lambda: self.set_progress(value, message))

        def worker() -> None:
            try:
                result = task(progress)
            except Exception as exc:  # noqa: BLE001 - GUI 顶层需要把错误显示给用户
                self.after(0, lambda exc=exc: self._finish_background_error(exc, error_message))
            else:
                self.after(0, lambda: self._finish_background_success(result, on_success))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_background_success(self, result: Any, on_success: Callable[[Any], None]) -> None:
        self._background_running = False
        if self.progress_bar is not None:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
        self.set_progress(100)
        on_success(result)

    def _finish_background_error(self, exc: Exception, error_message: str) -> None:
        self._background_running = False
        if self.progress_bar is not None:
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
        self.set_progress(0, error_message)
        messagebox.showerror("处理失败", str(exc), parent=self)
