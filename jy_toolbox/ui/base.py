from __future__ import annotations

import threading
from pathlib import Path
from typing import Any, Callable
import tkinter as tk
from tkinter import messagebox, ttk

from jy_toolbox.core.constants import *
from jy_toolbox.services.excel_io import list_excel_sheet_names_with_passwords
from jy_toolbox.ui.dialogs import DescriptionEditDialog
from jy_toolbox.ui.widgets import make_rounded_button

class BaseToolFrame(ttk.Frame):
    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, style="Surface.TFrame")
        self.app = app
        self.state = state
        self.description = description
        self.description_var = tk.StringVar(value=description)
        self._background_running = False
        self._sheet_loading_seq: dict[int, int] = {}
        self._sheet_loading_pending: dict[int, str] = {}
        self._sheet_status_before_loading = ""
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

    def add_sheet_selector(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        *,
        hint: str = "导入 Excel 后自动识别 sheet，下拉选择",
    ) -> ttk.Combobox:
        """增加统一的 sheet 下拉框，供当前和后续 Excel 工具复用。"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        combo = ttk.Combobox(parent, textvariable=var, state="readonly", values=())
        combo.grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(parent, text=hint).grid(row=row, column=2, sticky="w", padx=10, pady=8)
        return combo

    def _widget_exists(self, widget: tk.Widget) -> bool:
        """安全判断 Tk 控件是否仍然存在。

        后台线程完成后，用户可能已经切换工具或关闭窗口；此时 Tk 已经销毁了
        原来的 Combobox，再访问它会抛出 ``TclError: invalid command name``。
        """
        try:
            return bool(widget.winfo_exists())
        except tk.TclError:
            return False

    def _schedule_on_ui_thread(self, callback: Callable[[], None]) -> None:
        """尽量把回调派发到 UI 线程；窗口已销毁时直接忽略。"""
        if not self._widget_exists(self):
            return
        try:
            self.after(0, callback)
        except tk.TclError:
            return

    def populate_sheets_async(
        self,
        excel_path: str,
        combo: ttk.Combobox,
        var: tk.StringVar,
        *,
        show_errors: bool = True,
    ) -> None:
        """在后台识别 Excel sheet，避免大文件在选择文件或切换工具时卡住界面。"""
        if not self._widget_exists(combo):
            return
        if not excel_path:
            combo["values"] = ()
            return

        combo_key = id(combo)
        seq = self._sheet_loading_seq.get(combo_key, 0) + 1
        self._sheet_loading_seq[combo_key] = seq
        combo["values"] = ()
        var.set(var.get().strip())
        if hasattr(self, "status_var"):
            if not self._sheet_loading_pending:
                self._sheet_status_before_loading = self.status_var.get()
            self._sheet_loading_pending[combo_key] = Path(excel_path).name
            self._update_sheet_loading_status()

        def worker() -> None:
            try:
                sheet_names = list_excel_sheet_names_with_passwords(
                    Path(excel_path),
                    self.app.config.data.get("passwords", []),
                )
            except Exception as exc:  # noqa: BLE001 - GUI 顶层需要把错误显示给用户
                self._schedule_on_ui_thread(
                    lambda exc=exc: self._finish_sheet_loading_error(combo_key, seq, exc, show_errors)
                )
            else:
                self._schedule_on_ui_thread(
                    lambda: self._finish_sheet_loading_success(combo_key, seq, sheet_names, combo, var)
                )

        threading.Thread(target=worker, daemon=True).start()

    def load_configured_sheets_async(self, *, show_errors: bool = False) -> None:
        """按通用命名约定自动识别当前工具已填写的 Excel 工作表。"""
        mappings = (
            ("input_var", "sheet_combo", "sheet_var"),
            ("account_input_var", "account_sheet_combo", "account_sheet_var"),
            ("post_input_var", "post_sheet_combo", "post_sheet_var"),
            ("dictionary_input_var", "dictionary_sheet_combo", "dictionary_sheet_var"),
        )
        for path_attr, combo_attr, sheet_attr in mappings:
            if all(hasattr(self, attr) for attr in (path_attr, combo_attr, sheet_attr)):
                path_var = getattr(self, path_attr)
                combo = getattr(self, combo_attr)
                sheet_var = getattr(self, sheet_attr)
                self.populate_sheets_async(path_var.get().strip(), combo, sheet_var, show_errors=show_errors)

    def _update_sheet_loading_status(self) -> None:
        """显示当前仍在识别的 Excel，避免多个 sheet 任务互相恢复旧状态。"""
        if not hasattr(self, "status_var"):
            return
        pending_names = list(self._sheet_loading_pending.values())
        if not pending_names:
            self.status_var.set(self._sheet_status_before_loading or "工作表识别完成。")
            self._sheet_status_before_loading = ""
            return
        current_name = pending_names[-1]
        suffix = f"（剩余 {len(pending_names)} 个文件）" if len(pending_names) > 1 else ""
        self.status_var.set(f"正在后台识别工作表：{current_name}……{suffix}")

    def _finish_sheet_loading_success(
        self,
        combo_key: int,
        seq: int,
        sheet_names: list[str],
        combo: ttk.Combobox,
        var: tk.StringVar,
    ) -> None:
        if seq != self._sheet_loading_seq.get(combo_key):
            return
        self._sheet_loading_pending.pop(combo_key, None)
        if not self._widget_exists(combo):
            self._update_sheet_loading_status()
            return
        combo["values"] = sheet_names
        if sheet_names and var.get().strip() not in sheet_names:
            var.set(sheet_names[0])
        self._update_sheet_loading_status()

    def _finish_sheet_loading_error(
        self, combo_key: int, seq: int, exc: Exception, show_errors: bool
    ) -> None:
        if seq != self._sheet_loading_seq.get(combo_key):
            return
        self._sheet_loading_pending.pop(combo_key, None)
        if not self._widget_exists(self):
            return
        self._update_sheet_loading_status()
        message = str(exc)
        if not message.startswith("读取工作表失败"):
            message = f"读取工作表失败：{message}"
        if show_errors:
            messagebox.showwarning("提示", message, parent=self)

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
