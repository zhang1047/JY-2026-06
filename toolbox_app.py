#!/usr/bin/env python3
"""桌面工具箱：分类工具入口、全局密码本、工具独立配置保存。"""
from __future__ import annotations

import ctypes
import json
import os
import random
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_NAME = "JY 临时需求工具箱"
CONFIG_DIR = Path.home() / ".jy_toolbox"
CONFIG_FILE = CONFIG_DIR / "config.json"

COLOR_BG = "#eef1f4"
COLOR_SURFACE = "#f8f9fa"
COLOR_SURFACE_RAISED = "#f3f4f6"
COLOR_FIELD = "#fbfcfd"
COLOR_PRIMARY = "#2da44e"
COLOR_PRIMARY_DARK = "#1f7a3d"
COLOR_PRIMARY_HOVER = "#279247"
COLOR_PRIMARY_PRESSED = "#1f7a3d"
COLOR_TEXT = "#24292f"
COLOR_MUTED = "#57606a"
COLOR_DISABLED = "#8c959f"
COLOR_BORDER = "#d0d7de"
COLOR_BORDER_LIGHT = "#8c959f"
COLOR_ACCENT = "#eaf7ef"
COLOR_ACCENT_TEXT = "#1a7f37"
COLOR_BUTTON = "#f8f9fa"
COLOR_BUTTON_HOVER = "#eef1f4"
COLOR_BUTTON_PRESSED = "#e2e7ec"
COLOR_DANGER = "#cf222e"
COLOR_DANGER_DARK = "#a40e26"
COLOR_DANGER_HOVER = "#a40e26"
COLOR_DANGER_PRESSED = "#82071e"
COLOR_WARNING = "#bf8700"
COLOR_WARNING_DARK = "#9a6700"
COLOR_WARNING_PRESSED = "#7d4e00"
COLOR_TOOL_SELECTED = COLOR_PRIMARY
COLOR_TOOL_UNSELECTED = COLOR_BUTTON
APP_FONT = ("Microsoft YaHei UI", 9)
APP_FONT_BOLD = ("Microsoft YaHei UI", 9, "bold")
APP_FONT_SMALL = ("Microsoft YaHei UI", 8)
BUTTON_RADIUS = 5


def enable_light_title_bar(root: tk.Tk | tk.Toplevel) -> None:
    """在 Windows 上使用浅色标题栏，配合浅灰色整体主题。"""
    if os.name != "nt":
        return
    root.update_idletasks()
    hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
    value = ctypes.c_int(0)
    for attribute in (20, 19):
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))


def mask_password(password: str) -> str:
    """隐藏密码中间字符，例如 qwe123 -> q****3。"""
    if not password:
        return ""
    if len(password) == 1:
        return "*"
    if len(password) == 2:
        return password[0] + "*"
    return password[0] + ("*" * max(1, len(password) - 2)) + password[-1]


@dataclass
class ToolDefinition:
    key: str
    name: str
    default_category: str
    description: str
    factory: Callable[[tk.Widget, "ToolboxApp", dict[str, Any]], "BaseToolFrame"]


BUTTON_PALETTES = {
    "normal": {
        "bg": COLOR_BUTTON,
        "hover": COLOR_BUTTON_HOVER,
        "pressed": COLOR_BUTTON_PRESSED,
        "fg": COLOR_TEXT,
        "border": COLOR_BORDER,
    },
    "primary": {
        "bg": COLOR_PRIMARY,
        "hover": COLOR_PRIMARY_HOVER,
        "pressed": COLOR_PRIMARY_PRESSED,
        "fg": "#ffffff",
        "border": COLOR_PRIMARY_DARK,
    },
    "danger": {
        "bg": COLOR_DANGER,
        "hover": COLOR_DANGER_HOVER,
        "pressed": COLOR_DANGER_PRESSED,
        "fg": "#ffffff",
        "border": COLOR_DANGER_DARK,
    },
    "warning": {
        "bg": COLOR_WARNING,
        "hover": COLOR_WARNING_DARK,
        "pressed": COLOR_WARNING_PRESSED,
        "fg": "#ffffff",
        "border": COLOR_WARNING_DARK,
    },
    "selected": {
        "bg": COLOR_PRIMARY,
        "hover": COLOR_PRIMARY_HOVER,
        "pressed": COLOR_PRIMARY_PRESSED,
        "fg": "#ffffff",
        "border": COLOR_PRIMARY_DARK,
    },
}


class RoundedButton(tk.Canvas):
    """Canvas 绘制的圆角按钮，避开 ttk 在部分系统上直角且过大的默认外观。"""

    def __init__(
        self,
        parent: tk.Widget,
        text: str,
        command: Callable[[], None] | None = None,
        *,
        role: str = "normal",
        width: int | None = None,
        height: int = 26,
        padx: int = 9,
        font: tuple[str, int, str] | tuple[str, int] = APP_FONT,
    ) -> None:
        self.text = text
        self.command = command
        self.role = role if role in BUTTON_PALETTES else "normal"
        self.padx = padx
        self.fixed_width = width
        self._state = "bg"
        self._pressed = False
        palette = BUTTON_PALETTES[self.role]
        super().__init__(
            parent,
            width=width or max(52, len(text) * 11 + padx * 2),
            height=height,
            bg=self._parent_bg(parent),
            highlightthickness=0,
            bd=0,
            cursor="hand2",
        )
        self.configure(takefocus=True)
        self._font = font
        self._items: tuple[int, int] | None = None
        self._palette = palette
        self.bind("<Configure>", lambda _e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<space>", self._on_key)
        self.bind("<Return>", self._on_key)
        self._draw()

    @staticmethod
    def _parent_bg(parent: tk.Widget) -> str:
        try:
            return str(parent.cget("background"))
        except tk.TclError:
            pass
        try:
            style_name = str(parent.cget("style")) or parent.winfo_class()
            styled_bg = ttk.Style(parent).lookup(style_name, "background")
            return str(styled_bg or COLOR_BG)
        except tk.TclError:
            return COLOR_BG

    def _rounded_rect(self, x1: int, y1: int, x2: int, y2: int, radius: int, **kwargs: Any) -> int:
        points = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(points, smooth=True, splinesteps=12, **kwargs)

    def _draw(self) -> None:
        self.delete("all")
        width = max(1, self.winfo_width() or int(self.cget("width")))
        height = max(1, self.winfo_height() or int(self.cget("height")))
        fill = self._palette[self._state]
        border = self._palette["border"]
        self._rounded_rect(1, 1, width - 1, height - 1, BUTTON_RADIUS, fill=fill, outline=border, width=1)
        self.create_text(
            width // 2,
            height // 2,
            text=self.text,
            fill=self._palette["fg"],
            font=self._font,
        )

    def _on_enter(self, _event: tk.Event) -> None:
        if not self._pressed:
            self._state = "hover"
            self._draw()

    def _on_leave(self, _event: tk.Event) -> None:
        self._pressed = False
        self._state = "bg"
        self._draw()

    def _on_press(self, _event: tk.Event) -> None:
        self.focus_set()
        self._pressed = True
        self._state = "pressed"
        self._draw()

    def _on_release(self, event: tk.Event) -> None:
        was_pressed = self._pressed
        self._pressed = False
        inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        self._state = "hover" if inside else "bg"
        self._draw()
        if was_pressed and inside and self.command is not None:
            self.command()

    def _on_key(self, _event: tk.Event) -> str:
        if self.command is not None:
            self.command()
        return "break"


def make_rounded_button(
    parent: tk.Widget,
    text: str,
    command: Callable[[], None] | None = None,
    *,
    role: str = "normal",
    width: int | None = None,
    height: int = 26,
) -> RoundedButton:
    return RoundedButton(parent, text, command, role=role, width=width, height=height)


def rounded_rect_points(x1: int, y1: int, x2: int, y2: int, radius: int = BUTTON_RADIUS) -> list[int]:
    return [
        x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
        x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
        x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
    ]


class ConfigStore:
    """把全局设置、分类设置和每个工具自己的表单状态保存到本地 JSON。"""

    def __init__(self, path: Path = CONFIG_FILE) -> None:
        self.path = path
        self.data: dict[str, Any] = {
            "passwords": [],
            "categories": [],
            "tool_categories": {},
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


class PasswordBookDialog(tk.Toplevel):
    def __init__(self, app: "ToolboxApp") -> None:
        super().__init__(app.root)
        self.app = app
        self.title("密码本")
        self.geometry("520x420")
        self.configure(bg=COLOR_BG)
        enable_light_title_bar(self)
        self.transient(app.root)
        self.grab_set()
        self.check_vars: list[tk.BooleanVar] = []

        ttk.Label(
            self,
            text="用于打开加密文档，可预设多个，会挨个尝试",
            wraplength=420,
        ).pack(anchor="w", padx=14, pady=(14, 8))

        list_outer = ttk.Frame(self)
        list_outer.pack(fill="both", expand=True, padx=14, pady=6)
        self.canvas = tk.Canvas(list_outer, highlightthickness=0, bg=COLOR_SURFACE, bd=0)
        scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="Card.TFrame")
        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", padx=14, pady=(8, 14))
        make_rounded_button(buttons, "新增密码", self.add_password, role="primary", width=82).pack(side="left")
        make_rounded_button(buttons, "删除勾选", self.delete_checked, role="danger", width=82).pack(side="left", padx=8)
        make_rounded_button(buttons, "关闭", self.destroy, width=62).pack(side="right")
        self.refresh()

    def refresh(self) -> None:
        for child in self.inner.winfo_children():
            child.destroy()
        self.check_vars.clear()
        passwords = self.app.config.data.setdefault("passwords", [])
        if not passwords:
            ttk.Label(self.inner, text="暂无密码，请点击“新增密码”。", foreground=COLOR_MUTED).pack(anchor="w", pady=8)
            return
        for idx, password in enumerate(passwords):
            var = tk.BooleanVar(value=False)
            self.check_vars.append(var)
            row = ttk.Frame(self.inner, style="Card.TFrame")
            row.pack(fill="x", pady=3)
            ttk.Checkbutton(row, variable=var).pack(side="left")
            ttk.Label(row, text=f"{idx + 1}. {mask_password(str(password))}").pack(side="left", padx=8)

    def add_password(self) -> None:
        password = simpledialog.askstring("新增密码", "请输入文档打开密码：", parent=self, show="*")
        if password is None:
            return
        password = password.strip()
        if not password:
            messagebox.showwarning("提示", "密码不能为空。", parent=self)
            return
        self.app.config.data.setdefault("passwords", []).append(password)
        self.app.config.save()
        self.refresh()

    def delete_checked(self) -> None:
        passwords = self.app.config.data.setdefault("passwords", [])
        keep = [pwd for pwd, var in zip(passwords, self.check_vars) if not var.get()]
        if len(keep) == len(passwords):
            messagebox.showinfo("提示", "请先勾选要删除的密码。", parent=self)
            return
        self.app.config.data["passwords"] = keep
        self.app.config.save()
        self.refresh()


class DescriptionEditDialog(tk.Toplevel):
    """小型说明编辑弹窗，主界面只保留一两行展示和编辑入口。"""

    def __init__(self, tool_frame: "BaseToolFrame") -> None:
        super().__init__(tool_frame.app.root)
        self.tool_frame = tool_frame
        self.title("编辑工具说明")
        self.geometry("560x320")
        self.configure(bg=COLOR_BG)
        enable_light_title_bar(self)
        self.transient(tool_frame.app.root)
        self.grab_set()

        shell = ttk.Frame(self, style="Surface.TFrame", padding=(16, 14))
        shell.pack(fill="both", expand=True)
        ttk.Label(
            shell,
            text="说明只会在工具顶部简洁展示；这里可随时修改或恢复默认。",
            style="Muted.TLabel",
            wraplength=500,
        ).pack(anchor="w", pady=(0, 10))

        self.text = tk.Text(
            shell,
            height=7,
            wrap="word",
            bg=COLOR_FIELD,
            fg=COLOR_TEXT,
            insertbackground=COLOR_TEXT,
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_PRIMARY,
            font=APP_FONT,
            padx=10,
            pady=8,
        )
        self.text.pack(fill="both", expand=True)
        self.text.insert("1.0", tool_frame.description)
        self.text.focus_set()

        buttons = ttk.Frame(shell, style="Surface.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        make_rounded_button(buttons, "恢复默认", self.restore_default, width=84).pack(side="left")
        make_rounded_button(buttons, "取消", self.destroy, width=62).pack(side="right")
        make_rounded_button(buttons, "保存", self.save, role="primary", width=62).pack(side="right", padx=(0, 8))
        self.bind("<Control-s>", lambda _e: self.save())

    def restore_default(self) -> None:
        tool_key = self.tool_frame.app.current_tool_key
        if not tool_key:
            return
        default_description = self.tool_frame.app.tools[tool_key].description
        self.text.delete("1.0", "end")
        self.text.insert("1.0", default_description)

    def save(self) -> None:
        self.tool_frame.update_description(self.text.get("1.0", "end-1c").strip())
        self.destroy()


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


class PostDedupTool(BaseToolFrame):
    REQUIRED_COLUMNS = ["贴文url", "点赞数", "分享数", "评论数"]

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.input_var = tk.StringVar(value=state.get("input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.sheet_var = tk.StringVar(value=state.get("sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择 Excel 文件后开始处理。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="贴文去重", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "输入 Excel：", self.input_var, self.choose_input)
        self._path_row(form, 1, "输出 Excel：", self.output_var, self.choose_output)
        ttk.Label(form, text="工作表名：").grid(row=2, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.sheet_var).grid(row=2, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=2, column=2, sticky="w", padx=10, pady=8)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始去重", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.input_var.set(path)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_去重后.xlsx")))
        self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存处理结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state(
            "post_dedup",
            {
                "input_path": self.input_var.get().strip(),
                "output_path": self.output_var.get().strip(),
                "sheet_name": self.sheet_var.get().strip(),
            },
        )
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        input_path = self.input_var.get().strip()
        output_path = self.output_var.get().strip()
        sheet_name = self.sheet_var.get().strip() or 0
        if not input_path:
            messagebox.showwarning("提示", "请选择输入 Excel。", parent=self)
            return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            progress(10, "正在后台读取并去重 Excel……")
            return deduplicate_posts_excel(
                Path(input_path),
                Path(output_path),
                self.app.config.data.get("passwords", []),
                sheet_name,
                progress,
            )

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set(
                f"完成：原始 {result['original']} 行，删除 {result['removed']} 行，保留 {result['kept']} 行。输出：{output_path}"
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台执行贴文去重……")


def read_excel_with_passwords(path: Path, passwords: list[str], sheet_name: str | int) -> Any:
    """读取普通或加密 Excel。加密文件会按密码本顺序尝试。"""
    import pandas as pd

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as first_error:  # noqa: BLE001 - 需要判断是否可用密码继续尝试
        if not passwords:
            raise RuntimeError(f"读取失败；如果文件有打开密码，请先在密码本新增密码。原始错误：{first_error}") from first_error
        try:
            import msoffcrypto
        except ImportError as import_error:
            raise RuntimeError(
                "文件可能已加密，但缺少 msoffcrypto-tool 依赖；请执行：pip install msoffcrypto-tool"
            ) from import_error

        last_error: Exception | None = None
        for password in passwords:
            tmp_name = ""
            try:
                with path.open("rb") as source:
                    office_file = msoffcrypto.OfficeFile(source)
                    office_file.load_key(password=str(password))
                    with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as decrypted:
                        tmp_name = decrypted.name
                        office_file.decrypt(decrypted)
                return pd.read_excel(tmp_name, sheet_name=sheet_name)
            except Exception as exc:  # noqa: BLE001 - 尝试下一个密码
                last_error = exc
            finally:
                if tmp_name:
                    try:
                        os.unlink(tmp_name)
                    except OSError:
                        pass
        raise RuntimeError("读取失败：已按密码本逐个尝试，但没有密码可以打开该 Excel。") from last_error


def deduplicate_posts_excel(
    input_path: Path,
    output_path: Path,
    passwords: list[str],
    sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    import pandas as pd

    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    df = read_excel_with_passwords(input_path, passwords, sheet_name)
    if progress is not None:
        progress(35, "已读取 Excel，正在检查必要列……")
    missing = [col for col in PostDedupTool.REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要列：{', '.join(missing)}")

    work = df.copy()
    if progress is not None:
        progress(55, "正在计算重复贴文保留规则……")
    score_cols = ["点赞数", "分享数", "评论数"]
    numeric_scores = work[score_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    work["__dedup_score__"] = numeric_scores.sum(axis=1)
    work["__dedup_random__"] = [random.random() for _ in range(len(work))]
    work["__dedup_order__"] = range(len(work))
    kept = (
        work.sort_values(["贴文url", "__dedup_score__", "__dedup_random__"], ascending=[True, False, False])
        .drop_duplicates(subset=["贴文url"], keep="first")
        .sort_values("__dedup_order__")
        .drop(columns=["__dedup_score__", "__dedup_random__", "__dedup_order__"])
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(85, "正在写出处理结果……")
    kept.to_excel(output_path, index=False)
    return {"original": len(df), "removed": len(df) - len(kept), "kept": len(kept)}

class PostTypeRatioTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    LEGACY_OUTPUT_COLUMN = "帖子类型"
    OUTPUT_COLUMNS = ("文字帖占比", "图片帖占比", "视频帖占比", "未识别")
    REQUIRED_POST_COLUMNS = ["主页url", "图片附件", "创作类型", "标题", "帖子正文"]

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel 和贴文 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="贴文类型占比（%）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        ttk.Label(form, text="账号表工作表：").grid(row=3, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.account_sheet_var).grid(row=3, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=3, column=2, sticky="w", padx=10, pady=8)
        ttk.Label(form, text="贴文表工作表：").grid(row=4, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.post_sheet_var).grid(row=4, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=4, column=2, sticky="w", padx=10, pady=8)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择账号 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.account_input_var.set(path)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_帖子类型占比.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path)
            self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存账号表处理结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state(
            "post_type_ratio",
            {
                "account_input_path": self.account_input_var.get().strip(),
                "post_input_path": self.post_input_var.get().strip(),
                "output_path": self.output_var.get().strip(),
                "account_sheet_name": self.account_sheet_var.get().strip(),
                "post_sheet_name": self.post_sheet_var.get().strip(),
            },
        )
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip()
        post_input_path = self.post_input_var.get().strip()
        output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0
        post_sheet_name = self.post_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self)
            return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self)
            return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_post_type_ratios_excel(
                Path(account_input_path),
                Path(post_input_path),
                Path(output_path),
                self.app.config.data.get("passwords", []),
                account_sheet_name,
                post_sheet_name,
                progress,
            )

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set(
                "完成：账号 {accounts} 行，贴文 {posts} 行，已匹配 {matched_accounts} 个账号，"
                "有贴文占比的账号 {typed_accounts} 个。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    matched_accounts=result["matched_accounts"],
                    typed_accounts=result["typed_accounts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计贴文类型占比……")


def _is_non_empty_cell(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return bool(text) and text.lower() != "nan"


def _normalized_key(value: Any) -> str:
    return str(value).strip() if _is_non_empty_cell(value) else ""


def classify_post_type(row: Any) -> str:
    homepage_url = "" if not _is_non_empty_cell(row.get("主页url")) else str(row.get("主页url"))
    post_url = "" if not _is_non_empty_cell(row.get("贴文url")) else str(row.get("贴文url"))
    if "/videos/" in homepage_url.lower() or "/videos/" in post_url.lower():
        return "视频"
    attachments = "" if not _is_non_empty_cell(row.get("图片附件")) else str(row.get("图片附件"))
    if attachments.count("origin_url_md5") >= 2:
        return "图片"
    if str(row.get("创作类型", "")).strip().lower() == "common":
        if _is_non_empty_cell(row.get("标题")) or _is_non_empty_cell(row.get("帖子正文")):
            return "文字"
    return ""


def _format_post_type_ratios(counts: dict[str, int]) -> dict[str, str]:
    total = sum(counts.values())
    if total <= 0:
        return {column: "" for column in PostTypeRatioTool.OUTPUT_COLUMNS}

    text_ratio = round(counts.get("文字", 0) / total * 100, 2)
    image_ratio = round(counts.get("图片", 0) / total * 100, 2)
    video_ratio = round(counts.get("视频", 0) / total * 100, 2)
    unidentified_ratio = round(100 - text_ratio - image_ratio - video_ratio, 2)
    return {
        "文字帖占比": f"{text_ratio:.2f}%",
        "图片帖占比": f"{image_ratio:.2f}%",
        "视频帖占比": f"{video_ratio:.2f}%",
        "未识别": f"{unidentified_ratio:.2f}%",
    }


def calculate_post_type_ratios_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    import pandas as pd

    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if PostTypeRatioTool.ACCOUNT_URL_COLUMN not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{PostTypeRatioTool.ACCOUNT_URL_COLUMN}")
    missing_posts = [col for col in PostTypeRatioTool.REQUIRED_POST_COLUMNS if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(50, "正在识别每条贴文的文字、图片、视频类型……")
    work = post_df.copy()
    work["__post_type__"] = work.apply(classify_post_type, axis=1)
    work["__post_type__"] = work["__post_type__"].replace("", "未识别")

    empty_ratios = {column: "" for column in PostTypeRatioTool.OUTPUT_COLUMNS}
    ratios_by_homepage: dict[str, dict[str, str]] = {}
    if not work.empty:
        grouped = work.groupby(PostTypeRatioTool.POST_URL_COLUMN)["__post_type__"].value_counts()
        for homepage, counts_series in grouped.groupby(level=0):
            counts = {str(type_name): int(count) for (_, type_name), count in counts_series.items()}
            ratios_by_homepage[_normalized_key(homepage)] = _format_post_type_ratios(counts)

    if progress is not None:
        progress(75, "正在写回账号表最后四列类型占比……")
    output_df = account_df.copy()
    drop_columns = [
        column
        for column in (PostTypeRatioTool.LEGACY_OUTPUT_COLUMN, *PostTypeRatioTool.OUTPUT_COLUMNS)
        if column in output_df.columns
    ]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df[PostTypeRatioTool.ACCOUNT_URL_COLUMN].map(_normalized_key)
    for column in PostTypeRatioTool.OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(lambda key, column=column: ratios_by_homepage.get(key, empty_ratios).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(90, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(_normalized_key(v) for v in post_df[PostTypeRatioTool.POST_URL_COLUMN])).sum())
    typed_accounts = int(output_df[PostTypeRatioTool.OUTPUT_COLUMNS[0]].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "matched_accounts": matched_accounts,
        "typed_accounts": typed_accounts,
    }



class ToolListDialog(tk.Toplevel):
    """以弹窗承载工具列表，避免主界面左侧长期占位。"""

    def __init__(self, app: "ToolboxApp") -> None:
        super().__init__(app.root)
        self.app = app
        self.title("工具列表")
        self.geometry("420x560")
        self.minsize(360, 420)
        self.configure(bg=COLOR_BG)
        enable_light_title_bar(self)
        self.transient(app.root)
        self.protocol("WM_DELETE_WINDOW", self.close)

        shell = ttk.Frame(self, style="Surface.TFrame", padding=(18, 16))
        shell.pack(fill="both", expand=True)

        header = ttk.Frame(shell, style="Surface.TFrame")
        header.pack(fill="x", pady=(0, 14))
        title_group = ttk.Frame(header, style="Surface.TFrame")
        title_group.pack(side="left", fill="x", expand=True)
        ttk.Label(title_group, text="工具列表", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(title_group, text="按分类管理工具，拖动工具可移动分类", style="Muted.TLabel").pack(anchor="w", pady=(3, 0))
        make_rounded_button(header, "＋ 新增分类", app.add_category, role="primary", width=96).pack(side="right")

        list_card = ttk.Frame(shell, style="Card.TFrame", padding=(12, 12))
        list_card.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(list_card, highlightthickness=0, bg=COLOR_SURFACE, bd=0)
        scrollbar = ttk.Scrollbar(list_card, orient="vertical", command=self.canvas.yview)
        self.container = ttk.Frame(self.canvas, style="Card.TFrame")
        self.container.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.window_id = self.canvas.create_window((0, 0), window=self.container, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.window_id, width=e.width))
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        footer = ttk.Frame(shell, style="Surface.TFrame")
        footer.pack(fill="x", pady=(14, 0))
        make_rounded_button(footer, "关闭", self.close, width=62).pack(side="right")

    def close(self) -> None:
        self.app.tool_list_dialog = None
        self.destroy()


class ToolboxApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1080x720")
        enable_light_title_bar(self.root)
        self.config = ConfigStore()
        self.tools: dict[str, ToolDefinition] = {}
        self.current_tool_frame: BaseToolFrame | None = None
        self.current_tool_key: str | None = None
        self.drag_data: dict[str, Any] = {}
        self.category_frames: dict[str, tk.Frame] = {}
        self.category_body_frames: dict[str, ttk.Frame] = {}
        self.category_drop_widgets: dict[str, set[tk.Widget]] = {}
        self.tool_list_dialog: ToolListDialog | None = None
        self._configure_styles()
        self._register_tools()
        self._ensure_defaults()
        self._build_layout()
        self.refresh_tool_list()
        self.open_tool(next(iter(self.tools)))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _register_tools(self) -> None:
        self.add_tool(
            ToolDefinition(
                key="post_dedup",
                name="贴文去重",
                default_category="Excel 工具",
                description=(
                    "说明：根据“贴文url”列去重。若同一 URL 有重复行，会比较“点赞数”“分享数”“评论数”三列的数值总和，"
                    "优先保留总和更大的记录；如果总和相同，则随机保留其中一条。"
                ),
                factory=lambda parent, app, state: PostDedupTool(
                    parent, app, state, app.get_tool_description("post_dedup")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="post_type_ratio",
                name="贴文类型占比（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联，"
                    "按规则统计每个账号文字、图片、视频、未识别贴文占比，并在账号表最后新增四列占比。"
                ),
                factory=lambda parent, app, state: PostTypeRatioTool(
                    parent, app, state, app.get_tool_description("post_type_ratio")
                ),
            )
        )

    def add_tool(self, tool: ToolDefinition) -> None:
        self.tools[tool.key] = tool

    def get_tool_description(self, key: str) -> str:
        descriptions = self.config.data.setdefault("tool_descriptions", {})
        custom_description = descriptions.get(key)
        if isinstance(custom_description, str) and custom_description.strip():
            return custom_description.strip()
        return self.tools[key].description

    def _ensure_defaults(self) -> None:
        categories = self.config.data.setdefault("categories", [])
        if not categories:
            categories.extend(sorted({tool.default_category for tool in self.tools.values()}))
        tool_categories = self.config.data.setdefault("tool_categories", {})
        for key, tool in self.tools.items():
            assigned = tool_categories.get(key)
            if assigned not in categories:
                if tool.default_category not in categories:
                    categories.append(tool.default_category)
                tool_categories[key] = tool.default_category
        self.config.save()

    def _configure_styles(self) -> None:
        self.root.configure(bg=COLOR_BG)
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.root.option_add("*Font", APP_FONT)
        self.root.option_add("*Background", COLOR_BG)
        self.root.option_add("*Foreground", COLOR_TEXT)
        self.root.option_add("*Entry.Background", COLOR_FIELD)
        self.root.option_add("*Entry.Foreground", COLOR_TEXT)
        self.root.option_add("*Entry.insertBackground", COLOR_TEXT)
        self.root.option_add("*Text.Background", COLOR_FIELD)
        self.root.option_add("*Text.Foreground", COLOR_TEXT)
        self.root.option_add("*Text.insertBackground", COLOR_TEXT)
        self.root.option_add("*selectBackground", COLOR_PRIMARY_DARK)
        self.root.option_add("*selectForeground", "#ffffff")

        style.configure("TFrame", background=COLOR_BG)
        style.configure("Surface.TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_SURFACE, relief="flat")
        style.configure("Info.TFrame", background=COLOR_ACCENT, relief="flat")
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=APP_FONT)
        style.configure("HeroTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 15, "bold"))
        style.configure("HeroSubtitle.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=APP_FONT_SMALL)
        style.configure("SectionTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=APP_FONT_SMALL)
        style.configure("CardMuted.TLabel", background=COLOR_SURFACE, foreground=COLOR_MUTED, font=APP_FONT_SMALL)
        style.configure("InlineDescription.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=APP_FONT_SMALL)
        style.configure("Info.TLabel", background=COLOR_ACCENT, foreground=COLOR_ACCENT_TEXT, font=APP_FONT)
        style.configure(
            "TButton",
            font=APP_FONT,
            padding=(8, 3),
            borderwidth=1,
            relief="flat",
            background=COLOR_BUTTON,
            foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
        )
        style.map(
            "TButton",
            background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED), ("disabled", COLOR_SURFACE)],
            foreground=[("disabled", COLOR_DISABLED)],
        )
        style.configure("Primary.TButton", background=COLOR_PRIMARY, foreground="#ffffff", bordercolor=COLOR_PRIMARY_DARK)
        style.map("Primary.TButton", background=[("active", COLOR_PRIMARY_HOVER), ("pressed", COLOR_PRIMARY_PRESSED)])
        style.configure("Green.TButton", background=COLOR_PRIMARY, foreground="#ffffff", padding=(8, 3), bordercolor=COLOR_PRIMARY_DARK)
        style.map("Green.TButton", background=[("active", COLOR_PRIMARY_HOVER), ("pressed", COLOR_PRIMARY_PRESSED)])
        style.configure("Small.TButton", background=COLOR_BUTTON, foreground=COLOR_TEXT, padding=(7, 3), bordercolor=COLOR_BORDER)
        style.map("Small.TButton", background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED)])
        style.configure("Danger.TButton", background=COLOR_DANGER, foreground="#ffffff", padding=(7, 3), bordercolor=COLOR_DANGER_DARK)
        style.map("Danger.TButton", background=[("active", COLOR_DANGER_HOVER), ("pressed", COLOR_DANGER_PRESSED)])
        style.configure("Tool.TButton", background=COLOR_BUTTON, foreground=COLOR_TEXT, padding=(8, 3), bordercolor=COLOR_BORDER)
        style.map("Tool.TButton", background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED)])
        style.configure("SelectedTool.TButton", background=COLOR_PRIMARY, foreground="#ffffff", padding=(8, 3), bordercolor=COLOR_PRIMARY_DARK)
        style.map("SelectedTool.TButton", background=[("active", COLOR_PRIMARY_HOVER), ("pressed", COLOR_PRIMARY_PRESSED)])
        style.configure(
            "TEntry",
            fieldbackground=COLOR_FIELD,
            background=COLOR_FIELD,
            foreground=COLOR_TEXT,
            insertcolor=COLOR_TEXT,
            padding=(6, 3),
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
        )
        style.map(
            "TEntry",
            fieldbackground=[("focus", COLOR_SURFACE_RAISED), ("disabled", COLOR_BG)],
            foreground=[("disabled", COLOR_DISABLED)],
            bordercolor=[("focus", COLOR_BORDER_LIGHT)],
        )
        style.configure(
            "TCombobox",
            fieldbackground=COLOR_FIELD,
            background=COLOR_BUTTON,
            foreground=COLOR_TEXT,
            arrowcolor=COLOR_MUTED,
            bordercolor=COLOR_BORDER,
            lightcolor=COLOR_BORDER,
            darkcolor=COLOR_BORDER,
            padding=(6, 3),
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", COLOR_FIELD), ("focus", COLOR_SURFACE_RAISED), ("disabled", COLOR_BG)],
            foreground=[("readonly", COLOR_TEXT), ("disabled", COLOR_DISABLED)],
            background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED)],
            bordercolor=[("focus", COLOR_BORDER_LIGHT)],
        )
        style.configure(
            "TCheckbutton",
            background=COLOR_SURFACE,
            foreground=COLOR_TEXT,
            font=APP_FONT,
            indicatorcolor=COLOR_FIELD,
            bordercolor=COLOR_BORDER,
            focuscolor=COLOR_SURFACE,
        )
        style.map(
            "TCheckbutton",
            background=[("active", COLOR_SURFACE)],
            foreground=[("active", COLOR_TEXT), ("disabled", COLOR_DISABLED)],
            indicatorcolor=[("selected", COLOR_PRIMARY), ("!selected", COLOR_FIELD)],
        )
        style.configure(
            "Vertical.TScrollbar",
            background=COLOR_BUTTON,
            troughcolor=COLOR_SURFACE,
            bordercolor=COLOR_BORDER,
            arrowcolor=COLOR_MUTED,
            lightcolor=COLOR_BUTTON,
            darkcolor=COLOR_BUTTON,
        )
        style.map("Vertical.TScrollbar", background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED)])
        style.configure("Horizontal.TScrollbar", background=COLOR_BUTTON, troughcolor=COLOR_SURFACE, bordercolor=COLOR_BORDER, arrowcolor=COLOR_MUTED, lightcolor=COLOR_BUTTON, darkcolor=COLOR_BUTTON)
        style.map("Horizontal.TScrollbar", background=[("active", COLOR_BUTTON_HOVER), ("pressed", COLOR_BUTTON_PRESSED)])
        style.configure("TLabelframe", background=COLOR_BG, bordercolor=COLOR_BORDER, relief="solid")
        style.configure("TLabelframe.Label", background=COLOR_BG, foreground=COLOR_TEXT, font=APP_FONT_BOLD)
        style.configure("Card.TLabelframe", background=COLOR_SURFACE, bordercolor=COLOR_BORDER, relief="solid")
        style.configure("Card.TLabelframe.Label", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=APP_FONT_BOLD)

    def _build_layout(self) -> None:
        top = ttk.Frame(self.root, style="Surface.TFrame", padding=(24, 12))
        top.pack(fill="x")
        title_group = ttk.Frame(top, style="Surface.TFrame")
        title_group.pack(side="left", fill="x", expand=True)
        ttk.Label(title_group, text=APP_NAME, style="HeroTitle.TLabel").pack(anchor="w")
        actions = ttk.Frame(top, style="Surface.TFrame")
        actions.pack(side="right")
        make_rounded_button(actions, "工具列表", self.open_tool_list, role="primary", width=82).pack(side="left", padx=(0, 10))
        make_rounded_button(actions, "密码本", self.open_password_book, width=68).pack(side="left")

        body = ttk.Frame(self.root, style="Surface.TFrame", padding=(24, 0, 24, 24))
        body.pack(fill="both", expand=True)
        self.content = ttk.Frame(body, style="Surface.TFrame")
        self.content.pack(fill="both", expand=True)

    def open_password_book(self) -> None:
        PasswordBookDialog(self)

    def open_tool_list(self) -> None:
        if self.tool_list_dialog is not None and self.tool_list_dialog.winfo_exists():
            self.tool_list_dialog.lift()
            self.tool_list_dialog.focus_force()
            return
        self.tool_list_dialog = ToolListDialog(self)
        self.refresh_tool_list()

    def refresh_tool_list(self) -> None:
        if self.tool_list_dialog is None or not self.tool_list_dialog.winfo_exists():
            return
        container = self.tool_list_dialog.container
        for child in container.winfo_children():
            child.destroy()
        self.category_frames.clear()
        self.category_body_frames.clear()
        self.category_drop_widgets.clear()
        categories = list(self.config.data.setdefault("categories", []))
        for category in categories:
            outer = tk.Canvas(container, highlightthickness=0, height=96, bg=COLOR_SURFACE, bd=0)
            outer.pack(fill="x", padx=2, pady=8)
            card_id = outer.create_polygon(
                rounded_rect_points(2, 2, 10, 10, 8),
                smooth=True,
                splinesteps=12,
                fill=COLOR_SURFACE,
                outline=COLOR_BORDER,
                width=1,
            )
            inner = ttk.Frame(outer, style="Card.TFrame", padding=(12, 10))
            window_id = outer.create_window((10, 10), window=inner, anchor="nw")

            def resize(event: tk.Event, canvas: tk.Canvas = outer, win: int = window_id, card: int = card_id) -> None:
                width = max(220, canvas.winfo_width() - 20)
                canvas.itemconfigure(win, width=width)
                needed = event.height + 20
                canvas.configure(height=needed)
                canvas.coords(card, *rounded_rect_points(2, 2, max(4, canvas.winfo_width() - 2), needed - 2, 8))

            inner.bind("<Configure>", resize)
            outer.bind(
                "<Configure>",
                lambda e, c=outer, card=card_id: c.coords(
                    card,
                    *rounded_rect_points(2, 2, max(4, e.width - 2), max(4, int(c.cget("height")) - 2), 8),
                ),
            )
            self.category_frames[category] = outer
            self.category_drop_widgets[category] = {outer, inner}

            header = ttk.Frame(inner, style="Card.TFrame")
            header.pack(fill="x", pady=(0, 6))
            ttk.Label(
                header,
                text=category,
                background=COLOR_SURFACE,
                foreground=COLOR_TEXT,
                font=APP_FONT_BOLD,
            ).pack(side="left")
            make_rounded_button(header, "改名", lambda c=category: self.rename_category(c), width=48, height=24).pack(side="right")
            make_rounded_button(header, "删除", lambda c=category: self.delete_category(c), role="danger", width=48, height=24).pack(side="right", padx=6)

            body = ttk.Frame(inner, style="Card.TFrame")
            body.pack(fill="x")
            self.category_body_frames[category] = body
            self.category_drop_widgets[category].add(body)
            body.bind("<ButtonRelease-1>", lambda _e, c=category: self.drop_tool_to_category(c))

        for key, tool in self.tools.items():
            category = self.config.data.setdefault("tool_categories", {}).get(key, tool.default_category)
            body = self.category_body_frames.get(category)
            if body is None:
                continue
            role = "selected" if key == self.current_tool_key else "normal"
            btn = make_rounded_button(
                body,
                tool.name,
                lambda k=key: self.open_tool(k),
                role=role,
                height=26,
            )
            btn.pack(fill="x", pady=4)
            btn.bind("<ButtonPress-1>", lambda _e, k=key: self.start_drag(k), add="+")
            btn.bind("<ButtonRelease-1>", self.finish_drag, add="+")

    def start_drag(self, tool_key: str) -> None:
        self.drag_data = {"tool_key": tool_key}

    def finish_drag(self, event: tk.Event) -> None:
        tool_key = self.drag_data.get("tool_key")
        if not tool_key:
            return
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        while widget is not None:
            for category, widgets in self.category_drop_widgets.items():
                if widget in widgets:
                    self.move_tool(tool_key, category)
                    self.drag_data = {}
                    return
            widget = widget.master
        self.drag_data = {}

    def drop_tool_to_category(self, category: str) -> None:
        tool_key = self.drag_data.get("tool_key")
        if tool_key:
            self.move_tool(tool_key, category)
            self.drag_data = {}

    def move_tool(self, tool_key: str, category: str) -> None:
        self.config.data.setdefault("tool_categories", {})[tool_key] = category
        self.config.save()
        self.refresh_tool_list()

    def add_category(self) -> None:
        name = simpledialog.askstring("新增分类", "请输入分类名称：", parent=self.root)
        if not name:
            return
        name = name.strip()
        categories = self.config.data.setdefault("categories", [])
        if not name or name in categories:
            messagebox.showwarning("提示", "分类名称不能为空或重复。", parent=self.root)
            return
        categories.append(name)
        self.config.save()
        self.refresh_tool_list()

    def rename_category(self, old_name: str) -> None:
        new_name = simpledialog.askstring("分类改名", "请输入新的分类名称：", initialvalue=old_name, parent=self.root)
        if not new_name:
            return
        new_name = new_name.strip()
        categories = self.config.data.setdefault("categories", [])
        if not new_name or (new_name in categories and new_name != old_name):
            messagebox.showwarning("提示", "分类名称不能为空或重复。", parent=self.root)
            return
        self.config.data["categories"] = [new_name if c == old_name else c for c in categories]
        for key, category in list(self.config.data.setdefault("tool_categories", {}).items()):
            if category == old_name:
                self.config.data["tool_categories"][key] = new_name
        self.config.save()
        self.refresh_tool_list()

    def delete_category(self, category: str) -> None:
        categories = self.config.data.setdefault("categories", [])
        if len(categories) <= 1:
            messagebox.showwarning("提示", "至少保留一个分类。", parent=self.root)
            return
        target = next((c for c in categories if c != category), None)
        if not messagebox.askyesno("删除分类", f"确认删除分类“{category}”？其中工具将移动到“{target}”。", parent=self.root):
            return
        self.config.data["categories"] = [c for c in categories if c != category]
        for key, assigned in list(self.config.data.setdefault("tool_categories", {}).items()):
            if assigned == category:
                self.config.data["tool_categories"][key] = target
        self.config.save()
        self.refresh_tool_list()

    def open_tool(self, key: str) -> None:
        if self.current_tool_frame is not None:
            try:
                self.current_tool_frame.save_description()
                self.current_tool_frame.save_state()
            except Exception:
                pass
            self.current_tool_frame.destroy()
        self.current_tool_key = key
        tool = self.tools[key]
        state = self.config.get_tool_state(key)
        self.current_tool_frame = tool.factory(self.content, self, state)
        self.current_tool_frame.pack(fill="both", expand=True)
        self.refresh_tool_list()

    def on_close(self) -> None:
        if self.current_tool_frame is not None:
            self.current_tool_frame.save_description()
            self.current_tool_frame.save_state()
        self.config.save()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    ToolboxApp().run()


if __name__ == "__main__":
    main()
