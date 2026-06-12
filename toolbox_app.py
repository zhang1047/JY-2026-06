#!/usr/bin/env python3
"""桌面工具箱：分类工具入口、全局密码本、工具独立配置保存。"""
from __future__ import annotations

import json
import os
import random
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_NAME = "JY 临时需求工具箱"
CONFIG_DIR = Path.home() / ".jy_toolbox"
CONFIG_FILE = CONFIG_DIR / "config.json"

COLOR_BG = "#f6f7f8"
COLOR_SURFACE = "#ffffff"
COLOR_PRIMARY = "#18a558"
COLOR_PRIMARY_DARK = "#159149"
COLOR_TEXT = "#1f2933"
COLOR_MUTED = "#666666"
COLOR_BORDER = "#d9d9d9"
COLOR_ACCENT = "#edf8f2"
COLOR_DANGER = "#d93025"
COLOR_DANGER_DARK = "#b3261e"
COLOR_TOOL_SELECTED = "#2d2d2d"
COLOR_TOOL_UNSELECTED = "#f4f4f4"


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
        self.canvas = tk.Canvas(list_outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        buttons = ttk.Frame(self)
        buttons.pack(fill="x", padx=14, pady=(8, 14))
        ttk.Button(buttons, text="新增密码", command=self.add_password, style="Green.TButton").pack(side="left")
        ttk.Button(buttons, text="删除勾选", command=self.delete_checked, style="Danger.TButton").pack(side="left", padx=8)
        ttk.Button(buttons, text="关闭", command=self.destroy, style="Small.TButton").pack(side="right")
        self.refresh()

    def refresh(self) -> None:
        for child in self.inner.winfo_children():
            child.destroy()
        self.check_vars.clear()
        passwords = self.app.config.data.setdefault("passwords", [])
        if not passwords:
            ttk.Label(self.inner, text="暂无密码，请点击“新增密码”。", foreground="#666").pack(anchor="w", pady=8)
            return
        for idx, password in enumerate(passwords):
            var = tk.BooleanVar(value=False)
            self.check_vars.append(var)
            row = ttk.Frame(self.inner)
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
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            insertbackground=COLOR_TEXT,
            relief="solid",
            bd=1,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_PRIMARY,
            font=("Microsoft YaHei UI", 10),
            padx=10,
            pady=8,
        )
        self.text.pack(fill="both", expand=True)
        self.text.insert("1.0", tool_frame.description)
        self.text.focus_set()

        buttons = ttk.Frame(shell, style="Surface.TFrame")
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="恢复默认", command=self.restore_default, style="Small.TButton").pack(side="left")
        ttk.Button(buttons, text="取消", command=self.destroy, style="Small.TButton").pack(side="right")
        ttk.Button(buttons, text="保存", command=self.save, style="Green.TButton").pack(side="right", padx=(0, 8))
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

        desc_row = ttk.Frame(self, style="Surface.TFrame")
        desc_row.pack(fill="x", padx=22, pady=(18, 6))
        ttk.Label(desc_row, text="说明：", style="Muted.TLabel").pack(side="left", anchor="n", pady=(3, 0))
        ttk.Label(
            desc_row,
            textvariable=self.description_var,
            style="InlineDescription.TLabel",
            wraplength=780,
            justify="left",
        ).pack(side="left", fill="x", expand=True, padx=(2, 10))
        ttk.Button(desc_row, text="编辑", command=self.open_description_editor, style="Small.TButton").pack(side="right")

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
        form = ttk.LabelFrame(self, text="贴文去重", style="Card.TLabelframe", padding=(14, 12))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "输入 Excel：", self.input_var, self.choose_input)
        self._path_row(form, 1, "输出 Excel：", self.output_var, self.choose_output)
        ttk.Label(form, text="工作表名：").grid(row=2, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.sheet_var).grid(row=2, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=2, column=2, sticky="w", padx=10, pady=8)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        ttk.Button(actions, text="开始去重", command=self.run, style="Primary.TButton").pack(side="left")
        ttk.Button(actions, text="保存当前填写", command=self.save_state, style="Small.TButton").pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(14, 12))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        ttk.Button(parent, text="浏览", command=command).grid(row=row, column=2, padx=10, pady=8)

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
        try:
            result = deduplicate_posts_excel(
                Path(input_path),
                Path(output_path),
                self.app.config.data.get("passwords", []),
                sheet_name,
            )
        except Exception as exc:  # noqa: BLE001 - GUI 顶层需要把错误显示给用户
            messagebox.showerror("处理失败", str(exc), parent=self)
            self.status_var.set("处理失败，请检查文件、列名或密码本。")
            return
        self.status_var.set(
            f"完成：原始 {result['original']} 行，删除 {result['removed']} 行，保留 {result['kept']} 行。输出：{output_path}"
        )
        messagebox.showinfo("完成", self.status_var.get(), parent=self)


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


def deduplicate_posts_excel(input_path: Path, output_path: Path, passwords: list[str], sheet_name: str | int = 0) -> dict[str, int]:
    import pandas as pd

    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    df = read_excel_with_passwords(input_path, passwords, sheet_name)
    missing = [col for col in PostDedupTool.REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要列：{', '.join(missing)}")

    work = df.copy()
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
    kept.to_excel(output_path, index=False)
    return {"original": len(df), "removed": len(df) - len(kept), "kept": len(kept)}


class ToolListDialog(tk.Toplevel):
    """以弹窗承载工具列表，避免主界面左侧长期占位。"""

    def __init__(self, app: "ToolboxApp") -> None:
        super().__init__(app.root)
        self.app = app
        self.title("工具列表")
        self.geometry("420x560")
        self.minsize(360, 420)
        self.configure(bg=COLOR_BG)
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
        ttk.Button(header, text="＋ 新增分类", command=app.add_category, style="Green.TButton").pack(side="right")

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
        ttk.Button(footer, text="关闭", command=self.close, style="Small.TButton").pack(side="right")

    def close(self) -> None:
        self.app.tool_list_dialog = None
        self.destroy()


class ToolboxApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1080x720")
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
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Surface.TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_SURFACE, relief="flat")
        style.configure("Info.TFrame", background=COLOR_ACCENT, relief="flat")
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 10))
        style.configure("HeroTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("HeroSubtitle.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=("Microsoft YaHei UI", 10))
        style.configure("SectionTitle.TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("Muted.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=("Microsoft YaHei UI", 9))
        style.configure("CardMuted.TLabel", background=COLOR_SURFACE, foreground=COLOR_MUTED, font=("Microsoft YaHei UI", 9))
        style.configure("InlineDescription.TLabel", background=COLOR_BG, foreground=COLOR_MUTED, font=("Microsoft YaHei UI", 10))
        style.configure("Info.TLabel", background=COLOR_ACCENT, foreground=COLOR_PRIMARY_DARK, font=("Microsoft YaHei UI", 10))
        style.configure(
            "TButton",
            font=("Microsoft YaHei UI", 10, "bold"),
            padding=(12, 7),
            borderwidth=1,
            relief="flat",
            background=COLOR_TOOL_UNSELECTED,
            foreground=COLOR_MUTED,
        )
        style.map("TButton", background=[("active", "#ececec"), ("pressed", "#e3e3e3")])
        style.configure("Primary.TButton", background=COLOR_PRIMARY, foreground="#ffffff", bordercolor=COLOR_PRIMARY_DARK)
        style.map("Primary.TButton", background=[("active", COLOR_PRIMARY_DARK), ("pressed", COLOR_PRIMARY_DARK)])
        style.configure("Green.TButton", background=COLOR_PRIMARY, foreground="#ffffff", padding=(12, 5), bordercolor=COLOR_PRIMARY_DARK)
        style.map("Green.TButton", background=[("active", COLOR_PRIMARY_DARK), ("pressed", COLOR_PRIMARY_DARK)])
        style.configure("Small.TButton", background=COLOR_TOOL_UNSELECTED, foreground=COLOR_MUTED, padding=(10, 5), bordercolor=COLOR_BORDER)
        style.map("Small.TButton", background=[("active", "#ececec"), ("pressed", "#e3e3e3")])
        style.configure("Danger.TButton", background=COLOR_DANGER, foreground="#ffffff", padding=(10, 5), bordercolor=COLOR_DANGER_DARK)
        style.map("Danger.TButton", background=[("active", COLOR_DANGER_DARK), ("pressed", COLOR_DANGER_DARK)])
        style.configure("Tool.TButton", background=COLOR_TOOL_UNSELECTED, foreground=COLOR_MUTED, padding=(18, 8), bordercolor=COLOR_BORDER)
        style.map("Tool.TButton", background=[("active", "#ececec"), ("pressed", "#e3e3e3")])
        style.configure("SelectedTool.TButton", background=COLOR_TOOL_SELECTED, foreground="#ffffff", padding=(18, 8), bordercolor="#4a4a4a")
        style.map("SelectedTool.TButton", background=[("active", "#3a3a3a"), ("pressed", "#2d2d2d")])
        style.configure("TEntry", fieldbackground="#ffffff", padding=(8, 6), bordercolor=COLOR_BORDER)
        style.configure("TLabelframe", background=COLOR_BG, bordercolor=COLOR_BORDER, relief="solid")
        style.configure("Card.TLabelframe", background=COLOR_SURFACE, bordercolor=COLOR_BORDER, relief="solid")
        style.configure("Card.TLabelframe.Label", background=COLOR_SURFACE, foreground=COLOR_TEXT, font=("Microsoft YaHei UI", 11, "bold"))

    def _build_layout(self) -> None:
        top = ttk.Frame(self.root, style="Surface.TFrame", padding=(24, 18))
        top.pack(fill="x")
        title_group = ttk.Frame(top, style="Surface.TFrame")
        title_group.pack(side="left", fill="x", expand=True)
        ttk.Label(title_group, text=APP_NAME, style="HeroTitle.TLabel").pack(anchor="w")
        ttk.Label(
            title_group,
            text="轻量、清爽的临时需求处理工作台",
            style="HeroSubtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        actions = ttk.Frame(top, style="Surface.TFrame")
        actions.pack(side="right")
        ttk.Button(
            actions,
            text="工具列表",
            command=self.open_tool_list,
            style="Primary.TButton",
        ).pack(side="left", padx=(0, 10))
        ttk.Button(actions, text="密码本", command=self.open_password_book, style="Small.TButton").pack(side="left")

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
            inner = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
            window_id = outer.create_window((8, 8), window=inner, anchor="nw")
            rect_id = outer.create_rectangle(2, 2, 10, 10, dash=(5, 3), outline=COLOR_BORDER, width=2)

            def resize(event: tk.Event, canvas: tk.Canvas = outer, win: int = window_id, rect: int = rect_id) -> None:
                width = max(220, canvas.winfo_width() - 16)
                canvas.itemconfigure(win, width=width)
                needed = event.height + 16
                canvas.configure(height=needed)
                canvas.coords(rect, 2, 2, max(4, canvas.winfo_width() - 2), needed - 2)

            inner.bind("<Configure>", resize)
            outer.bind(
                "<Configure>",
                lambda e, c=outer, r=rect_id: c.coords(
                    r, 2, 2, max(4, e.width - 2), max(4, int(c.cget("height")) - 2)
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
                font=("Microsoft YaHei UI", 10, "bold"),
            ).pack(side="left")
            ttk.Button(header, text="改名", command=lambda c=category: self.rename_category(c), style="Small.TButton").pack(side="right")
            ttk.Button(header, text="删除", command=lambda c=category: self.delete_category(c), style="Danger.TButton").pack(side="right", padx=6)

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
            style_name = "SelectedTool.TButton" if key == self.current_tool_key else "Tool.TButton"
            btn = ttk.Button(body, text=f"  {tool.name}", command=lambda k=key: self.open_tool(k), style=style_name)
            btn.pack(fill="x", pady=4)
            btn.bind("<ButtonPress-1>", lambda _e, k=key: self.start_drag(k))
            btn.bind("<ButtonRelease-1>", self.finish_drag)

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
