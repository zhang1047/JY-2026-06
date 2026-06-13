from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from jy_toolbox.core.constants import *
from jy_toolbox.core.platform import enable_light_title_bar
from jy_toolbox.core.security import mask_password
from jy_toolbox.ui.widgets import make_rounded_button, rounded_rect_points

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
