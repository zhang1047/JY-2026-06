from __future__ import annotations

from pathlib import Path
from typing import Any, Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from jy_toolbox.core.constants import *
from jy_toolbox.services.analytics import (
    calculate_added_opinion_share_rate_excel,
    calculate_average_post_length_excel,
    calculate_post_type_ratios_excel,
    calculate_source_media_camp_ratios_excel,
    deduplicate_posts_excel,
)
from jy_toolbox.ui.base import BaseToolFrame
from jy_toolbox.ui.widgets import make_rounded_button

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

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
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

class PostTypeRatioTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    LEGACY_OUTPUT_COLUMN = "帖子类型"
    OUTPUT_COLUMNS = ("文字帖占比", "图片帖占比", "视频帖占比")
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

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
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
                "有可判断类型的账号 {typed_accounts} 个。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    matched_accounts=result["matched_accounts"],
                    typed_accounts=result["typed_accounts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计贴文类型占比……")

class AveragePostLengthTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    BODY_COLUMN = "帖子正文"
    OUTPUT_COLUMN = "平均发帖长度"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, BODY_COLUMN]

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
        form = ttk.LabelFrame(self, text="平均发帖长度", style="Card.TLabelframe", padding=(12, 9))
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

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
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
            self.output_var.set(str(p.with_name(f"{p.stem}_平均发帖长度.xlsx")))
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
            "average_post_length",
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
            return calculate_average_post_length_excel(
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
                "已写入 {averaged_accounts} 个账号。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    matched_accounts=result["matched_accounts"],
                    averaged_accounts=result["averaged_accounts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计平均发帖长度……")

class AddedOpinionShareRateTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    CREATION_TYPE_COLUMN = "创作类型"
    TITLE_COLUMN = "标题"
    BODY_COLUMN = "帖子正文"
    OUTPUT_COLUMN = "附加观点转发率"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, CREATION_TYPE_COLUMN, TITLE_COLUMN, BODY_COLUMN]

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
        form = ttk.LabelFrame(self, text="附加观点转发率（%）", style="Card.TLabelframe", padding=(12, 9))
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

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
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
            self.output_var.set(str(p.with_name(f"{p.stem}_附加观点转发率.xlsx")))
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
            "added_opinion_share_rate",
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
            return calculate_added_opinion_share_rate_excel(
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
                "完成：账号 {accounts} 行，贴文 {posts} 行，其中转发贴 {share_posts} 行，"
                "已匹配 {matched_accounts} 个账号，已写入 {rated_accounts} 个账号。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    share_posts=result["share_posts"],
                    matched_accounts=result["matched_accounts"],
                    rated_accounts=result["rated_accounts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计附加观点转发率……")

class SourceMediaCampRatioTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    CREATION_TYPE_COLUMN = "创作类型"
    SHARED_ACCOUNT_COLUMN = "分享贴账号名"
    STANCE_COLUMN = "账号立场归属"
    ACCOUNT_TYPE_COLUMN = "账号类型归属"
    OFFICIAL_SOURCE_TYPE = "政府/军警机关"
    OUTPUT_COLUMNS = ("账号立场归属占比", "账号类型归属占比", "官方信源占比")
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, CREATION_TYPE_COLUMN, SHARED_ACCOUNT_COLUMN]
    REQUIRED_DICT_COLUMNS = [SHARED_ACCOUNT_COLUMN, STANCE_COLUMN, ACCOUNT_TYPE_COLUMN]

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.dictionary_input_var = tk.StringVar(value=state.get("dictionary_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.dictionary_sheet_var = tk.StringVar(value=state.get("dictionary_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel、贴文 Excel 和账号名字典 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(
            self,
            text="信息来源的媒体阵营分布（%）",
            style="Card.TLabelframe",
            padding=(12, 9),
        )
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        ttk.Label(form, text="账号表工作表：").grid(row=4, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.account_sheet_var).grid(row=4, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=4, column=2, sticky="w", padx=10, pady=8)
        ttk.Label(form, text="贴文表工作表：").grid(row=5, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.post_sheet_var).grid(row=5, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="留空则读取第一个工作表").grid(row=5, column=2, sticky="w", padx=10, pady=8)
        ttk.Label(form, text="字典表工作表：").grid(row=6, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(form, textvariable=self.dictionary_sheet_var).grid(row=6, column=1, sticky="ew", padx=10, pady=8)
        ttk.Label(form, text="请填写字典 Excel 的 sheet 名；留空则读取第一个工作表").grid(
            row=6,
            column=2,
            sticky="w",
            padx=10,
            pady=8,
        )
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)

    def _path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        var: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
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
            self.output_var.set(str(p.with_name(f"{p.stem}_媒体阵营分布.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path)
            self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择账号名字典 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.dictionary_input_var.set(path)
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
            "source_media_camp_ratio",
            {
                "account_input_path": self.account_input_var.get().strip(),
                "post_input_path": self.post_input_var.get().strip(),
                "dictionary_input_path": self.dictionary_input_var.get().strip(),
                "output_path": self.output_var.get().strip(),
                "account_sheet_name": self.account_sheet_var.get().strip(),
                "post_sheet_name": self.post_sheet_var.get().strip(),
                "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
            },
        )
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip()
        post_input_path = self.post_input_var.get().strip()
        dictionary_input_path = self.dictionary_input_var.get().strip()
        output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0
        post_sheet_name = self.post_sheet_var.get().strip() or 0
        dictionary_sheet_name = self.dictionary_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self)
            return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self)
            return
        if not dictionary_input_path:
            messagebox.showwarning("提示", "请选择账号名字典 Excel。", parent=self)
            return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_source_media_camp_ratios_excel(
                Path(account_input_path),
                Path(post_input_path),
                Path(dictionary_input_path),
                Path(output_path),
                self.app.config.data.get("passwords", []),
                account_sheet_name,
                post_sheet_name,
                dictionary_sheet_name,
                progress,
            )

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set(
                "完成：账号 {accounts} 行，贴文 {posts} 行，其中分享贴 {share_posts} 行，"
                "字典匹配分享贴 {matched_share_posts} 行，已写入 {classified_accounts} 个账号。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    share_posts=result["share_posts"],
                    matched_share_posts=result["matched_share_posts"],
                    classified_accounts=result["classified_accounts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计信息来源的媒体阵营分布……")
