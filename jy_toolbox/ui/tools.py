from __future__ import annotations

from pathlib import Path
from typing import Any, Callable
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from jy_toolbox.core.constants import *
from jy_toolbox.services.analytics import (
    calculate_active_day_ratio_excel,
    calculate_added_opinion_share_rate_excel,
    calculate_average_original_post_interactions_excel,
    calculate_average_daily_original_posts_excel,
    calculate_average_post_length_excel,
    calculate_daily_active_span_excel,
    calculate_post_type_ratios_excel,
    calculate_post_theme_ratios_excel,
    calculate_posting_period_type_excel,
    calculate_sensitive_topic_participation_rate_excel,
    calculate_sentiment_expression_excel,
    calculate_source_media_camp_ratios_excel,
    calculate_stance_tendency_excel,
    calculate_weekly_post_frequency_excel,
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
        self.sheet_combo = self.add_sheet_selector(form, 2, "工作表名：", self.sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始去重", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.populate_sheets_async(path, self.sheet_combo, self.sheet_var)
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


class PostingPeriodTypeTool(BaseToolFrame):
    OUTPUT_COLUMNS = ("高频发帖时段", "高频发帖类型")

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.dictionary_input_var = tk.StringVar(value=state.get("dictionary_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.dictionary_sheet_var = tk.StringVar(value=state.get("dictionary_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel、贴文 Excel 和发帖时段字典 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="高频发帖时段类型", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var)
        form.columnconfigure(1, weight=1)
        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        self.account_input_var.set(path)
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path); self.output_var.set(str(p.with_name(f"{p.stem}_高频发帖时段类型.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(title="选择贴文 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.post_input_var.set(path); self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var); self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择发帖时段字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path); self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var); self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存账号表处理结果", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path); self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state("posting_period_type", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "dictionary_input_path": self.dictionary_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
            "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
        })
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip(); post_input_path = self.post_input_var.get().strip(); dictionary_input_path = self.dictionary_input_var.get().strip(); output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0; post_sheet_name = self.post_sheet_var.get().strip() or 0; dictionary_sheet_name = self.dictionary_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self); return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self); return
        if not dictionary_input_path:
            messagebox.showwarning("提示", "请选择发帖时段字典 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self); return
        self.save_state()
        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_posting_period_type_excel(Path(account_input_path), Path(post_input_path), Path(dictionary_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, dictionary_sheet_name, progress)
        def on_success(result: dict[str, int]) -> None:
            self.status_var.set("完成：账号 {accounts} 行，贴文 {posts} 行，字典时段 {periods} 个，有效时间贴文 {valid_time_posts} 行，匹配时段贴文 {matched_posts} 行，已写入 {typed_accounts} 个账号，其中混乱型 {disorder_accounts} 个。输出：{output}".format(**result, output=output_path))
            messagebox.showinfo("完成", self.status_var.get(), parent=self)
        self.run_in_background(task, on_success, start_message="已开始后台统计高频发帖时段类型……")

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
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
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
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
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
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
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
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
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


class AverageOriginalPostInteractionsTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    CREATION_TYPE_COLUMN = "创作类型"
    LIKE_COLUMN = "点赞数"
    SHARE_COLUMN = "分享数"
    COMMENT_COLUMN = "评论数"
    OUTPUT_COLUMN = "平均原创单帖互动数（条）"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, CREATION_TYPE_COLUMN, LIKE_COLUMN, SHARE_COLUMN, COMMENT_COLUMN]

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
        form = ttk.LabelFrame(self, text="平均原创单帖互动数（条）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_平均原创单帖互动数.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
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
            "average_original_post_interactions",
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
            return calculate_average_original_post_interactions_excel(
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
                "完成：账号 {accounts} 行，贴文 {posts} 行，其中原创帖 {original_posts} 行，已匹配 {matched_accounts} 个账号，"
                "已写入 {averaged_accounts} 个账号。输出：{output}".format(
                    accounts=result["accounts"],
                    posts=result["posts"],
                    matched_accounts=result["matched_accounts"],
                    averaged_accounts=result["averaged_accounts"],
                    original_posts=result["original_posts"],
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计平均原创单帖互动数……")


class AverageDailyOriginalPostsTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    CREATION_TYPE_COLUMN = "创作类型"
    POST_TIME_COLUMN = "贴文发布时间"
    OUTPUT_COLUMN = "日均原创量（条）"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, CREATION_TYPE_COLUMN, POST_TIME_COLUMN]

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
        form = ttk.LabelFrame(self, text="日均原创量（条）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        self.account_input_var.set(path)
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_日均原创量.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(title="选择贴文 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
            self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存账号表处理结果", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state(
            "average_daily_original_posts",
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
            return calculate_average_daily_original_posts_excel(
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
                "完成：账号 {accounts} 行，贴文 {posts} 行，有有效发布时间的贴文 {valid_time_posts} 行，"
                "其中原创帖 {original_posts} 行，已匹配 {matched_accounts} 个账号，"
                "已写入 {averaged_accounts} 个账号。输出：{output}".format(**result, output=output_path)
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计日均原创量……")


class WeeklyPostFrequencyTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    POST_TIME_COLUMN = "贴文发布时间"
    OUTPUT_COLUMN = "每周发布帖子频率（次）"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, POST_TIME_COLUMN]

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
        form = ttk.LabelFrame(self, text="每周发布帖子频率（次）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        self.account_input_var.set(path)
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_每周发布帖子频率.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(title="选择贴文 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
            self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存账号表处理结果", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state("weekly_post_frequency", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
        })
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip()
        post_input_path = self.post_input_var.get().strip()
        output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0
        post_sheet_name = self.post_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self); return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self); return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_weekly_post_frequency_excel(Path(account_input_path), Path(post_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, progress)

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set(
                "完成：账号 {accounts} 行，贴文 {posts} 行，有有效发布时间的贴文 {valid_time_posts} 行，"
                "已匹配 {matched_accounts} 个账号，已写入 {calculated_accounts} 个账号。输出：{output}".format(**result, output=output_path)
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计每周发布帖子频率……")


class DailyActiveSpanTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    POST_TIME_COLUMN = "贴文发布时间"
    OUTPUT_COLUMN = "日均在线活跃时段跨度（小时/天）"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, POST_TIME_COLUMN]

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
        form = ttk.LabelFrame(self, text="日均在线活跃时段跨度（小时/天）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=2, padx=10, pady=8)

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        self.account_input_var.set(path)
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_日均在线活跃时段跨度.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(title="选择贴文 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
            self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存账号表处理结果", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state(
            "daily_active_span",
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
            return calculate_daily_active_span_excel(
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
                "完成：账号 {accounts} 行，贴文 {posts} 行，有有效发布时间的贴文 {valid_time_posts} 行，"
                "已匹配 {matched_accounts} 个账号，符合单日 2 条及以上的日期 {qualified_days} 个，"
                "已写入 {spanned_accounts} 个账号。输出：{output}".format(**result, output=output_path)
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计日均在线活跃时段跨度……")


class ActiveDayRatioTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    POST_TIME_COLUMN = "贴文发布时间"
    POST_COUNT_COLUMN = "帖子数量"
    POST_SPAN_DAYS_COLUMN = "帖子时间跨度天数"
    ACTIVE_DAYS_COLUMN = "活跃天数"
    OUTPUT_COLUMN = "活跃天数占比"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, POST_TIME_COLUMN]

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
        form = ttk.LabelFrame(self, text="活跃天数占比（%）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_活跃天数占比.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
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
            "active_day_ratio",
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
            return calculate_active_day_ratio_excel(
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
                "完成：账号 {accounts} 行，贴文 {posts} 行，有有效发布时间的贴文 {valid_time_posts} 行，"
                "已匹配 {matched_accounts} 个账号，已写入帖子数量 {post_count_accounts} 个账号，"
                "已写入帖子时间跨度天数 {active_span_accounts} 个账号，"
                "已写入活跃天数 {active_days_accounts} 个账号，"
                "已写入活跃天数占比 {active_ratio_accounts} 个账号。输出：{output}".format(
                    **result,
                    output=output_path,
                )
            )
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计活跃天数占比……")

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
        self.account_sheet_combo = self.add_sheet_selector(form, 3, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 4, "贴文表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
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
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
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
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var)
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
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
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
            self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择账号名字典 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.dictionary_input_var.set(path)
            self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var)
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


class PostThemeRatioTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    BODY_COLUMN = "帖子正文"
    THEME_COLUMN = "内容偏好"
    OUTPUT_PREFIX = "主题占比-"
    REQUIRED_POST_COLUMNS = [POST_URL_COLUMN, BODY_COLUMN]
    REQUIRED_DICTIONARY_COLUMNS = [BODY_COLUMN, THEME_COLUMN]

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.dictionary_input_var = tk.StringVar(value=state.get("dictionary_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.dictionary_sheet_var = tk.StringVar(value=state.get("dictionary_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel、贴文 Excel 和内容偏好字典 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="帖子主题占比（%）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var, hint="选择含第一列“帖子正文”和“内容偏好”列的 sheet")
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_帖子主题占比.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)
            self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择内容偏好字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path)
            self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var)
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
        self.app.config.set_tool_state("post_theme_ratio", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "dictionary_input_path": self.dictionary_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
            "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
        })
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
            messagebox.showwarning("提示", "请选择内容偏好字典 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_post_theme_ratios_excel(Path(account_input_path), Path(post_input_path), Path(dictionary_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, dictionary_sheet_name, progress)

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set("完成：账号 {accounts} 行，贴文 {posts} 行，字典 {dictionary_rows} 行，识别主题 {themes} 个，匹配贴文 {matched_posts} 行，已写入 {themed_accounts} 个账号。输出：{output}".format(**result, output=output_path))
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计帖子主题占比……")


class SensitiveTopicParticipationRateTool(PostThemeRatioTool):
    OUTPUT_COLUMN = "敏感话题参与率（%）"

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.status_var.set("请选择账号 Excel、贴文 Excel 和敏感话题关键词字典 Excel 后开始统计。")

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="敏感话题参与率（%）", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "关键词字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var, hint="选择第一列为关键词、无标题的 sheet")
        form.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        self.account_input_var.set(path)
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path)
            self.output_var.set(str(p.with_name(f"{p.stem}_敏感话题参与率.xlsx")))
        self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择敏感话题关键词字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path)
            self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var)
            self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state("sensitive_topic_participation_rate", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "dictionary_input_path": self.dictionary_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
            "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
        })
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
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self); return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self); return
        if not dictionary_input_path:
            messagebox.showwarning("提示", "请选择敏感话题关键词字典 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self); return
        self.save_state()

        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_sensitive_topic_participation_rate_excel(Path(account_input_path), Path(post_input_path), Path(dictionary_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, dictionary_sheet_name, progress)

        def on_success(result: dict[str, int]) -> None:
            self.status_var.set("完成：账号 {accounts} 行，贴文 {posts} 行，关键词 {dictionary_keywords} 个，匹配账号 {matched_accounts} 个，敏感话题贴文 {sensitive_posts} 行，已写入 {rated_accounts} 个账号。输出：{output}".format(**result, output=output_path))
            messagebox.showinfo("完成", self.status_var.get(), parent=self)

        self.run_in_background(task, on_success, start_message="已开始后台统计敏感话题参与率……")


class SentimentExpressionTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    BODY_COLUMN = "帖子正文"
    SENTIMENT_COLUMN = "情感表达倾向"
    OUTPUT_COLUMNS = (
        "情感表达-分数",
        "情感表达-正面（数量）",
        "情感表达-正面（占比）",
        "情感表达-负面（数量）",
        "情感表达-负面（占比）",
        "情感表达-中性（数量）",
        "情感表达-中性（占比）",
    )

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.dictionary_input_var = tk.StringVar(value=state.get("dictionary_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.dictionary_sheet_var = tk.StringVar(value=state.get("dictionary_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel、贴文 Excel 和情感表达字典 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="情感表达分数&数量占比", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var)
        form.columnconfigure(1, weight=1)
        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path); self.output_var.set(str(p.with_name(f"{p.stem}_情感表达统计.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path); self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var); self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择情感表达字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path); self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var); self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存账号表处理结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_var.set(path); self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state("sentiment_expression", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "dictionary_input_path": self.dictionary_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
            "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
        })
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip(); post_input_path = self.post_input_var.get().strip(); dictionary_input_path = self.dictionary_input_var.get().strip(); output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0; post_sheet_name = self.post_sheet_var.get().strip() or 0; dictionary_sheet_name = self.dictionary_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self)
            return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self)
            return
        if not dictionary_input_path:
            messagebox.showwarning("提示", "请选择情感表达字典 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()
        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_sentiment_expression_excel(Path(account_input_path), Path(post_input_path), Path(dictionary_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, dictionary_sheet_name, progress)
        def on_success(result: dict[str, int]) -> None:
            self.status_var.set("完成：账号 {accounts} 行，贴文 {posts} 行，字典 {dictionary_rows} 行，匹配贴文 {matched_posts} 行，已写入 {sentiment_accounts} 个账号。输出：{output}".format(**result, output=output_path))
            messagebox.showinfo("完成", self.status_var.get(), parent=self)
        self.run_in_background(task, on_success, start_message="已开始后台统计情感表达分数和数量占比……")


class StanceTendencyTool(BaseToolFrame):
    ACCOUNT_URL_COLUMN = "FB主页"
    POST_URL_COLUMN = "主页url"
    BODY_COLUMN = "帖子正文"
    STANCE_COLUMN = "两岸议题立场倾向"
    OUTPUT_COLUMNS = (
        "立场倾向-分数",
        "立场倾向-偏蓝（数量）",
        "立场倾向-偏蓝（占比）",
        "立场倾向-偏绿（数量）",
        "立场倾向-偏绿（占比）",
        "立场倾向-中立（数量）",
        "立场倾向-中立（占比）",
    )

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", state: dict[str, Any], description: str) -> None:
        super().__init__(parent, app, state, description)
        self.account_input_var = tk.StringVar(value=state.get("account_input_path", ""))
        self.post_input_var = tk.StringVar(value=state.get("post_input_path", ""))
        self.dictionary_input_var = tk.StringVar(value=state.get("dictionary_input_path", ""))
        self.output_var = tk.StringVar(value=state.get("output_path", ""))
        self.account_sheet_var = tk.StringVar(value=state.get("account_sheet_name", ""))
        self.post_sheet_var = tk.StringVar(value=state.get("post_sheet_name", ""))
        self.dictionary_sheet_var = tk.StringVar(value=state.get("dictionary_sheet_name", ""))
        self.status_var = tk.StringVar(value="请选择账号 Excel、贴文 Excel 和立场倾向字典 Excel 后开始统计。")
        self._build_form()

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text="立场倾向分数&数量占比", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "贴文 Excel：", self.post_input_var, self.choose_post_input)
        self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "贴文表工作表：", self.post_sheet_var)
        self.dictionary_sheet_combo = self.add_sheet_selector(form, 6, "字典表工作表：", self.dictionary_sheet_var)
        form.columnconfigure(1, weight=1)
        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "开始统计", self.run, role="primary", width=82).pack(side="left")
        make_rounded_button(actions, "保存当前填写", self.save_state, width=98).pack(side="left", padx=10)
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")
        self.add_progress_bar(status_card)
        self.load_configured_sheets_async()

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
        self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)
        if not self.output_var.get().strip():
            p = Path(path); self.output_var.set(str(p.with_name(f"{p.stem}_立场倾向统计.xlsx")))
        self.save_state()

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择贴文 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.post_input_var.set(path); self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var); self.save_state()

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择立场倾向字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path); self.populate_sheets_async(path, self.dictionary_sheet_combo, self.dictionary_sheet_var); self.save_state()

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="保存账号表处理结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_var.set(path); self.save_state()

    def save_state(self) -> None:
        self.app.config.set_tool_state("stance_tendency", {
            "account_input_path": self.account_input_var.get().strip(),
            "post_input_path": self.post_input_var.get().strip(),
            "dictionary_input_path": self.dictionary_input_var.get().strip(),
            "output_path": self.output_var.get().strip(),
            "account_sheet_name": self.account_sheet_var.get().strip(),
            "post_sheet_name": self.post_sheet_var.get().strip(),
            "dictionary_sheet_name": self.dictionary_sheet_var.get().strip(),
        })
        self.status_var.set("当前工具填写内容已保存。")

    def run(self) -> None:
        account_input_path = self.account_input_var.get().strip(); post_input_path = self.post_input_var.get().strip(); dictionary_input_path = self.dictionary_input_var.get().strip(); output_path = self.output_var.get().strip()
        account_sheet_name = self.account_sheet_var.get().strip() or 0; post_sheet_name = self.post_sheet_var.get().strip() or 0; dictionary_sheet_name = self.dictionary_sheet_var.get().strip() or 0
        if not account_input_path:
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self)
            return
        if not post_input_path:
            messagebox.showwarning("提示", "请选择贴文 Excel。", parent=self)
            return
        if not dictionary_input_path:
            messagebox.showwarning("提示", "请选择立场倾向字典 Excel。", parent=self); return
        if not output_path:
            messagebox.showwarning("提示", "请选择输出 Excel。", parent=self)
            return
        self.save_state()
        def task(progress: Callable[[float, str | None], None]) -> dict[str, int]:
            return calculate_stance_tendency_excel(Path(account_input_path), Path(post_input_path), Path(dictionary_input_path), Path(output_path), self.app.config.data.get("passwords", []), account_sheet_name, post_sheet_name, dictionary_sheet_name, progress)
        def on_success(result: dict[str, int]) -> None:
            self.status_var.set("完成：账号 {accounts} 行，贴文 {posts} 行，字典 {dictionary_rows} 行，匹配贴文 {matched_posts} 行，已写入 {stance_accounts} 个账号。输出：{output}".format(**result, output=output_path))
            messagebox.showinfo("完成", self.status_var.get(), parent=self)
        self.run_in_background(task, on_success, start_message="已开始后台统计立场倾向分数和数量占比……")
