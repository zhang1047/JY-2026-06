from __future__ import annotations

from pathlib import Path
from typing import Any, Callable
import os
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from jy_toolbox.core.config import ConfigStore
from jy_toolbox.core.constants import *
from jy_toolbox.core.models import ToolDefinition
from jy_toolbox.core.platform import enable_light_title_bar
from jy_toolbox.ui.dialogs import PasswordBookDialog, ToolListDialog
from jy_toolbox.ui.tools import (
    ActiveDayRatioTool,
    AddedOpinionShareRateTool,
    AverageDailyOriginalPostsTool,
    AverageOriginalPostInteractionsTool,
    AveragePostLengthTool,
    CustomKeywordFrequencyTool,
    DailyActiveSpanTool,
    PostDedupTool,
    PostingPeriodTypeTool,
    PostTypeRatioTool,
    PostThemeRatioTool,
    SensitiveTopicParticipationRateTool,
    SentimentExpressionTool,
    SourceMediaCampRatioTool,
    StanceTendencyTool,
    WeeklyPostFrequencyTool,
)
from jy_toolbox.ui.base import BaseToolFrame
from jy_toolbox.ui.widgets import make_rounded_button, rounded_rect_points

class ToolboxApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1080x720")
        self.root.minsize(1080, 720)
        enable_light_title_bar(self.root)
        self.config = ConfigStore()
        self.tools: dict[str, ToolDefinition] = {}
        self.current_tool_frame: BaseToolFrame | None = None
        self.current_tool_key: str | None = None
        self.drag_data: dict[str, Any] = {}
        self.drag_threshold = 6
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
                name="帖去重",
                default_category="Excel 工具",
                description=(
                    "说明：根据“帖url”列去重。若同一 URL 有重复行，会比较“点赞数”“分享数”“评论数”三列的数值总和，"
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
                name="帖类型占比（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联，"
                    "按规则统计每个账号文字、图片、视频帖占比，并在账号表最后新增三列占比。"
                ),
                factory=lambda parent, app, state: PostTypeRatioTool(
                    parent, app, state, app.get_tool_description("post_type_ratio")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="posting_period_type",
                name="高频发帖时段类型",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和发帖时段字典 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "按字典中的“时段类型 / 起始时段 / 结束时段”动态判断每条帖发布时间所属时段，"
                    "再按账号写回“高频发帖时段”和“高频发帖类型”。若账号在所有字典时段均有发帖且各时段占比差值不高于 10%，"
                    "则写为混乱型且不写回具体时段。"
                ),
                factory=lambda parent, app, state: PostingPeriodTypeTool(
                    parent, app, state, app.get_tool_description("posting_period_type")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="post_theme_ratio",
                name="帖主题占比（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和内容偏好字典 Excel，通过账号表“FB主页”与帖表“主页url”关联，"
                    "再用帖表“帖正文”匹配字典 sheet 第一列“帖正文”的“内容偏好”分组，"
                    "按账号计算各内容偏好分组占比，并在账号表最后新增“主题占比-分组名”列。"
                ),
                factory=lambda parent, app, state: PostThemeRatioTool(
                    parent, app, state, app.get_tool_description("post_theme_ratio")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="sensitive_topic_participation_rate",
                name="敏感话题参与率（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和敏感话题关键词字典 Excel，通过账号表“FB主页”与帖表“主页url”关联，"
                    "用字典 sheet 第一列（无标题）关键词快速扫描帖“标题”和“帖正文”，"
                    "按账号计算包含任一关键词的帖数占该账号全部帖数的百分比，并在账号表最后新增“敏感话题参与率（%）”列。"
                ),
                factory=lambda parent, app, state: SensitiveTopicParticipationRateTool(
                    parent, app, state, app.get_tool_description("sensitive_topic_participation_rate")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="custom_keyword_frequency",
                name="自定义关键词统计",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "在工具界面新增多个关键词后，按每个关键词统计每个账号所有帖标题和正文中的出现次数，"
                    "并在账号表最后新增“词频-关键词”列。关键词列表支持增删，并可一键转换为繁体或简体。"
                ),
                factory=lambda parent, app, state: CustomKeywordFrequencyTool(
                    parent, app, state, app.get_tool_description("custom_keyword_frequency")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="sentiment_expression",
                name="情感表达分数&数量占比",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和情感表达字典 Excel，通过账号表“FB主页”与帖表“主页url”关联，"
                    "再用帖表“帖正文”匹配字典 sheet 第一列“帖正文”的“情感表达倾向”枚举，"
                    "按账号计算情感表达分数（正面 +1、负面 -1、中性 0）以及正面、负面、中性的数量和占比。"
                ),
                factory=lambda parent, app, state: SentimentExpressionTool(
                    parent, app, state, app.get_tool_description("sentiment_expression")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="stance_tendency",
                name="立场倾向分数&数量占比",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和立场倾向字典 Excel，通过账号表“FB主页”与帖表“主页url”关联，"
                    "再用帖表“帖正文”匹配字典 sheet 第一列“帖正文”的“两岸议题立场倾向”枚举，"
                    "按账号计算立场倾向分数（偏蓝 +1、中立 0、偏绿 -1）以及偏蓝、偏绿、中立的数量和占比。"
                ),
                factory=lambda parent, app, state: StanceTendencyTool(
                    parent, app, state, app.get_tool_description("stance_tendency")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="daily_active_span",
                name="日均在线活跃时段跨度（小时/天）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "仅统计每个账号单日发帖 2 条及以上的日期，先按日期计算当天最早到最晚发帖时间间隔（小时），"
                    "再对这些日期的间隔取平均值，并在账号表最后新增“日均在线活跃时段跨度（小时/天）”列。"
                ),
                factory=lambda parent, app, state: DailyActiveSpanTool(
                    parent, app, state, app.get_tool_description("daily_active_span")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="active_day_ratio",
                name="活跃天数占比（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "按每个账号最早到最晚的帖发布时间计算账号发帖时间范围天数，"
                    "再用该账号实际发帖日期数除以时间范围天数，并在账号表最后依次新增“帖数量”、"
                    "“帖时间跨度天数”、“活跃天数”和“活跃天数占比”列。"
                ),
                factory=lambda parent, app, state: ActiveDayRatioTool(
                    parent, app, state, app.get_tool_description("active_day_ratio")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="average_post_length",
                name="平均发帖长度",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "只统计帖表“帖正文”列的文字长度，按账号计算平均每个帖的正文长度，"
                    "并在账号表最后新增“平均发帖长度”列。"
                ),
                factory=lambda parent, app, state: AveragePostLengthTool(
                    parent, app, state, app.get_tool_description("average_post_length")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="average_daily_original_posts",
                name="日均原创量（条）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "按每个账号最早到最晚的帖发布时间计算统计时间范围天数，"
                    "仅排除帖表“创作类型”为 share 的转发帖后统计原创帖数量，"
                    "用原创帖数量除以统计时间范围天数，并在账号表最后新增“日均原创量（条）”列。"
                ),
                factory=lambda parent, app, state: AverageDailyOriginalPostsTool(
                    parent, app, state, app.get_tool_description("average_daily_original_posts")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="weekly_post_frequency",
                name="每周发布帖频率（次）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "按每个账号最早到最晚的帖发布日期计算统计自然周数（含首尾日期，向上取整且最少 1 周），"
                    "先在账号表最后新增“跨越周数”列，再新增“每周发布帖频率（次）”列。"
                ),
                factory=lambda parent, app, state: WeeklyPostFrequencyTool(
                    parent, app, state, app.get_tool_description("weekly_post_frequency")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="average_original_post_interactions",
                name="平均原创单帖互动数（条）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "仅统计帖表“创作类型”为 common 的原创帖，将每帖“点赞数”“评论数”“分享数”相加得到互动数，"
                    "按账号计算平均原创单帖互动数，并在账号表最后新增“平均原创单帖互动数（条）”列。"
                ),
                factory=lambda parent, app, state: AverageOriginalPostInteractionsTool(
                    parent, app, state, app.get_tool_description("average_original_post_interactions")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="added_opinion_share_rate",
                name="附加观点转发率（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和帖 Excel，通过账号表“FB主页”与帖表“主页url”关联；"
                    "仅统计“创作类型”为 share 的转发贴，其中“标题”或“帖正文”任一不为空即视为附加观点，"
                    "按账号计算附加观点转发率，并在账号表最后新增“附加观点转发率”列。"
                ),
                factory=lambda parent, app, state: AddedOpinionShareRateTool(
                    parent, app, state, app.get_tool_description("added_opinion_share_rate")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="source_media_camp_ratio",
                name="信息来源的媒体阵营分布（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、帖 Excel 和账号名字典 Excel；仅统计帖表中“创作类型”为 share 的转发贴，"
                    "用“分享贴账号名”匹配字典中的“账号立场归属”和“账号类型归属”，"
                    "按账号汇总各分组占比，并在账号表最后新增两列占比。"
                ),
                factory=lambda parent, app, state: SourceMediaCampRatioTool(
                    parent, app, state, app.get_tool_description("source_media_camp_ratio")
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
        self._normalize_tool_orders()
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

    def refresh_current_excel_sheets(self) -> None:
        """密码本更新后，重新识别当前工具里已选择的 Excel 工作表。"""
        if self.current_tool_frame is not None:
            self.current_tool_frame.load_configured_sheets_async(show_errors=True)

    def open_tool_list(self) -> None:
        if self.tool_list_dialog is not None and self.tool_list_dialog.winfo_exists():
            self.tool_list_dialog.lift()
            self.tool_list_dialog.focus_force()
            return
        self.tool_list_dialog = ToolListDialog(self)
        self.refresh_tool_list()


    def _normalize_tool_orders(self) -> None:
        categories = self.config.data.setdefault("categories", [])
        tool_categories = self.config.data.setdefault("tool_categories", {})
        orders = self.config.data.setdefault("tool_orders", {})
        for category in categories:
            existing_order = [key for key in orders.get(category, []) if key in self.tools and tool_categories.get(key) == category]
            missing_keys = [key for key, tool in self.tools.items() if tool_categories.get(key, tool.default_category) == category and key not in existing_order]
            orders[category] = existing_order + missing_keys
        for category in list(orders):
            if category not in categories:
                orders.pop(category, None)

    def tools_in_category(self, category: str) -> list[str]:
        self._normalize_tool_orders()
        return list(self.config.data.setdefault("tool_orders", {}).get(category, []))

    def move_tool_order(self, tool_key: str, direction: int) -> None:
        category = self.config.data.setdefault("tool_categories", {}).get(tool_key, self.tools[tool_key].default_category)
        order = self.tools_in_category(category)
        if tool_key not in order:
            return
        index = order.index(tool_key)
        new_index = index + direction
        if new_index < 0 or new_index >= len(order):
            return
        order[index], order[new_index] = order[new_index], order[index]
        self.config.data.setdefault("tool_orders", {})[category] = order
        self.config.save()
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
            make_rounded_button(header, "删除", lambda c=category: self.delete_category(c), role="danger", width=48, height=24).pack(side="right")
            make_rounded_button(header, "改名", lambda c=category: self.rename_category(c), width=48, height=24).pack(side="right", padx=6)
            make_rounded_button(header, "一键执行", lambda c=category: self.open_group_run(c), role="primary", width=72, height=24).pack(side="right")

            body = ttk.Frame(inner, style="Card.TFrame")
            body.pack(fill="x")
            self.category_body_frames[category] = body
            self.category_drop_widgets[category].add(body)
            body.bind("<ButtonRelease-1>", lambda _e, c=category: self.drop_tool_to_category(c))

        self._normalize_tool_orders()
        for category in categories:
            body = self.category_body_frames.get(category)
            if body is None:
                continue
            order = self.tools_in_category(category)
            for index, key in enumerate(order):
                tool = self.tools.get(key)
                if tool is None:
                    continue
                row = ttk.Frame(body, style="Card.TFrame")
                row.pack(fill="x", pady=4)
                role = "selected" if key == self.current_tool_key else "normal"
                btn = make_rounded_button(row, tool.name, lambda k=key: self.open_tool(k), role=role, height=26)
                btn.pack(side="left", fill="x", expand=True)
                btn.bind("<ButtonPress-1>", lambda e, k=key: self.start_drag(k, e), add="+")
                btn.bind("<B1-Motion>", self.update_drag, add="+")
                btn.bind("<ButtonRelease-1>", self.finish_drag, add="+")
                up = make_rounded_button(row, "↑", lambda k=key: self.move_tool_order(k, -1), width=28, height=24)
                up.pack(side="left", padx=(6, 2))
                down = make_rounded_button(row, "↓", lambda k=key: self.move_tool_order(k, 1), width=28, height=24)
                down.pack(side="left")
                if index == 0:
                    up.configure(state="disabled")
                if index == len(order) - 1:
                    down.configure(state="disabled")

    def start_drag(self, tool_key: str, event: tk.Event) -> None:
        self.drag_data = {
            "tool_key": tool_key,
            "start_x_root": event.x_root,
            "start_y_root": event.y_root,
            "dragging": False,
        }

    def update_drag(self, event: tk.Event) -> None:
        tool_key = self.drag_data.get("tool_key")
        if not tool_key:
            return
        distance_x = abs(event.x_root - int(self.drag_data.get("start_x_root", event.x_root)))
        distance_y = abs(event.y_root - int(self.drag_data.get("start_y_root", event.y_root)))
        if distance_x >= self.drag_threshold or distance_y >= self.drag_threshold:
            self.drag_data["dragging"] = True

    def finish_drag(self, event: tk.Event) -> None:
        tool_key = self.drag_data.get("tool_key")
        if not tool_key:
            return
        if not self.drag_data.get("dragging"):
            self.drag_data = {}
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
        if tool_key and self.drag_data.get("dragging"):
            self.move_tool(tool_key, category)
            self.drag_data = {}

    def move_tool(self, tool_key: str, category: str) -> None:
        tool_categories = self.config.data.setdefault("tool_categories", {})
        old_category = tool_categories.get(tool_key, self.tools[tool_key].default_category)
        tool_categories[tool_key] = category
        orders = self.config.data.setdefault("tool_orders", {})
        for order_category in {old_category, category}:
            orders[order_category] = [key for key in orders.get(order_category, []) if key != tool_key]
        orders.setdefault(category, []).append(tool_key)
        self._normalize_tool_orders()
        self.config.save()
        self.refresh_tool_list()

    def add_category(self) -> None:
        name = simpledialog.askstring("新增分组", "请输入分组名称：", parent=self.root)
        if not name:
            return
        name = name.strip()
        categories = self.config.data.setdefault("categories", [])
        if not name or name in categories:
            messagebox.showwarning("提示", "分组名称不能为空或重复。", parent=self.root)
            return
        categories.append(name)
        self.config.save()
        self.refresh_tool_list()

    def rename_category(self, old_name: str) -> None:
        new_name = simpledialog.askstring("分组改名", "请输入新的分组名称：", initialvalue=old_name, parent=self.root)
        if not new_name:
            return
        new_name = new_name.strip()
        categories = self.config.data.setdefault("categories", [])
        if not new_name or (new_name in categories and new_name != old_name):
            messagebox.showwarning("提示", "分组名称不能为空或重复。", parent=self.root)
            return
        self.config.data["categories"] = [new_name if c == old_name else c for c in categories]
        orders = self.config.data.setdefault("tool_orders", {})
        if old_name in orders:
            orders[new_name] = orders.pop(old_name)
        for key, category in list(self.config.data.setdefault("tool_categories", {}).items()):
            if category == old_name:
                self.config.data["tool_categories"][key] = new_name
        self.config.save()
        self.refresh_tool_list()

    def delete_category(self, category: str) -> None:
        categories = self.config.data.setdefault("categories", [])
        if len(categories) <= 1:
            messagebox.showwarning("提示", "至少保留一个分组。", parent=self.root)
            return
        target = next((c for c in categories if c != category), None)
        if not messagebox.askyesno("删除分组", f"确认删除分组“{category}”？其中工具将移动到“{target}”。", parent=self.root):
            return
        self.config.data["categories"] = [c for c in categories if c != category]
        orders = self.config.data.setdefault("tool_orders", {})
        moved_order = orders.pop(category, [])
        orders.setdefault(target, []).extend([key for key in moved_order if key not in orders.get(target, [])])
        for key, assigned in list(self.config.data.setdefault("tool_categories", {}).items()):
            if assigned == category:
                self.config.data["tool_categories"][key] = target
        self._normalize_tool_orders()
        self.config.save()
        self.refresh_tool_list()

    def open_group_run(self, category: str) -> None:
        self.drag_data = {}
        if self.current_tool_frame is not None:
            try:
                self.current_tool_frame.save_description()
                self.current_tool_frame.save_state()
            except Exception:
                pass
            self.current_tool_frame.destroy()
        tool_keys = self.tools_in_category(category)
        self.current_tool_key = None
        self.current_tool_frame = GroupRunFrame(self.content, self, category, tool_keys)
        self.current_tool_frame.pack(fill="both", expand=True)
        self.refresh_tool_list()

    def open_tool(self, key: str) -> None:
        self.drag_data = {}
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

class GroupRunFrame(BaseToolFrame):
    """按分组顺序一键执行多个工具。"""

    def __init__(self, parent: tk.Widget, app: "ToolboxApp", category: str, tool_keys: list[str]) -> None:
        self.category = category
        self.tool_keys = tool_keys
        super().__init__(parent, app, {}, f"一键执行“{category}”分组中的 {len(tool_keys)} 个工具。账号表、帖表、字典表共用；每个工具可单独选择字典 sheet，并沿用该工具已保存的其他设置。")
        self.account_input_var = tk.StringVar(value="")
        self.post_input_var = tk.StringVar(value="")
        self.dictionary_input_var = tk.StringVar(value="")
        self.output_var = tk.StringVar(value="")
        self.account_sheet_var = tk.StringVar(value="")
        self.post_sheet_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="请选择共用表格后开始一键执行。")
        self.progress_text_var = tk.StringVar(value=f"0/{len(tool_keys)}")
        self.progress_var = tk.DoubleVar(value=0)
        self.dictionary_sheet_vars: dict[str, tk.StringVar] = {}
        self.dictionary_sheet_combos: dict[str, ttk.Combobox] = {}
        self.custom_keyword_text_vars: dict[str, tk.StringVar] = {}
        self.custom_keyword_input_vars: dict[str, tk.StringVar] = {}
        self.custom_ignore_script_vars: dict[str, tk.BooleanVar] = {}
        self.dictionary_loading_var = tk.StringVar(value="")
        self._batch_frames: list[BaseToolFrame] = []
        self._batch_index = 0
        self._current_account_path = ""
        self._temp_dir: str | None = None
        self._original_showinfo: Callable[..., object] | None = None
        self._build_form()

    def save_state(self) -> None:
        return

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, command: Callable[[], None]) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=8)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=10, pady=8)
        if label == "字典 Excel：":
            ttk.Label(parent, textvariable=self.dictionary_loading_var, foreground=COLOR_PRIMARY).grid(row=row, column=2, sticky="e", padx=(10, 4), pady=8)
            make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=3, padx=(4, 10), pady=8)
        else:
            make_rounded_button(parent, "浏览", command, width=54).grid(row=row, column=3, padx=10, pady=8)

    def _tool_needs_dictionary_sheet(self, key: str) -> bool:
        return key in {
            "posting_period_type",
            "source_media_camp_ratio",
            "post_theme_ratio",
            "sentiment_expression",
            "stance_tendency",
            "sensitive_topic_participation_rate",
        }

    def _tool_needs_common_dictionary(self, key: str) -> bool:
        return self._tool_needs_dictionary_sheet(key)

    def _tool_has_extra_group_settings(self, key: str) -> bool:
        return key == "custom_keyword_frequency"

    def _add_custom_keyword_setting_row(self, parent: tk.Widget, key: str) -> None:
        state = self.app.config.get_tool_state(key)
        keywords = [str(item).strip() for item in state.get("keywords", []) if str(item).strip()]
        text_var = tk.StringVar(value="、".join(keywords))
        input_var = tk.StringVar(value="")
        ignore_var = tk.BooleanVar(value=bool(state.get("ignore_chinese_script", False)))
        self.custom_keyword_text_vars[key] = text_var
        self.custom_keyword_input_vars[key] = input_var
        self.custom_ignore_script_vars[key] = ignore_var

        box = ttk.Frame(parent, style="Card.TFrame")
        box.pack(fill="x", padx=(28, 0), pady=(2, 8))
        ttk.Label(box, text="自定义关键词：", background=COLOR_SURFACE, foreground=COLOR_MUTED).grid(row=0, column=0, sticky="nw", padx=(0, 6), pady=4)
        entry = ttk.Entry(box, textvariable=input_var)
        entry.grid(row=0, column=1, sticky="ew", pady=4)
        make_rounded_button(box, "新增", lambda k=key: self._add_group_keyword(k), width=54).grid(row=0, column=2, padx=6, pady=4)
        make_rounded_button(box, "转繁体", lambda k=key: self._convert_group_keywords(k, True), width=66).grid(row=0, column=3, padx=6, pady=4)
        make_rounded_button(box, "转简体", lambda k=key: self._convert_group_keywords(k, False), width=66).grid(row=0, column=4, padx=6, pady=4)
        ttk.Checkbutton(box, text="不区分简繁体", variable=ignore_var).grid(row=0, column=5, padx=6, pady=4, sticky="w")
        text = tk.Text(box, height=4, wrap="word", bg="white", relief="solid", borderwidth=1)
        text._tool_key = key  # type: ignore[attr-defined]
        text.insert("1.0", text_var.get())
        text.grid(row=1, column=1, columnspan=5, sticky="ew", pady=(0, 4))
        text.bind("<KeyRelease>", lambda _e, v=text_var, w=text: v.set(w.get("1.0", "end").strip()))
        entry.bind("<Return>", lambda _e, k=key: (self._add_group_keyword(k), "break")[-1])
        box.keyword_text_widget = text  # type: ignore[attr-defined]
        box.columnconfigure(1, weight=1)

    def _group_keywords(self, key: str) -> list[str]:
        raw = self.custom_keyword_text_vars[key].get().replace("，", "、").replace(",", "、").replace("\n", "、")
        return list(dict.fromkeys(item.strip() for item in raw.split("、") if item.strip()))

    def _set_group_keywords(self, key: str, keywords: list[str]) -> None:
        value = "、".join(keywords)
        self.custom_keyword_text_vars[key].set(value)
        for widget in self._keyword_text_widgets():
            if getattr(widget, "_tool_key", None) == key:
                widget.delete("1.0", "end")
                widget.insert("1.0", value)

    def _keyword_text_widgets(self) -> list[tk.Text]:
        widgets: list[tk.Text] = []
        def walk(w: tk.Widget) -> None:
            if isinstance(w, tk.Text):
                widgets.append(w)
            for child in w.winfo_children():
                walk(child)
        walk(self)
        return widgets

    def _add_group_keyword(self, key: str) -> None:
        keyword = self.custom_keyword_input_vars[key].get().strip()
        if not keyword:
            return
        keywords = self._group_keywords(key)
        if keyword not in keywords:
            keywords.append(keyword)
        self.custom_keyword_input_vars[key].set("")
        self._set_group_keywords(key, keywords)

    def _convert_group_keywords(self, key: str, to_traditional: bool) -> None:
        from jy_toolbox.core.chinese import convert_chinese_text
        self._set_group_keywords(key, list(dict.fromkeys(convert_chinese_text(item, to_traditional=to_traditional) for item in self._group_keywords(key))))

    def _build_form(self) -> None:
        form = ttk.LabelFrame(self, text=f"{self.category} - 一键执行", style="Card.TLabelframe", padding=(12, 9))
        form.pack(fill="x", padx=22, pady=12)
        self._path_row(form, 0, "账号 Excel：", self.account_input_var, self.choose_account_input)
        self._path_row(form, 1, "帖 Excel：", self.post_input_var, self.choose_post_input)
        if any(self._tool_needs_common_dictionary(key) for key in self.tool_keys):
            self._path_row(form, 2, "字典 Excel：", self.dictionary_input_var, self.choose_dictionary_input)
        self._path_row(form, 3, "最终输出 Excel：", self.output_var, self.choose_output)
        self.account_sheet_combo = self.add_sheet_selector(form, 4, "账号表工作表：", self.account_sheet_var)
        self.post_sheet_combo = self.add_sheet_selector(form, 5, "帖表工作表：", self.post_sheet_var)
        form.columnconfigure(1, weight=1)

        list_card = ttk.LabelFrame(self, text="分组工具", style="Card.TLabelframe", padding=(12, 9))
        list_card.pack(fill="both", expand=True, padx=22, pady=8)
        canvas = tk.Canvas(list_card, highlightthickness=0, bg=COLOR_SURFACE)
        scrollbar = ttk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
        tools_body = ttk.Frame(canvas, style="Card.TFrame")
        window_id = canvas.create_window((0, 0), window=tools_body, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        tools_body.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))
        def on_tools_mousewheel(event: tk.Event) -> str:
            pointer_x = self.winfo_pointerx()
            pointer_y = self.winfo_pointery()
            x = canvas.winfo_rootx()
            y = canvas.winfo_rooty()
            if not (x <= pointer_x < x + canvas.winfo_width() and y <= pointer_y < y + canvas.winfo_height()):
                return ""
            delta = 0
            if getattr(event, "num", None) == 4:
                delta = -1
            elif getattr(event, "num", None) == 5:
                delta = 1
            elif event.delta:
                delta = -int(event.delta / 120)
            if delta:
                canvas.yview_scroll(delta, "units")
            return "break"
        canvas.bind_all("<MouseWheel>", on_tools_mousewheel)
        canvas.bind_all("<Button-4>", on_tools_mousewheel)
        canvas.bind_all("<Button-5>", on_tools_mousewheel)
        for index, key in enumerate(self.tool_keys, start=1):
            tool = self.app.tools[key]
            row = ttk.Frame(tools_body, style="Card.TFrame")
            row.pack(fill="x", pady=4)
            row.columnconfigure(1, weight=1)
            ttk.Label(row, text=f"{index}. {tool.name}", background=COLOR_SURFACE).grid(row=0, column=0, sticky="w", padx=(0, 8))
            if self._tool_needs_dictionary_sheet(key):
                ttk.Label(row, text="- - - - - -", background=COLOR_SURFACE, foreground=COLOR_MUTED).grid(row=0, column=1, sticky="ew", padx=(0, 8))
                ttk.Label(row, text="字典 sheet：", background=COLOR_SURFACE, foreground=COLOR_MUTED).grid(row=0, column=2, sticky="e", padx=(0, 4))
                var = tk.StringVar(value=self.app.config.get_tool_state(key).get("dictionary_sheet_name", ""))
                self.dictionary_sheet_vars[key] = var
                combo = ttk.Combobox(row, textvariable=var, state="readonly", values=(), width=22)
                self.dictionary_sheet_combos[key] = combo
                combo.grid(row=0, column=3, sticky="e")
            if self._tool_has_extra_group_settings(key):
                self._add_custom_keyword_setting_row(tools_body, key)

        actions = ttk.Frame(self, style="Surface.TFrame")
        actions.pack(fill="x", padx=22, pady=12)
        make_rounded_button(actions, "一键执行", self.run_batch, role="primary", width=82).pack(side="left")
        ttk.Progressbar(actions, variable=self.progress_var, maximum=max(1, len(self.tool_keys)), length=260).pack(side="left", padx=12)
        ttk.Label(actions, textvariable=self.progress_text_var, style="Muted.TLabel").pack(side="left")
        status_card = ttk.Frame(self, style="Info.TFrame", padding=(12, 9))
        status_card.pack(fill="x", padx=22, pady=8)
        ttk.Label(status_card, textvariable=self.status_var, wraplength=820, style="Info.TLabel").pack(fill="x")

    def choose_account_input(self) -> None:
        path = filedialog.askopenfilename(title="选择账号 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.account_input_var.set(path)
            self.use_first_sheet_by_default(self.account_sheet_combo, self.account_sheet_var)

    def choose_post_input(self) -> None:
        path = filedialog.askopenfilename(title="选择帖 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.post_input_var.set(path)
            self.use_first_sheet_by_default(self.post_sheet_combo, self.post_sheet_var)

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(title="保存最终账号表处理结果", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_var.set(path)

    def choose_dictionary_input(self) -> None:
        path = filedialog.askopenfilename(title="选择字典 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")])
        if path:
            self.dictionary_input_var.set(path)
            self.dictionary_loading_var.set("识别sheet名中...")
            for widget in self._dictionary_combos():
                self.populate_sheets_async(path, widget[1], widget[0], show_errors=True)
            self._wait_dictionary_sheet_loading()

    def _wait_dictionary_sheet_loading(self) -> None:
        pending_combo_keys = {id(combo) for _, combo in self._dictionary_combos()}
        if pending_combo_keys & set(self._sheet_loading_pending):
            self.after(120, self._wait_dictionary_sheet_loading)
            return
        self.dictionary_loading_var.set("")

    def _dictionary_combos(self) -> list[tuple[tk.StringVar, ttk.Combobox]]:
        return [(self.dictionary_sheet_vars[key], combo) for key, combo in self.dictionary_sheet_combos.items()]

    def run_batch(self) -> None:
        if not self.tool_keys:
            messagebox.showinfo("提示", "当前分组没有工具。", parent=self); return
        if not self.account_input_var.get().strip():
            messagebox.showwarning("提示", "请选择账号 Excel。", parent=self); return
        if any(self._tool_needs_common_dictionary(key) for key in self.tool_keys) and not self.dictionary_input_var.get().strip():
            messagebox.showwarning("提示", "当前分组包含需要字典 Excel 的工具，请选择字典 Excel。", parent=self); return
        if not self.output_var.get().strip():
            messagebox.showwarning("提示", "请选择最终输出 Excel。", parent=self); return
        self._batch_frames = []
        self._batch_index = 0
        self._current_account_path = self.account_input_var.get().strip()
        if self._temp_dir:
            shutil.rmtree(self._temp_dir, ignore_errors=True)
        self._temp_dir = tempfile.mkdtemp(prefix="jy_group_run_")
        if self._original_showinfo is None:
            self._original_showinfo = messagebox.showinfo
            messagebox.showinfo = lambda *args, **kwargs: None
        self.progress_var.set(0)
        self.progress_text_var.set(f"0/{len(self.tool_keys)}")
        self.status_var.set("正在准备一键执行……")
        self._run_next_tool()

    def _run_next_tool(self) -> None:
        if self._batch_index >= len(self.tool_keys):
            self.status_var.set("一键执行完成。")
            self._restore_showinfo()
            self._cleanup_temp_dir()
            BatchCompleteDialog(self, Path(self.output_var.get().strip()))
            return
        key = self.tool_keys[self._batch_index]
        tool = self.app.tools[key]
        frame = tool.factory(self, self.app, self.app.config.get_tool_state(key))
        frame.pack_forget()
        self._patch_frame_error_handler(frame)
        self._batch_frames.append(frame)
        step_output_path = self._step_output_path(key)
        values = (
            ("account_input_var", self._current_account_path),
            ("post_input_var", self.post_input_var.get().strip()),
            ("dictionary_input_var", self.dictionary_input_var.get().strip()),
            ("output_var", step_output_path),
            ("account_sheet_var", self.account_sheet_var.get().strip() if self._batch_index == 0 else ""),
            ("post_sheet_var", self.post_sheet_var.get().strip()),
            ("dictionary_sheet_var", self.dictionary_sheet_vars[key].get().strip() if key in self.dictionary_sheet_vars else ""),
        )
        for attr, value in values:
            if hasattr(frame, attr):
                getattr(frame, attr).set(value)
        if key == "custom_keyword_frequency":
            frame.keywords = self._group_keywords(key)
            if hasattr(frame, "ignore_chinese_script_var"):
                frame.ignore_chinese_script_var.set(self.custom_ignore_script_vars[key].get())
            if hasattr(frame, "refresh_keyword_listbox"):
                frame.refresh_keyword_listbox()
        self.status_var.set(f"正在执行：{tool.name}")
        frame.run()
        if not getattr(frame, "_background_running", False):
            self._restore_showinfo()
            self._cleanup_temp_dir()
            self.status_var.set(f"一键执行已中断：{tool.name} 未启动。")
            return
        self.after(300, lambda f=frame, k=key: self._wait_tool_done(f, k))

    def _patch_frame_error_handler(self, frame: BaseToolFrame) -> None:
        original_error = frame._finish_background_error
        def finish_error(exc: Exception, error_message: str) -> None:
            self._restore_showinfo()
            self._cleanup_temp_dir()
            self.status_var.set("一键执行已中断，请处理失败工具后重试。")
            original_error(exc, error_message)
        frame._finish_background_error = finish_error  # type: ignore[method-assign]

    def _restore_showinfo(self) -> None:
        if self._original_showinfo is not None:
            messagebox.showinfo = self._original_showinfo
            self._original_showinfo = None

    def _wait_tool_done(self, frame: BaseToolFrame, key: str) -> None:
        if getattr(frame, "_background_running", False):
            self.after(300, lambda: self._wait_tool_done(frame, key))
            return
        if hasattr(frame, "account_input_var") and hasattr(frame, "output_var"):
            output_path = getattr(frame, "output_var").get().strip()
            if output_path:
                self._current_account_path = output_path
        self._batch_index += 1
        self.progress_var.set(self._batch_index)
        self.progress_text_var.set(f"{self._batch_index}/{len(self.tool_keys)}")
        self.after(100, self._run_next_tool)

    def _step_output_path(self, key: str) -> str:
        if self._batch_index == len(self.tool_keys) - 1:
            return self.output_var.get().strip()
        temp_dir = Path(self._temp_dir or tempfile.mkdtemp(prefix="jy_group_run_"))
        self._temp_dir = str(temp_dir)
        safe_key = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in key)
        return str(temp_dir / f"{self._batch_index + 1:02d}_{safe_key}.xlsx")

    def _cleanup_temp_dir(self) -> None:
        if self._temp_dir:
            shutil.rmtree(self._temp_dir, ignore_errors=True)
            self._temp_dir = None

class BatchCompleteDialog(tk.Toplevel):
    """一键执行完成提示，支持打开最终输出目录。"""

    def __init__(self, parent: tk.Widget, output_path: Path) -> None:
        super().__init__(parent)
        self.output_path = output_path
        self.title("完成")
        self.configure(bg=COLOR_BG)
        self.transient(parent.winfo_toplevel())
        self.grab_set()

        shell = ttk.Frame(self, style="Surface.TFrame", padding=(18, 16))
        shell.pack(fill="both", expand=True)
        ttk.Label(shell, text=f"分组工具已全部执行完成。\n最终输出：{output_path}", wraplength=520).pack(anchor="w")
        buttons = ttk.Frame(shell, style="Surface.TFrame")
        buttons.pack(fill="x", pady=(16, 0))
        make_rounded_button(buttons, "关闭", self.destroy, width=62).pack(side="right")
        make_rounded_button(buttons, "打开文件所在目录", self.open_output_folder, role="primary", width=132).pack(side="right", padx=(0, 8))

    def open_output_folder(self) -> None:
        folder = self.output_path.parent
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder)])
        else:
            subprocess.Popen(["xdg-open", str(folder)])
        self.destroy()
