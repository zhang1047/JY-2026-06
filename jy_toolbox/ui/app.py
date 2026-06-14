from __future__ import annotations

from typing import Any
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from jy_toolbox.core.config import ConfigStore
from jy_toolbox.core.constants import *
from jy_toolbox.core.models import ToolDefinition
from jy_toolbox.core.platform import enable_light_title_bar
from jy_toolbox.ui.dialogs import PasswordBookDialog, ToolListDialog
from jy_toolbox.ui.tools import (
    ActiveDayRatioTool,
    AddedOpinionShareRateTool,
    AverageOriginalPostInteractionsTool,
    AveragePostLengthTool,
    DailyActiveSpanTool,
    PostDedupTool,
    PostTypeRatioTool,
    PostThemeRatioTool,
    SentimentExpressionTool,
    SourceMediaCampRatioTool,
    StanceTendencyTool,
)
from jy_toolbox.ui.base import BaseToolFrame
from jy_toolbox.ui.widgets import make_rounded_button, rounded_rect_points

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
                    "按规则统计每个账号文字、图片、视频贴文占比，并在账号表最后新增三列占比。"
                ),
                factory=lambda parent, app, state: PostTypeRatioTool(
                    parent, app, state, app.get_tool_description("post_type_ratio")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="post_theme_ratio",
                name="帖子主题占比（%）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、贴文 Excel 和内容偏好字典 Excel，通过账号表“FB主页”与贴文表“主页url”关联，"
                    "再用贴文表“帖子正文”匹配字典 sheet 第一列“帖子正文”的“内容偏好”分类，"
                    "按账号计算各内容偏好分类占比，并在账号表最后新增“主题占比-分类名”列。"
                ),
                factory=lambda parent, app, state: PostThemeRatioTool(
                    parent, app, state, app.get_tool_description("post_theme_ratio")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="sentiment_expression",
                name="情感表达分数&数量占比",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel、贴文 Excel 和情感表达字典 Excel，通过账号表“FB主页”与贴文表“主页url”关联，"
                    "再用贴文表“帖子正文”匹配字典 sheet 第一列“帖子正文”的“情感表达倾向”枚举，"
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
                    "说明：选择账号 Excel、贴文 Excel 和立场倾向字典 Excel，通过账号表“FB主页”与贴文表“主页url”关联，"
                    "再用贴文表“帖子正文”匹配字典 sheet 第一列“帖子正文”的“两岸议题立场倾向”枚举，"
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
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联；"
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
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联；"
                    "按每个账号最早到最晚的贴文发布时间计算账号发帖时间范围天数，"
                    "再用该账号实际发帖日期数除以时间范围天数，并在账号表最后依次新增“帖子数量”、"
                    "“帖子时间跨度天数”、“活跃天数”和“活跃天数占比”列。"
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
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联；"
                    "只统计贴文表“帖子正文”列的文字长度，按账号计算平均每个帖子的正文长度，"
                    "并在账号表最后新增“平均发帖长度”列。"
                ),
                factory=lambda parent, app, state: AveragePostLengthTool(
                    parent, app, state, app.get_tool_description("average_post_length")
                ),
            )
        )
        self.add_tool(
            ToolDefinition(
                key="average_original_post_interactions",
                name="平均原创单帖互动数（条）",
                default_category="Excel 工具",
                description=(
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联；"
                    "仅统计贴文表“创作类型”为 common 的原创帖，将每帖“点赞数”“评论数”“分享数”相加得到互动数，"
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
                    "说明：选择账号 Excel 和贴文 Excel，通过账号表“FB主页”与贴文表“主页url”关联；"
                    "仅统计“创作类型”为 share 的转发贴，其中“标题”或“帖子正文”任一不为空即视为附加观点，"
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
                    "说明：选择账号 Excel、贴文 Excel 和账号名字典 Excel；仅统计贴文表中“创作类型”为 share 的转发贴，"
                    "用“分享贴账号名”匹配字典中的“账号立场归属”和“账号类型归属”，"
                    "按账号汇总各分类占比，并在账号表最后新增两列占比。"
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
            make_rounded_button(header, "改名", lambda c=category: self.rename_category(c), width=48, height=24).pack(side="right")
            make_rounded_button(header, "删除", lambda c=category: self.delete_category(c), role="danger", width=48, height=24).pack(side="right", padx=6)

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
            messagebox.showwarning("提示", "至少保留一个分类。", parent=self.root)
            return
        target = next((c for c in categories if c != category), None)
        if not messagebox.askyesno("删除分类", f"确认删除分类“{category}”？其中工具将移动到“{target}”。", parent=self.root):
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
