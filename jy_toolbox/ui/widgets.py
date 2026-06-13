from __future__ import annotations

from typing import Any, Callable
import tkinter as tk
from tkinter import ttk

from jy_toolbox.core.constants import *

BUTTON_PALETTES = {
    "normal": {"bg": COLOR_BUTTON, "hover": COLOR_BUTTON_HOVER, "pressed": COLOR_BUTTON_PRESSED, "fg": COLOR_TEXT, "border": COLOR_BORDER},
    "primary": {"bg": COLOR_PRIMARY, "hover": COLOR_PRIMARY_HOVER, "pressed": COLOR_PRIMARY_PRESSED, "fg": "#ffffff", "border": COLOR_PRIMARY_DARK},
    "danger": {"bg": COLOR_DANGER, "hover": COLOR_DANGER_HOVER, "pressed": COLOR_DANGER_PRESSED, "fg": "#ffffff", "border": COLOR_DANGER_DARK},
    "warning": {"bg": COLOR_WARNING, "hover": COLOR_WARNING_DARK, "pressed": COLOR_WARNING_PRESSED, "fg": "#ffffff", "border": COLOR_WARNING_DARK},
    "selected": {"bg": COLOR_PRIMARY, "hover": COLOR_PRIMARY_HOVER, "pressed": COLOR_PRIMARY_PRESSED, "fg": "#ffffff", "border": COLOR_PRIMARY_DARK},
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
