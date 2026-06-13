from __future__ import annotations

import ctypes
import os
import tkinter as tk

def enable_light_title_bar(root: tk.Tk | tk.Toplevel) -> None:
    """在 Windows 上使用浅色标题栏，配合浅灰色整体主题。"""
    if os.name != "nt":
        return
    root.update_idletasks()
    hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
    value = ctypes.c_int(0)
    for attribute in (20, 19):
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))
