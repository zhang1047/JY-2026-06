#!/usr/bin/env python3
"""应用启动入口。"""
from __future__ import annotations

from jy_toolbox.ui.app import ToolboxApp


def main() -> None:
    app = ToolboxApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()
