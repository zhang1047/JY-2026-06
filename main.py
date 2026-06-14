#!/usr/bin/env python3
"""应用启动入口。"""
from __future__ import annotations

import os
import sys
import threading

from jy_toolbox.ui.app import ToolboxApp


_LIBPNG_SRGB_WARNING = b"libpng warning: iCCP: known incorrect sRGB profile"


def _forward_filtered_stderr(read_fd: int, original_stderr_fd: int) -> None:
    """Forward stderr from a pipe while dropping known noisy warnings."""
    with os.fdopen(read_fd, "rb", closefd=True) as reader:
        with os.fdopen(original_stderr_fd, "wb", closefd=True) as writer:
            for line in reader:
                if _LIBPNG_SRGB_WARNING in line:
                    continue
                writer.write(line)
                writer.flush()


def _install_known_stderr_filter() -> None:
    """Suppress noisy platform Tk/libpng profile warnings while preserving other stderr output."""
    if os.environ.get("JY_TOOLBOX_NO_STDERR_FILTER"):
        return
    try:
        read_fd, write_fd = os.pipe()
        original_stderr_fd = os.dup(2)
        os.dup2(write_fd, 2)
        os.close(write_fd)
    except OSError:
        return

    threading.Thread(
        target=_forward_filtered_stderr,
        args=(read_fd, original_stderr_fd),
        name="stderr-filter",
        daemon=True,
    ).start()
    sys.stderr = os.fdopen(os.dup(2), "w", encoding=sys.stderr.encoding or "utf-8", errors="replace", buffering=1)


def main() -> None:
    _install_known_stderr_filter()
    app = ToolboxApp()
    app.root.mainloop()


if __name__ == "__main__":
    main()
