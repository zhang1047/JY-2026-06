from __future__ import annotations

import os
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any


def _password_candidates(passwords: list[str]) -> list[str]:
    """按密码本顺序生成尝试密码；兼容误输入前后空格的常见情况。"""
    candidates: list[str] = []
    seen: set[str] = set()
    for password in passwords:
        text = str(password)
        for candidate in (text, text.strip()):
            if candidate and candidate not in seen:
                candidates.append(candidate)
                seen.add(candidate)
    return candidates


def _read_decrypted_excel(path: Path, passwords: list[str], reader: Callable[[str], Any], failure_prefix: str) -> Any:
    candidates = _password_candidates(passwords)
    if not candidates:
        raise RuntimeError(f"{failure_prefix}；如果文件有打开密码，请先在密码本新增密码。")
    try:
        import msoffcrypto
    except ImportError as import_error:
        raise RuntimeError(
            "文件可能已加密，但缺少 msoffcrypto-tool 依赖；请执行：pip install msoffcrypto-tool"
        ) from import_error

    last_error: Exception | None = None
    for password in candidates:
        tmp_name = ""
        try:
            with path.open("rb") as source:
                office_file = msoffcrypto.OfficeFile(source)
                office_file.load_key(password=password, verify_password=True)
                with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as decrypted:
                    tmp_name = decrypted.name
                    office_file.decrypt(decrypted)
            return reader(tmp_name)
        except Exception as exc:  # noqa: BLE001 - 尝试下一个密码
            last_error = exc
        finally:
            if tmp_name:
                try:
                    os.unlink(tmp_name)
                except OSError:
                    pass
    raise RuntimeError(f"{failure_prefix}：已按密码本逐个尝试，但没有密码可以打开该 Excel。") from last_error


def list_excel_sheet_names_with_passwords(path: Path, passwords: list[str]) -> list[str]:
    """列出普通或加密 Excel 的工作表名称。"""
    import pandas as pd

    try:
        return list(pd.ExcelFile(path).sheet_names)
    except Exception as first_error:  # noqa: BLE001 - 需要判断是否可用密码继续尝试
        try:
            return _read_decrypted_excel(
                path,
                passwords,
                lambda tmp_name: list(pd.ExcelFile(tmp_name).sheet_names),
                "读取工作表失败",
            )
        except RuntimeError as encrypted_error:
            raise RuntimeError(f"{encrypted_error} 原始错误：{first_error}") from encrypted_error


def read_excel_with_passwords(path: Path, passwords: list[str], sheet_name: str | int) -> Any:
    """读取普通或加密 Excel。加密文件会按密码本顺序尝试。"""
    import pandas as pd

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as first_error:  # noqa: BLE001 - 需要判断是否可用密码继续尝试
        try:
            return _read_decrypted_excel(
                path,
                passwords,
                lambda tmp_name: pd.read_excel(tmp_name, sheet_name=sheet_name),
                "读取失败",
            )
        except RuntimeError as encrypted_error:
            raise RuntimeError(f"{encrypted_error} 原始错误：{first_error}") from encrypted_error
