from __future__ import annotations

import os
import tempfile
from pathlib import Path
from typing import Any

def list_excel_sheet_names_with_passwords(path: Path, passwords: list[str]) -> list[str]:
    """列出普通或加密 Excel 的工作表名称。"""
    import pandas as pd

    try:
        return list(pd.ExcelFile(path).sheet_names)
    except Exception as first_error:  # noqa: BLE001 - 需要判断是否可用密码继续尝试
        if not passwords:
            raise RuntimeError(f"读取工作表失败；如果文件有打开密码，请先在密码本新增密码。原始错误：{first_error}") from first_error
        try:
            import msoffcrypto
        except ImportError as import_error:
            raise RuntimeError(
                "文件可能已加密，但缺少 msoffcrypto-tool 依赖；请执行：pip install msoffcrypto-tool"
            ) from import_error

        last_error: Exception | None = None
        for password in passwords:
            tmp_name = ""
            try:
                with path.open("rb") as source:
                    office_file = msoffcrypto.OfficeFile(source)
                    office_file.load_key(password=str(password))
                    with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as decrypted:
                        tmp_name = decrypted.name
                        office_file.decrypt(decrypted)
                return list(pd.ExcelFile(tmp_name).sheet_names)
            except Exception as exc:  # noqa: BLE001 - 尝试下一个密码
                last_error = exc
            finally:
                if tmp_name:
                    try:
                        os.unlink(tmp_name)
                    except OSError:
                        pass
        raise RuntimeError("读取工作表失败：已按密码本逐个尝试，但没有密码可以打开该 Excel。") from last_error


def read_excel_with_passwords(path: Path, passwords: list[str], sheet_name: str | int) -> Any:
    """读取普通或加密 Excel。加密文件会按密码本顺序尝试。"""
    import pandas as pd

    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception as first_error:  # noqa: BLE001 - 需要判断是否可用密码继续尝试
        if not passwords:
            raise RuntimeError(f"读取失败；如果文件有打开密码，请先在密码本新增密码。原始错误：{first_error}") from first_error
        try:
            import msoffcrypto
        except ImportError as import_error:
            raise RuntimeError(
                "文件可能已加密，但缺少 msoffcrypto-tool 依赖；请执行：pip install msoffcrypto-tool"
            ) from import_error

        last_error: Exception | None = None
        for password in passwords:
            tmp_name = ""
            try:
                with path.open("rb") as source:
                    office_file = msoffcrypto.OfficeFile(source)
                    office_file.load_key(password=str(password))
                    with tempfile.NamedTemporaryFile(delete=False, suffix=path.suffix) as decrypted:
                        tmp_name = decrypted.name
                        office_file.decrypt(decrypted)
                return pd.read_excel(tmp_name, sheet_name=sheet_name)
            except Exception as exc:  # noqa: BLE001 - 尝试下一个密码
                last_error = exc
            finally:
                if tmp_name:
                    try:
                        os.unlink(tmp_name)
                    except OSError:
                        pass
        raise RuntimeError("读取失败：已按密码本逐个尝试，但没有密码可以打开该 Excel。") from last_error
