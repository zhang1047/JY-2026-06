def mask_password(password: str) -> str:
    """隐藏密码中间字符，例如 qwe123 -> q****3。"""
    if not password:
        return ""
    if len(password) == 1:
        return "*"
    if len(password) == 2:
        return password[0] + "*"
    return password[0] + ("*" * max(1, len(password) - 2)) + password[-1]
