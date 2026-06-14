from __future__ import annotations

import random
from pathlib import Path
from typing import Any, Callable

from jy_toolbox.services.excel_io import read_excel_with_passwords

SENTIMENT_VALUES = ("正面", "负面", "中性")
SENTIMENT_SCORE_MAP = {"正面": 1, "负面": -1, "中性": 0}
SENTIMENT_OUTPUT_COLUMNS = (
    "情感表达-分数",
    "情感表达-正面（数量）",
    "情感表达-正面（占比）",
    "情感表达-负面（数量）",
    "情感表达-负面（占比）",
    "情感表达-中性（数量）",
    "情感表达-中性（占比）",
)

STANCE_VALUES = ("偏蓝", "中立", "偏绿")
STANCE_SCORE_MAP = {"偏蓝": 1, "中立": 0, "偏绿": -1}
STANCE_OUTPUT_COLUMNS = (
    "立场倾向-分数",
    "立场倾向-偏蓝（数量）",
    "立场倾向-偏蓝（占比）",
    "立场倾向-偏绿（数量）",
    "立场倾向-偏绿（占比）",
    "立场倾向-中立（数量）",
    "立场倾向-中立（占比）",
)


def deduplicate_posts_excel(
    input_path: Path,
    output_path: Path,
    passwords: list[str],
    sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    import pandas as pd

    if not input_path.exists():
        raise FileNotFoundError(f"输入文件不存在：{input_path}")
    df = read_excel_with_passwords(input_path, passwords, sheet_name)
    if progress is not None:
        progress(35, "已读取 Excel，正在检查必要列……")
    missing = [col for col in ["贴文url", "点赞数", "分享数", "评论数"] if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要列：{', '.join(missing)}")

    work = df.copy()
    if progress is not None:
        progress(55, "正在计算重复贴文保留规则……")
    score_cols = ["点赞数", "分享数", "评论数"]
    numeric_scores = work[score_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    work["__dedup_score__"] = numeric_scores.sum(axis=1)
    work["__dedup_random__"] = [random.random() for _ in range(len(work))]
    work["__dedup_order__"] = range(len(work))
    kept = (
        work.sort_values(["贴文url", "__dedup_score__", "__dedup_random__"], ascending=[True, False, False])
        .drop_duplicates(subset=["贴文url"], keep="first")
        .sort_values("__dedup_order__")
        .drop(columns=["__dedup_score__", "__dedup_random__", "__dedup_order__"])
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(85, "正在写出处理结果……")
    kept.to_excel(output_path, index=False)
    return {"original": len(df), "removed": len(df) - len(kept), "kept": len(kept)}

def _is_non_empty_cell(value: Any) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return bool(text) and text.lower() != "nan"

def _normalized_key(value: Any) -> str:
    return str(value).strip() if _is_non_empty_cell(value) else ""

def classify_post_type(row: Any) -> str:
    homepage_url = "" if not _is_non_empty_cell(row.get("主页url")) else str(row.get("主页url"))
    post_url = "" if not _is_non_empty_cell(row.get("贴文url")) else str(row.get("贴文url"))
    if "/videos/" in homepage_url.lower() or "/videos/" in post_url.lower():
        return "视频"
    attachments = "" if not _is_non_empty_cell(row.get("图片附件")) else str(row.get("图片附件"))
    if attachments.count("origin_url_md5") >= 2:
        return "图片"
    if str(row.get("创作类型", "")).strip().lower() == "common":
        if _is_non_empty_cell(row.get("标题")) or _is_non_empty_cell(row.get("帖子正文")):
            return "文字"
    return ""

def _format_post_type_ratios(counts: dict[str, int]) -> dict[str, str]:
    total = sum(counts.values())
    if total <= 0:
        return {column: "" for column in ("文字帖占比", "图片帖占比", "视频帖占比")}
    return {
        "文字帖占比": f"{counts.get('文字', 0) / total * 100:.2f}%",
        "图片帖占比": f"{counts.get('图片', 0) / total * 100:.2f}%",
        "视频帖占比": f"{counts.get('视频', 0) / total * 100:.2f}%",
    }

def calculate_post_type_ratios_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "图片附件", "创作类型", "标题", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(50, "正在识别每条贴文的文字、图片、视频类型……")
    work = post_df.copy()
    work["__post_type__"] = work.apply(classify_post_type, axis=1)
    typed_work = work[work["__post_type__"].isin(["文字", "图片", "视频"])].copy()

    empty_ratios = {column: "" for column in ("文字帖占比", "图片帖占比", "视频帖占比")}
    ratios_by_homepage: dict[str, dict[str, str]] = {}
    if not typed_work.empty:
        grouped = typed_work.groupby("主页url")["__post_type__"].value_counts()
        for homepage, counts_series in grouped.groupby(level=0):
            counts = {str(type_name): int(count) for (_, type_name), count in counts_series.items()}
            ratios_by_homepage[_normalized_key(homepage)] = _format_post_type_ratios(counts)

    if progress is not None:
        progress(75, "正在写回账号表最后三列类型占比……")
    output_df = account_df.copy()
    drop_columns = [
        column
        for column in (PostTypeRatioTool.LEGACY_OUTPUT_COLUMN, *("文字帖占比", "图片帖占比", "视频帖占比"))
        if column in output_df.columns
    ]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in ("文字帖占比", "图片帖占比", "视频帖占比"):
        output_df[column] = account_keys.map(lambda key, column=column: ratios_by_homepage.get(key, empty_ratios).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(90, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(_normalized_key(v) for v in post_df["主页url"])).sum())
    typed_accounts = int(output_df[("文字帖占比", "图片帖占比", "视频帖占比")[0]].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "matched_accounts": matched_accounts,
        "typed_accounts": typed_accounts,
    }

def _post_body_text_length(value: Any) -> int:
    if not _is_non_empty_cell(value):
        return 0
    return len(str(value).strip())

def calculate_average_post_length_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在计算每条贴文正文长度……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work = work[work["__homepage_key__"] != ""].copy()
    work["__body_length__"] = work["帖子正文"].map(_post_body_text_length)

    if progress is not None:
        progress(72, "正在按账号汇总平均发帖长度……")
    average_by_homepage: dict[str, float] = {}
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        average_by_homepage[str(homepage)] = round(float(homepage_rows["__body_length__"].mean()), 2)

    if progress is not None:
        progress(84, "正在写回账号表平均发帖长度列……")
    output_df = account_df.copy()
    if "平均发帖长度" in output_df.columns:
        output_df = output_df.drop(columns=["平均发帖长度"])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df["平均发帖长度"] = account_keys.map(lambda key: average_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(work["__homepage_key__"])).sum())
    averaged_accounts = int(output_df["平均发帖长度"].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "matched_accounts": matched_accounts,
        "averaged_accounts": averaged_accounts,
    }



def calculate_daily_active_span_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "贴文发布时间"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在解析贴文发布时间并筛选单日 2 条及以上的发帖日期……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__post_datetime__"] = pd.to_datetime(work["贴文发布时间"], errors="coerce")
    work = work[(work["__homepage_key__"] != "") & work["__post_datetime__"].notna()].copy()
    work["__post_date__"] = work["__post_datetime__"].dt.date

    if progress is not None:
        progress(72, "正在按账号计算日均在线活跃时段跨度……")
    span_by_homepage: dict[str, float] = {}
    qualified_days_count = 0
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        daily_spans: list[float] = []
        for _, day_rows in homepage_rows.groupby("__post_date__", sort=False):
            if len(day_rows) < 2:
                continue
            span_hours = (day_rows["__post_datetime__"].max() - day_rows["__post_datetime__"].min()).total_seconds() / 3600
            daily_spans.append(float(span_hours))
        if daily_spans:
            span_by_homepage[str(homepage)] = round(sum(daily_spans) / len(daily_spans), 2)
            qualified_days_count += len(daily_spans)

    if progress is not None:
        progress(84, "正在写回账号表日均在线活跃时段跨度列……")
    output_df = account_df.copy()
    output_column = "日均在线活跃时段跨度（小时/天）"
    if output_column in output_df.columns:
        output_df = output_df.drop(columns=[output_column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[output_column] = account_keys.map(lambda key: span_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(work["__homepage_key__"])).sum())
    spanned_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "valid_time_posts": len(work),
        "matched_accounts": matched_accounts,
        "qualified_days": qualified_days_count,
        "spanned_accounts": spanned_accounts,
    }

def calculate_active_day_ratio_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "贴文发布时间"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在统计帖子数量并解析贴文发布时间……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    post_counts_by_homepage = (
        work[work["__homepage_key__"] != ""].groupby("__homepage_key__", sort=False).size().to_dict()
    )
    work["__post_datetime__"] = pd.to_datetime(work["贴文发布时间"], errors="coerce")
    work = work[(work["__homepage_key__"] != "") & work["__post_datetime__"].notna()].copy()
    work["__post_date__"] = work["__post_datetime__"].dt.date

    if progress is not None:
        progress(72, "正在按账号汇总活跃天数占比……")
    span_days_by_homepage: dict[str, int] = {}
    active_days_by_homepage: dict[str, int] = {}
    ratios_by_homepage: dict[str, str] = {}
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        first_date = homepage_rows["__post_date__"].min()
        last_date = homepage_rows["__post_date__"].max()
        span_days = (last_date - first_date).days + 1
        active_days = int(homepage_rows["__post_date__"].nunique())
        homepage_key = str(homepage)
        span_days_by_homepage[homepage_key] = span_days
        active_days_by_homepage[homepage_key] = active_days
        ratios_by_homepage[homepage_key] = _format_single_percentage(active_days, span_days)

    if progress is not None:
        progress(84, "正在写回账号表帖子数量、帖子时间跨度天数、活跃天数和活跃天数占比列……")
    output_df = account_df.copy()
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df["帖子数量"] = account_keys.map(lambda key: int(post_counts_by_homepage.get(key, 0)))
    output_df["帖子时间跨度天数"] = account_keys.map(lambda key: span_days_by_homepage.get(key, ""))
    output_df["活跃天数"] = account_keys.map(lambda key: active_days_by_homepage.get(key, ""))
    output_df["活跃天数占比"] = account_keys.map(lambda key: ratios_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(post_counts_by_homepage)).sum())
    post_count_accounts = int((output_df["帖子数量"] > 0).sum())
    active_span_accounts = int(output_df["帖子时间跨度天数"].map(_is_non_empty_cell).sum())
    active_days_accounts = int(output_df["活跃天数"].map(_is_non_empty_cell).sum())
    active_ratio_accounts = int(output_df["活跃天数占比"].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "valid_time_posts": len(work),
        "matched_accounts": matched_accounts,
        "post_count_accounts": post_count_accounts,
        "active_span_accounts": active_span_accounts,
        "active_days_accounts": active_days_accounts,
        "active_ratio_accounts": active_ratio_accounts,
    }

def _format_category_ratios(counts: dict[str, int]) -> str:
    total = sum(counts.values())
    if total <= 0:
        return ""
    return "；".join(f"{name}：{count / total * 100:.2f}%" for name, count in counts.items())

def _format_single_percentage(count: int, total: int) -> str:
    if total <= 0:
        return ""
    return f"{count / total * 100:.2f}%"

def calculate_added_opinion_share_rate_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "创作类型", "标题", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在识别转发贴是否附加标题或帖子正文……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__creation_type__"] = work["创作类型"].map(
        lambda value: str(value).strip().lower() if _is_non_empty_cell(value) else ""
    )
    share_work = work[(work["__creation_type__"] == "share") & (work["__homepage_key__"] != "")].copy()
    share_work["__has_added_opinion__"] = share_work.apply(
        lambda row: _is_non_empty_cell(row.get("标题"))
        or _is_non_empty_cell(row.get("帖子正文")),
        axis=1,
    )

    if progress is not None:
        progress(72, "正在按账号汇总附加观点转发率……")
    ratios_by_homepage: dict[str, str] = {}
    for homepage, homepage_rows in share_work.groupby("__homepage_key__", sort=False):
        added_opinion_count = int(homepage_rows["__has_added_opinion__"].sum())
        ratios_by_homepage[str(homepage)] = _format_single_percentage(added_opinion_count, len(homepage_rows))

    if progress is not None:
        progress(84, "正在写回账号表附加观点转发率列……")
    output_df = account_df.copy()
    if "附加观点转发率" in output_df.columns:
        output_df = output_df.drop(columns=["附加观点转发率"])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df["附加观点转发率"] = account_keys.map(lambda key: ratios_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(share_work["__homepage_key__"])).sum())
    rated_accounts = int(output_df["附加观点转发率"].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "share_posts": len(share_work),
        "matched_accounts": matched_accounts,
        "rated_accounts": rated_accounts,
    }

def calculate_source_media_camp_ratios_excel(
    account_input_path: Path,
    post_input_path: Path,
    dictionary_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    dictionary_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取账号名字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "创作类型", "分享贴账号名"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")
    missing_dictionary = [col for col in ["分享贴账号名", "账号立场归属", "账号类型归属"] if col not in dictionary_df.columns]
    if missing_dictionary:
        raise ValueError(f"字典 Excel 缺少必要列：{', '.join(missing_dictionary)}")

    if progress is not None:
        progress(55, "正在按字典匹配分享贴账号名……")
    dictionary_work = dictionary_df.copy()
    dictionary_work["__shared_account_key__"] = dictionary_work["分享贴账号名"].map(
        _normalized_key
    )
    dictionary_work = dictionary_work[dictionary_work["__shared_account_key__"] != ""].drop_duplicates(
        subset=["__shared_account_key__"],
        keep="first",
    )
    stance_by_shared_account = dictionary_work.set_index("__shared_account_key__")[
        SourceMediaCampRatioTool.STANCE_COLUMN
    ].to_dict()
    type_by_shared_account = dictionary_work.set_index("__shared_account_key__")[
        SourceMediaCampRatioTool.ACCOUNT_TYPE_COLUMN
    ].to_dict()

    post_work = post_df.copy()
    post_work["__homepage_key__"] = post_work["主页url"].map(_normalized_key)
    post_work["__shared_account_key__"] = post_work["分享贴账号名"].map(
        _normalized_key
    )
    post_work["__creation_type__"] = post_work["创作类型"].map(
        lambda value: str(value).strip().lower() if _is_non_empty_cell(value) else ""
    )
    share_work = post_work[
        (post_work["__creation_type__"] == "share")
        & (post_work["__homepage_key__"] != "")
        & (post_work["__shared_account_key__"] != "")
    ].copy()
    share_work["__stance__"] = share_work["__shared_account_key__"].map(stance_by_shared_account)
    share_work["__account_type__"] = share_work["__shared_account_key__"].map(type_by_shared_account)
    dictionary_keys = set(dictionary_work["__shared_account_key__"])
    matched_share_posts = int(share_work["__shared_account_key__"].isin(dictionary_keys).sum())

    if progress is not None:
        progress(70, "正在汇总每个账号的立场、类型和官方信源占比……")
    ratios_by_homepage: dict[str, dict[str, str]] = {}
    for homepage, homepage_rows in share_work.groupby("__homepage_key__", sort=False):
        stance_counts = {
            str(name): int(count)
            for name, count in homepage_rows.loc[
                homepage_rows["__stance__"].map(_is_non_empty_cell), "__stance__"
            ].value_counts(sort=False).items()
        }
        type_counts = {
            str(name): int(count)
            for name, count in homepage_rows.loc[
                homepage_rows["__account_type__"].map(_is_non_empty_cell), "__account_type__"
            ].value_counts(sort=False).items()
        }
        official_source_count = type_counts.get(SourceMediaCampRatioTool.OFFICIAL_SOURCE_TYPE, 0)
        total_typed_sources = sum(type_counts.values())
        ratios_by_homepage[str(homepage)] = {
            "账号立场归属占比": _format_category_ratios(stance_counts),
            "账号类型归属占比": _format_category_ratios(type_counts),
            "官方信源占比": _format_single_percentage(official_source_count, total_typed_sources),
        }

    if progress is not None:
        progress(82, "正在写回账号表最后三列媒体阵营分布……")
    output_df = account_df.copy()
    drop_columns = [column for column in SourceMediaCampRatioTool.OUTPUT_COLUMNS if column in output_df.columns]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in SourceMediaCampRatioTool.OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(
            lambda key, column=column: ratios_by_homepage.get(key, {}).get(column, "")
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(92, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    classified_accounts = int(
        output_df[list(SourceMediaCampRatioTool.OUTPUT_COLUMNS)].apply(
            lambda row: any(_is_non_empty_cell(value) for value in row),
            axis=1,
        ).sum()
    )
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "share_posts": len(share_work),
        "matched_share_posts": matched_share_posts,
        "classified_accounts": classified_accounts,
    }


def calculate_post_theme_ratios_excel(
    account_input_path: Path,
    post_input_path: Path,
    dictionary_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    dictionary_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取内容偏好字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError('账号 Excel 缺少必要列：FB主页')
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")
    if len(dictionary_df.columns) < 1 or str(dictionary_df.columns[0]).strip() != "帖子正文":
        raise ValueError("字典 Excel 第一列必须是：帖子正文")
    if "内容偏好" not in dictionary_df.columns:
        raise ValueError("字典 Excel 缺少必要列：内容偏好")

    if progress is not None:
        progress(55, "正在按帖子正文匹配内容偏好分类……")
    dictionary_work = dictionary_df.copy()
    dictionary_work["__body_key__"] = dictionary_work["帖子正文"].map(_normalized_key)
    dictionary_work["__theme__"] = dictionary_work["内容偏好"].map(_normalized_key)
    dictionary_work = dictionary_work[
        (dictionary_work["__body_key__"] != "") & (dictionary_work["__theme__"] != "")
    ].drop_duplicates(subset=["__body_key__"], keep="first")
    theme_by_body = dictionary_work.set_index("__body_key__")["__theme__"].to_dict()
    theme_names = list(dict.fromkeys(str(value) for value in dictionary_work["__theme__"]))
    output_columns = [f"主题占比-{theme}" for theme in theme_names]

    post_work = post_df.copy()
    post_work["__homepage_key__"] = post_work["主页url"].map(_normalized_key)
    post_work["__body_key__"] = post_work["帖子正文"].map(_normalized_key)
    post_work["__theme__"] = post_work["__body_key__"].map(theme_by_body)
    matched_work = post_work[
        (post_work["__homepage_key__"] != "") & post_work["__theme__"].map(_is_non_empty_cell)
    ].copy()

    if progress is not None:
        progress(72, "正在按账号汇总帖子主题占比……")
    ratios_by_homepage: dict[str, dict[str, str]] = {}
    for homepage, homepage_rows in matched_work.groupby("__homepage_key__", sort=False):
        counts = {str(name): int(count) for name, count in homepage_rows["__theme__"].value_counts(sort=False).items()}
        total = sum(counts.values())
        ratios_by_homepage[str(homepage)] = {
            f"主题占比-{theme}": _format_single_percentage(counts.get(theme, 0), total)
            for theme in theme_names
        }

    if progress is not None:
        progress(84, "正在写回账号表帖子主题占比列……")
    output_df = account_df.copy()
    existing_theme_columns = [column for column in output_df.columns if str(column).startswith("主题占比-")]
    if existing_theme_columns:
        output_df = output_df.drop(columns=existing_theme_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in output_columns:
        output_df[column] = account_keys.map(lambda key, column=column: ratios_by_homepage.get(key, {}).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    themed_accounts = int(
        output_df[output_columns].apply(lambda row: any(_is_non_empty_cell(value) for value in row), axis=1).sum()
    ) if output_columns else 0
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "dictionary_rows": len(dictionary_df),
        "themes": len(theme_names),
        "matched_posts": len(matched_work),
        "themed_accounts": themed_accounts,
    }

def calculate_sentiment_expression_excel(
    account_input_path: Path,
    post_input_path: Path,
    dictionary_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    dictionary_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取情感表达字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")
    if len(dictionary_df.columns) < 1 or str(dictionary_df.columns[0]).strip() != "帖子正文":
        raise ValueError("字典 Excel 第一列必须是：帖子正文")
    if "情感表达倾向" not in dictionary_df.columns:
        raise ValueError("字典 Excel 缺少必要列：情感表达倾向")

    if progress is not None:
        progress(55, "正在按帖子正文匹配情感表达倾向……")
    dictionary_work = dictionary_df.copy()
    dictionary_work["__body_key__"] = dictionary_work["帖子正文"].map(_normalized_key)
    dictionary_work["__sentiment__"] = dictionary_work["情感表达倾向"].map(_normalized_key)
    dictionary_work = dictionary_work[
        (dictionary_work["__body_key__"] != "") & dictionary_work["__sentiment__"].isin(SENTIMENT_VALUES)
    ].drop_duplicates(subset=["__body_key__"], keep="first")
    sentiment_by_body = dictionary_work.set_index("__body_key__")["__sentiment__"].to_dict()

    post_work = post_df.copy()
    post_work["__homepage_key__"] = post_work["主页url"].map(_normalized_key)
    post_work["__body_key__"] = post_work["帖子正文"].map(_normalized_key)
    post_work["__sentiment__"] = post_work["__body_key__"].map(sentiment_by_body)
    matched_work = post_work[
        (post_work["__homepage_key__"] != "") & post_work["__sentiment__"].map(_is_non_empty_cell)
    ].copy()

    if progress is not None:
        progress(72, "正在按账号汇总情感表达分数、数量和占比……")
    stats_by_homepage: dict[str, dict[str, str | int]] = {}
    for homepage, homepage_rows in matched_work.groupby("__homepage_key__", sort=False):
        counts = {name: int((homepage_rows["__sentiment__"] == name).sum()) for name in SENTIMENT_VALUES}
        total = sum(counts.values())
        stats: dict[str, str | int] = {
            "情感表达-分数": sum(counts[name] * SENTIMENT_SCORE_MAP[name] for name in SENTIMENT_VALUES),
        }
        for name in SENTIMENT_VALUES:
            stats[f"情感表达-{name}（数量）"] = counts[name]
            stats[f"情感表达-{name}（占比）"] = _format_single_percentage(counts[name], total)
        stats_by_homepage[str(homepage)] = stats

    if progress is not None:
        progress(84, "正在写回账号表情感表达统计列……")
    output_df = account_df.copy()
    drop_columns = [column for column in SENTIMENT_OUTPUT_COLUMNS if column in output_df.columns]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in SENTIMENT_OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(lambda key, column=column: stats_by_homepage.get(key, {}).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    sentiment_accounts = int(output_df[list(SENTIMENT_OUTPUT_COLUMNS)].apply(lambda row: any(_is_non_empty_cell(value) for value in row), axis=1).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "dictionary_rows": len(dictionary_df),
        "matched_posts": len(matched_work),
        "sentiment_accounts": sentiment_accounts,
    }


def calculate_stance_tendency_excel(
    account_input_path: Path,
    post_input_path: Path,
    dictionary_input_path: Path,
    output_path: Path,
    passwords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    dictionary_sheet_name: str | int = 0,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"贴文文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取贴文 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取立场倾向字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"贴文 Excel 缺少必要列：{', '.join(missing_posts)}")
    if len(dictionary_df.columns) < 1 or str(dictionary_df.columns[0]).strip() != "帖子正文":
        raise ValueError("字典 Excel 第一列必须是：帖子正文")
    if "两岸议题立场倾向" not in dictionary_df.columns:
        raise ValueError("字典 Excel 缺少必要列：两岸议题立场倾向")

    if progress is not None:
        progress(55, "正在按帖子正文匹配两岸议题立场倾向……")
    dictionary_work = dictionary_df.copy()
    dictionary_work["__body_key__"] = dictionary_work["帖子正文"].map(_normalized_key)
    dictionary_work["__stance__"] = dictionary_work["两岸议题立场倾向"].map(_normalized_key)
    dictionary_work = dictionary_work[
        (dictionary_work["__body_key__"] != "") & dictionary_work["__stance__"].isin(STANCE_VALUES)
    ].drop_duplicates(subset=["__body_key__"], keep="first")
    stance_by_body = dictionary_work.set_index("__body_key__")["__stance__"].to_dict()

    post_work = post_df.copy()
    post_work["__homepage_key__"] = post_work["主页url"].map(_normalized_key)
    post_work["__body_key__"] = post_work["帖子正文"].map(_normalized_key)
    post_work["__stance__"] = post_work["__body_key__"].map(stance_by_body)
    matched_work = post_work[
        (post_work["__homepage_key__"] != "") & post_work["__stance__"].map(_is_non_empty_cell)
    ].copy()

    if progress is not None:
        progress(72, "正在按账号汇总立场倾向分数、数量和占比……")
    stats_by_homepage: dict[str, dict[str, str | int]] = {}
    for homepage, homepage_rows in matched_work.groupby("__homepage_key__", sort=False):
        counts = {name: int((homepage_rows["__stance__"] == name).sum()) for name in STANCE_VALUES}
        total = sum(counts.values())
        stats: dict[str, str | int] = {
            "立场倾向-分数": sum(counts[name] * STANCE_SCORE_MAP[name] for name in STANCE_VALUES),
        }
        for name in STANCE_VALUES:
            stats[f"立场倾向-{name}（数量）"] = counts[name]
            stats[f"立场倾向-{name}（占比）"] = _format_single_percentage(counts[name], total)
        stats_by_homepage[str(homepage)] = stats

    if progress is not None:
        progress(84, "正在写回账号表立场倾向统计列……")
    output_df = account_df.copy()
    drop_columns = [column for column in STANCE_OUTPUT_COLUMNS if column in output_df.columns]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in STANCE_OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(lambda key, column=column: stats_by_homepage.get(key, {}).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    stance_accounts = int(output_df[list(STANCE_OUTPUT_COLUMNS)].apply(lambda row: any(_is_non_empty_cell(value) for value in row), axis=1).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "dictionary_rows": len(dictionary_df),
        "matched_posts": len(matched_work),
        "stance_accounts": stance_accounts,
    }
