from __future__ import annotations

import random
from pathlib import Path
from typing import Any, Callable

from jy_toolbox.core.chinese import convert_chinese_text
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

POSTING_PERIOD_OUTPUT_COLUMNS = ("高频发帖时段", "高频发帖类型")
POST_TYPE_LEGACY_OUTPUT_COLUMN = "帖子类型"
SOURCE_MEDIA_STANCE_COLUMN = "账号立场归属"
SOURCE_MEDIA_ACCOUNT_TYPE_COLUMN = "账号类型归属"
SOURCE_MEDIA_OFFICIAL_SOURCE_TYPE = "政府/军警机关"
SOURCE_MEDIA_OUTPUT_COLUMNS = ("账号立场归属占比", "账号类型归属占比", "官方信源占比")
CUSTOM_KEYWORD_FREQUENCY_PREFIX = "词频-"


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
    missing = [col for col in ["帖子url", "点赞数", "分享数", "评论数"] if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要列：{', '.join(missing)}")

    work = df.copy()
    if progress is not None:
        progress(55, "正在计算重复帖子保留规则……")
    score_cols = ["点赞数", "分享数", "评论数"]
    numeric_scores = work[score_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    work["__dedup_score__"] = numeric_scores.sum(axis=1)
    work["__dedup_random__"] = [random.random() for _ in range(len(work))]
    work["__dedup_order__"] = range(len(work))
    kept = (
        work.sort_values(["帖子url", "__dedup_score__", "__dedup_random__"], ascending=[True, False, False])
        .drop_duplicates(subset=["帖子url"], keep="first")
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

def _parse_time_seconds(value: Any) -> int | None:
    import pandas as pd

    if not _is_non_empty_cell(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        return int(value.hour) * 3600 + int(value.minute) * 60 + int(value.second)
    text = str(value).strip()
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return int(parsed.hour) * 3600 + int(parsed.minute) * 60 + int(parsed.second)

def _format_time_seconds(seconds: int) -> str:
    seconds %= 24 * 3600
    hour = seconds // 3600
    minute = (seconds % 3600) // 60
    second = seconds % 60
    return f"{hour}:{minute:02d}:{second:02d}"

def _time_in_range(seconds: int, start: int, end: int) -> bool:
    if start <= end:
        return start <= seconds <= end
    return seconds >= start or seconds <= end

def calculate_posting_period_type_excel(
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
    import pandas as pd

    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取发帖时段字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子发布时间"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
    missing_dictionary = [col for col in ["时段类型", "起始时段", "结束时段"] if col not in dictionary_df.columns]
    if missing_dictionary:
        raise ValueError(f"字典 Excel 缺少必要列：{', '.join(missing_dictionary)}")

    periods: list[dict[str, Any]] = []
    disorder_type = "混乱型"
    for _, row in dictionary_df.iterrows():
        period_type = _normalized_key(row.get("时段类型"))
        start = _parse_time_seconds(row.get("起始时段"))
        end = _parse_time_seconds(row.get("结束时段"))
        if start is None or end is None:
            if period_type:
                disorder_type = period_type
            continue
        periods.append({
            "type": period_type,
            "start": start,
            "end": end,
            "range": f"{_format_time_seconds(start)}~{_format_time_seconds(end)}",
        })
    if not periods:
        raise ValueError("字典 Excel 中没有可用的时段定义。")

    if progress is not None:
        progress(58, "正在按帖子发布时间匹配发帖时段……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__post_datetime__"] = pd.to_datetime(work["帖子发布时间"], errors="coerce", format="mixed")
    work = work[(work["__homepage_key__"] != "") & work["__post_datetime__"].notna()].copy()
    work["__seconds__"] = (
        work["__post_datetime__"].dt.hour * 3600
        + work["__post_datetime__"].dt.minute * 60
        + work["__post_datetime__"].dt.second
    )

    def match_period(seconds: int) -> str:
        for index, period in enumerate(periods):
            if _time_in_range(int(seconds), int(period["start"]), int(period["end"])):
                return str(index)
        return ""

    work["__period_index__"] = work["__seconds__"].map(match_period)
    matched_work = work[work["__period_index__"] != ""].copy()

    if progress is not None:
        progress(75, "正在按账号计算高频发帖时段类型……")
    result_by_homepage: dict[str, dict[str, str]] = {}
    for homepage, homepage_rows in matched_work.groupby("__homepage_key__", sort=False):
        counts = {int(index): 0 for index in range(len(periods))}
        for index, count in homepage_rows["__period_index__"].value_counts(sort=False).items():
            counts[int(index)] = int(count)
        total = sum(counts.values())
        if total <= 0:
            continue
        values = [counts[index] / total * 100 for index in range(len(periods))]
        is_disorder = all(counts[index] > 0 for index in range(len(periods))) and (
            max(values) - min(values) <= 10
        )
        if is_disorder:
            result_by_homepage[str(homepage)] = {"高频发帖时段": "", "高频发帖类型": disorder_type}
            continue
        top_index = max(range(len(periods)), key=lambda index: (counts[index], -index))
        top_period = periods[top_index]
        result_by_homepage[str(homepage)] = {
            "高频发帖时段": str(top_period["range"]),
            "高频发帖类型": str(top_period["type"]),
        }

    if progress is not None:
        progress(86, "正在写回账号表高频发帖时段和类型列……")
    output_df = account_df.copy()
    for column in POSTING_PERIOD_OUTPUT_COLUMNS:
        if column in output_df.columns:
            output_df = output_df.drop(columns=[column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in POSTING_PERIOD_OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(lambda key, column=column: result_by_homepage.get(key, {}).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(94, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    typed_accounts = int(output_df["高频发帖类型"].map(_is_non_empty_cell).sum())
    disorder_accounts = int((output_df["高频发帖类型"] == disorder_type).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "dictionary_rows": len(dictionary_df),
        "periods": len(periods),
        "valid_time_posts": len(work),
        "matched_posts": len(matched_work),
        "typed_accounts": typed_accounts,
        "disorder_accounts": disorder_accounts,
    }

def classify_post_type(row: Any) -> str:
    homepage_url = "" if not _is_non_empty_cell(row.get("主页url")) else str(row.get("主页url"))
    post_url = "" if not _is_non_empty_cell(row.get("帖子url")) else str(row.get("帖子url"))
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
        return {column: "" for column in ("文字帖子占比", "图片帖子占比", "视频帖子占比")}
    return {
        "文字帖子占比": f"{counts.get('文字', 0) / total * 100:.2f}%",
        "图片帖子占比": f"{counts.get('图片', 0) / total * 100:.2f}%",
        "视频帖子占比": f"{counts.get('视频', 0) / total * 100:.2f}%",
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "图片附件", "创作类型", "标题", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(50, "正在识别每条帖子的文字、图片、视频类型……")
    work = post_df.copy()
    work["__post_type__"] = work.apply(classify_post_type, axis=1)
    typed_work = work[work["__post_type__"].isin(["文字", "图片", "视频"])].copy()

    empty_ratios = {column: "" for column in ("文字帖子占比", "图片帖子占比", "视频帖子占比")}
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
        for column in (POST_TYPE_LEGACY_OUTPUT_COLUMN, *("文字帖子占比", "图片帖子占比", "视频帖子占比"))
        if column in output_df.columns
    ]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in ("文字帖子占比", "图片帖子占比", "视频帖子占比"):
        output_df[column] = account_keys.map(lambda key, column=column: ratios_by_homepage.get(key, empty_ratios).get(column, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(90, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(_normalized_key(v) for v in post_df["主页url"])).sum())
    typed_accounts = int(output_df[("文字帖子占比", "图片帖子占比", "视频帖子占比")[0]].map(_is_non_empty_cell).sum())
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在计算每条帖子正文长度……")
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



def calculate_average_original_post_interactions_excel(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    required_post_columns = ["主页url", "创作类型", "点赞数", "分享数", "评论数"]
    missing_posts = [col for col in required_post_columns if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在筛选原创帖子并计算每帖互动数……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__creation_type__"] = work["创作类型"].map(
        lambda value: str(value).strip().lower() if _is_non_empty_cell(value) else ""
    )
    original_work = work[(work["__homepage_key__"] != "") & (work["__creation_type__"] == "common")].copy()
    interaction_columns = ["点赞数", "分享数", "评论数"]
    original_work["__interactions__"] = original_work[interaction_columns].apply(
        pd.to_numeric,
        errors="coerce",
    ).fillna(0).sum(axis=1)

    if progress is not None:
        progress(74, "正在按账号计算平均原创单帖子互动数……")
    average_by_homepage: dict[str, float] = {}
    for homepage, homepage_rows in original_work.groupby("__homepage_key__", sort=False):
        average_by_homepage[str(homepage)] = round(float(homepage_rows["__interactions__"].mean()), 2)

    if progress is not None:
        progress(84, "正在写回账号表平均原创单帖子互动数列……")
    output_column = "平均原创单帖子互动数（条）"
    output_df = account_df.copy()
    if output_column in output_df.columns:
        output_df = output_df.drop(columns=[output_column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[output_column] = account_keys.map(lambda key: average_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(original_work["__homepage_key__"])).sum())
    averaged_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "original_posts": len(original_work),
        "matched_accounts": matched_accounts,
        "averaged_accounts": averaged_accounts,
    }


def calculate_average_daily_original_posts_excel(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    required_post_columns = ["主页url", "创作类型", "帖子发布时间"]
    missing_posts = [col for col in required_post_columns if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在解析帖子发布时间并筛选原创帖子……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__creation_type__"] = work["创作类型"].map(
        lambda value: str(value).strip().lower() if _is_non_empty_cell(value) else ""
    )
    work["__post_datetime__"] = pd.to_datetime(work["帖子发布时间"], errors="coerce", format="mixed")
    work = work[(work["__homepage_key__"] != "") & work["__post_datetime__"].notna()].copy()
    work["__post_date__"] = work["__post_datetime__"].dt.date
    work["__is_original__"] = work["__creation_type__"] != "share"

    if progress is not None:
        progress(74, "正在按账号计算日均原创量……")
    average_by_homepage: dict[str, float] = {}
    original_counts_by_homepage: dict[str, int] = {}
    active_days_by_homepage: dict[str, int] = {}
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        active_days = int(homepage_rows["__post_date__"].nunique())
        if active_days <= 0:
            continue
        original_count = int(homepage_rows["__is_original__"].sum())
        homepage_key = str(homepage)
        active_days_by_homepage[homepage_key] = active_days
        original_counts_by_homepage[homepage_key] = original_count
        average_by_homepage[homepage_key] = round(original_count / active_days, 2)

    if progress is not None:
        progress(84, "正在写回账号表日均原创量列……")
    output_column = "日均原创量（条）"
    output_df = account_df.copy()
    if output_column in output_df.columns:
        output_df = output_df.drop(columns=[output_column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[output_column] = account_keys.map(lambda key: average_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(work["__homepage_key__"])).sum())
    averaged_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "valid_time_posts": len(work),
        "original_posts": sum(original_counts_by_homepage.values()),
        "matched_accounts": matched_accounts,
        "averaged_accounts": averaged_accounts,
    }



def calculate_original_post_ratio_excel(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    required_post_columns = ["主页url", "创作类型"]
    missing_posts = [col for col in required_post_columns if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在按创作类型识别原创和转发帖子……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__creation_type__"] = work["创作类型"].map(
        lambda value: str(value).strip().lower() if _is_non_empty_cell(value) else ""
    )
    classified_work = work[
        (work["__homepage_key__"] != "")
        & (work["__creation_type__"].isin(["common", "share"]))
    ].copy()
    classified_work["__is_original__"] = classified_work["__creation_type__"] == "common"

    if progress is not None:
        progress(72, "正在按账号汇总原创帖子占比……")
    ratios_by_homepage: dict[str, str] = {}
    original_posts = 0
    for homepage, homepage_rows in classified_work.groupby("__homepage_key__", sort=False):
        original_count = int(homepage_rows["__is_original__"].sum())
        original_posts += original_count
        ratios_by_homepage[str(homepage)] = _format_single_percentage(original_count, len(homepage_rows))

    if progress is not None:
        progress(84, "正在写回账号表原创帖子占比列……")
    output_column = "原创帖子占比"
    output_df = account_df.copy()
    if output_column in output_df.columns:
        output_df = output_df.drop(columns=[output_column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[output_column] = account_keys.map(lambda key: ratios_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(classified_work["__homepage_key__"])).sum())
    rated_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "classified_posts": len(classified_work),
        "original_posts": original_posts,
        "matched_accounts": matched_accounts,
        "rated_accounts": rated_accounts,
    }

def calculate_weekly_post_frequency_excel(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    required_post_columns = ["主页url", "帖子发布时间"]
    missing_posts = [col for col in required_post_columns if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在解析帖子发布时间并统计账号发帖频率……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__post_datetime__"] = pd.to_datetime(work["帖子发布时间"], errors="coerce", format="mixed")
    work = work[(work["__homepage_key__"] != "") & work["__post_datetime__"].notna()].copy()

    if progress is not None:
        progress(74, "正在按账号计算每周发布帖子频率……")
    span_weeks_by_homepage: dict[str, int] = {}
    frequency_by_homepage: dict[str, float] = {}
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        first_date = homepage_rows["__post_datetime__"].min().date()
        last_date = homepage_rows["__post_datetime__"].max().date()
        span_days = (last_date - first_date).days + 1
        if span_days <= 0:
            continue
        # Use inclusive calendar weeks instead of exact seconds between the
        # first and last post. Exact-second spans can be only a few minutes for
        # accounts with one burst of posts, which inflates the weekly rate into
        # thousands. A minimum one-week denominator keeps short observation
        # windows interpretable (e.g. one post in the data means 1 time/week).
        span_weeks = max(1, (span_days + 6) // 7)
        span_weeks_by_homepage[str(homepage)] = span_weeks
        frequency_by_homepage[str(homepage)] = round(len(homepage_rows) / span_weeks, 2)

    if progress is not None:
        progress(84, "正在写回账号表跨越周数和每周发布帖子频率列……")
    span_weeks_column = "跨越周数"
    output_column = "每周发布帖子频率（次）"
    output_df = account_df.copy()
    existing_output_columns = [col for col in [span_weeks_column, output_column] if col in output_df.columns]
    if existing_output_columns:
        output_df = output_df.drop(columns=existing_output_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[span_weeks_column] = account_keys.map(lambda key: span_weeks_by_homepage.get(key, ""))
    output_df[output_column] = account_keys.map(lambda key: frequency_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(93, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(work["__homepage_key__"])).sum())
    calculated_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "valid_time_posts": len(work),
        "matched_accounts": matched_accounts,
        "calculated_accounts": calculated_accounts,
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子发布时间"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在解析帖子发布时间并筛选单日 2 条及以上的发帖日期……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work["__post_datetime__"] = pd.to_datetime(work["帖子发布时间"], errors="coerce")
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    import pandas as pd

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子发布时间"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在统计帖子数量并解析帖子发布时间……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    post_counts_by_homepage = (
        work[work["__homepage_key__"] != ""].groupby("__homepage_key__", sort=False).size().to_dict()
    )
    work["__post_datetime__"] = pd.to_datetime(work["帖子发布时间"], errors="coerce")
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "创作类型", "标题", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    if progress is not None:
        progress(55, "正在识别转发帖子是否附加标题或帖子正文……")
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取账号名字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError(f"账号 Excel 缺少必要列：{"FB主页"}")
    missing_posts = [col for col in ["主页url", "创作类型", "分享帖账号名"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
    missing_dictionary = [col for col in ["分享帖账号名", "账号立场归属", "账号类型归属"] if col not in dictionary_df.columns]
    if missing_dictionary:
        raise ValueError(f"字典 Excel 缺少必要列：{', '.join(missing_dictionary)}")

    if progress is not None:
        progress(55, "正在按字典匹配分享帖账号名……")
    dictionary_work = dictionary_df.copy()
    dictionary_work["__shared_account_key__"] = dictionary_work["分享帖账号名"].map(
        _normalized_key
    )
    dictionary_work = dictionary_work[dictionary_work["__shared_account_key__"] != ""].drop_duplicates(
        subset=["__shared_account_key__"],
        keep="first",
    )
    stance_by_shared_account = dictionary_work.set_index("__shared_account_key__")[
        SOURCE_MEDIA_STANCE_COLUMN
    ].to_dict()
    type_by_shared_account = dictionary_work.set_index("__shared_account_key__")[
        SOURCE_MEDIA_ACCOUNT_TYPE_COLUMN
    ].to_dict()

    post_work = post_df.copy()
    post_work["__homepage_key__"] = post_work["主页url"].map(_normalized_key)
    post_work["__shared_account_key__"] = post_work["分享帖账号名"].map(
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
        official_source_count = type_counts.get(SOURCE_MEDIA_OFFICIAL_SOURCE_TYPE, 0)
        total_typed_sources = sum(type_counts.values())
        ratios_by_homepage[str(homepage)] = {
            "账号立场归属占比": _format_category_ratios(stance_counts),
            "账号类型归属占比": _format_category_ratios(type_counts),
            "官方信源占比": _format_single_percentage(official_source_count, total_typed_sources),
        }

    if progress is not None:
        progress(82, "正在写回账号表最后三列媒体阵营分布……")
    output_df = account_df.copy()
    drop_columns = [column for column in SOURCE_MEDIA_OUTPUT_COLUMNS if column in output_df.columns]
    if drop_columns:
        output_df = output_df.drop(columns=drop_columns)
    account_keys = output_df["FB主页"].map(_normalized_key)
    for column in SOURCE_MEDIA_OUTPUT_COLUMNS:
        output_df[column] = account_keys.map(
            lambda key, column=column: ratios_by_homepage.get(key, {}).get(column, "")
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(92, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    classified_accounts = int(
        output_df[list(SOURCE_MEDIA_OUTPUT_COLUMNS)].apply(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取内容偏好字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError('账号 Excel 缺少必要列：FB主页')
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
    if len(dictionary_df.columns) < 1 or str(dictionary_df.columns[0]).strip() != "帖子正文":
        raise ValueError("字典 Excel 第一列必须是：帖子正文")
    if "内容偏好" not in dictionary_df.columns:
        raise ValueError("字典 Excel 缺少必要列：内容偏好")

    if progress is not None:
        progress(55, "正在按帖子正文匹配内容偏好分组……")
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取情感表达字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取立场倾向字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
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

class _KeywordMatcher:
    def __init__(self, keywords: list[str]) -> None:
        self._next: list[dict[str, int]] = [{}]
        self._fail: list[int] = [0]
        self._terminal: list[bool] = [False]
        for keyword in keywords:
            state = 0
            for char in keyword:
                nxt = self._next[state].get(char)
                if nxt is None:
                    nxt = len(self._next)
                    self._next[state][char] = nxt
                    self._next.append({})
                    self._fail.append(0)
                    self._terminal.append(False)
                state = nxt
            self._terminal[state] = True

        from collections import deque

        queue: deque[int] = deque()
        for child in self._next[0].values():
            queue.append(child)
        while queue:
            current = queue.popleft()
            if self._terminal[self._fail[current]]:
                self._terminal[current] = True
            for char, child in self._next[current].items():
                fail_state = self._fail[current]
                while fail_state and char not in self._next[fail_state]:
                    fail_state = self._fail[fail_state]
                self._fail[child] = self._next[fail_state].get(char, 0)
                queue.append(child)

    def contains(self, text: Any) -> bool:
        if not _is_non_empty_cell(text):
            return False
        state = 0
        for char in str(text):
            while state and char not in self._next[state]:
                state = self._fail[state]
            state = self._next[state].get(char, 0)
            if self._terminal[state]:
                return True
        return False


def _load_first_column_keywords(dictionary_df: Any) -> list[str]:
    if len(dictionary_df.columns) < 1:
        return []
    first_column = dictionary_df.columns[0]
    seen: set[str] = set()
    keywords: list[str] = []
    for value in dictionary_df[first_column]:
        keyword = _normalized_key(value)
        if keyword and keyword not in seen:
            seen.add(keyword)
            keywords.append(keyword)
    return keywords


def _combine_post_text(row: Any) -> str:
    parts = []
    for column in ("标题", "帖子正文"):
        if column in row and _is_non_empty_cell(row.get(column)):
            parts.append(str(row.get(column)))
    return "\n".join(parts)


def calculate_custom_keyword_frequencies_excel(
    account_input_path: Path,
    post_input_path: Path,
    output_path: Path,
    passwords: list[str],
    keywords: list[str],
    account_sheet_name: str | int = 0,
    post_sheet_name: str | int = 0,
    ignore_chinese_script: bool = False,
    progress: Callable[[float, str | None], None] | None = None,
) -> dict[str, int]:
    if not account_input_path.exists():
        raise FileNotFoundError(f"账号文件不存在：{account_input_path}")
    if not post_input_path.exists():
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")

    normalized_keywords = list(dict.fromkeys(_normalized_key(keyword) for keyword in keywords))
    normalized_keywords = [keyword for keyword in normalized_keywords if keyword]
    if not normalized_keywords:
        raise ValueError("请至少新增 1 个可用关键词。")
    keyword_variants = {
        keyword: list(dict.fromkeys((
            keyword,
            convert_chinese_text(keyword, to_traditional=True),
            convert_chinese_text(keyword, to_traditional=False),
        )))
        if ignore_chinese_script
        else [keyword]
        for keyword in normalized_keywords
    }

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(30, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")
    if "标题" not in post_df.columns and "帖子正文" not in post_df.columns:
        raise ValueError("帖子 Excel 至少需要包含“标题”或“帖子正文”列。")

    if progress is not None:
        progress(55, f"正在扫描帖子并统计 {len(normalized_keywords)} 个自定义关键词……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work = work[work["__homepage_key__"] != ""].copy()
    work["__post_text__"] = work.apply(_combine_post_text, axis=1)

    counts_by_homepage: dict[str, dict[str, int]] = {}
    for homepage, homepage_rows in work.groupby("__homepage_key__", sort=False):
        joined_text = "\n".join(str(value) for value in homepage_rows["__post_text__"] if _is_non_empty_cell(value))
        counts_by_homepage[str(homepage)] = {
            keyword: int(sum(joined_text.count(variant) for variant in keyword_variants[keyword]))
            for keyword in normalized_keywords
        }

    if progress is not None:
        progress(82, "正在写回账号表自定义关键词词频列……")
    output_df = account_df.copy()
    output_columns = [f"{CUSTOM_KEYWORD_FREQUENCY_PREFIX}{keyword}" for keyword in normalized_keywords]
    for column in output_columns:
        if column in output_df.columns:
            output_df = output_df.drop(columns=[column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    for keyword, column in zip(normalized_keywords, output_columns, strict=True):
        output_df[column] = account_keys.map(lambda key, keyword=keyword: counts_by_homepage.get(key, {}).get(keyword, 0))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(94, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(
        output_df[output_columns].apply(lambda row: any(int(value) > 0 for value in row), axis=1).sum()
    )
    total_occurrences = int(output_df[output_columns].sum(numeric_only=True).sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "keywords": len(normalized_keywords),
        "matched_accounts": matched_accounts,
        "total_occurrences": total_occurrences,
    }


def calculate_sensitive_topic_participation_rate_excel(
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
        raise FileNotFoundError(f"帖子文件不存在：{post_input_path}")
    if not dictionary_input_path.exists():
        raise FileNotFoundError(f"字典文件不存在：{dictionary_input_path}")

    if progress is not None:
        progress(10, "正在后台读取账号 Excel……")
    account_df = read_excel_with_passwords(account_input_path, passwords, account_sheet_name)
    if progress is not None:
        progress(25, "正在后台读取帖子 Excel……")
    post_df = read_excel_with_passwords(post_input_path, passwords, post_sheet_name)
    if progress is not None:
        progress(40, "正在后台读取敏感话题关键词字典 Excel……")
    dictionary_df = read_excel_with_passwords(dictionary_input_path, passwords, dictionary_sheet_name, header=None)

    if "FB主页" not in account_df.columns:
        raise ValueError("账号 Excel 缺少必要列：FB主页")
    missing_posts = [col for col in ["主页url", "帖子正文"] if col not in post_df.columns]
    if missing_posts:
        raise ValueError(f"帖子 Excel 缺少必要列：{', '.join(missing_posts)}")

    keywords = _load_first_column_keywords(dictionary_df)
    if not keywords:
        raise ValueError("字典 Excel 第一列没有可用关键词。")

    if progress is not None:
        progress(55, f"正在构建 {len(keywords)} 个关键词的快速匹配索引……")
    matcher = _KeywordMatcher(keywords)

    if progress is not None:
        progress(65, "正在扫描帖子标题和正文中的敏感话题关键词……")
    work = post_df.copy()
    work["__homepage_key__"] = work["主页url"].map(_normalized_key)
    work = work[work["__homepage_key__"] != ""].copy()
    work["__post_text__"] = work.apply(_combine_post_text, axis=1)
    work["__is_sensitive_topic__"] = work["__post_text__"].map(matcher.contains)

    if progress is not None:
        progress(78, "正在按账号汇总敏感话题参与率……")
    total_counts_by_homepage = work.groupby("__homepage_key__", sort=False).size().to_dict()
    sensitive_counts_by_homepage = (
        work[work["__is_sensitive_topic__"]].groupby("__homepage_key__", sort=False).size().to_dict()
    )
    ratios_by_homepage = {
        str(homepage): f"{int(sensitive_counts_by_homepage.get(homepage, 0)) / int(total) * 100:.2f}%"
        for homepage, total in total_counts_by_homepage.items()
        if int(total) > 0
    }

    if progress is not None:
        progress(86, "正在写回账号表敏感话题参与率列……")
    output_column = "敏感话题参与率（%）"
    output_df = account_df.copy()
    if output_column in output_df.columns:
        output_df = output_df.drop(columns=[output_column])
    account_keys = output_df["FB主页"].map(_normalized_key)
    output_df[output_column] = account_keys.map(lambda key: ratios_by_homepage.get(key, ""))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress(94, "正在保存处理后的账号 Excel……")
    output_df.to_excel(output_path, index=False)

    matched_accounts = int(account_keys.isin(set(total_counts_by_homepage)).sum())
    rated_accounts = int(output_df[output_column].map(_is_non_empty_cell).sum())
    sensitive_posts = int(work["__is_sensitive_topic__"].sum())
    return {
        "accounts": len(account_df),
        "posts": len(post_df),
        "dictionary_keywords": len(keywords),
        "matched_accounts": matched_accounts,
        "sensitive_posts": sensitive_posts,
        "rated_accounts": rated_accounts,
    }
