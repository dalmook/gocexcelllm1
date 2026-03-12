from __future__ import annotations

import os
import re
import json
import sys
import uuid
import math
import difflib
import poplib
import html as html_lib
from datetime import datetime, date, timedelta
from dataclasses import dataclass, field
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
from tkinter import scrolledtext
from email import policy
from email.header import decode_header
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime

import requests
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


# =========================================================
# 1) GPT-OSS API 설정
# =========================================================

API_BASE_URL = os.getenv(
    "API_BASE_URL",
    "http://apigw.samsungds.net:8000/gpt-oss/1/gpt-oss-120b/v1/chat/completions",
)

CREDENTIAL_KEY = os.getenv("CREDENTIAL_KEY", "credential:TICKET-96f7bce0-efab-4516-8e62-5501b07ab43c:ST0000107488-PROD:CTXLCkSDRGWtI5HdVHkPAQgol2o-RyQiq2I1vCHHOgGw:-1:Q1RYTENrU0RSR1d0STVIZFZIa1BBUWdvbDJvLVJ5UWlxMkkxdkNISE9nR3c=:signature=eRa1UcfmWGfKTDBt-Xnz2wFhW0OvMX0WESZUpoNVgCA5uNVgpgax59LZ3osPOp8whnZwQay8s5TUvxJGtmsCD9iK-HpcsyUOcE5P58W0Weyg-YQ3KRTWFiA==")
USER_ID = os.getenv("USER_ID", "sungmook.cho")
USER_TYPE = os.getenv("USER_TYPE", "AD_ID")
SEND_SYSTEM_NAME = os.getenv("SEND_SYSTEM_NAME", "GOC_MAIL_RAG_PIPELINE")


# =========================================================
# 2) LLM 호출
# =========================================================

def call_gpt_oss(
    prompt: str,
    system_prompt: str | None = None,
    temperature: float = 0.2,
    max_tokens: int = 900,
) -> dict:
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    payload = json.dumps({
        "model": "openai/gpt-oss-120b",
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False
    })

    headers = {
        "x-dep-ticket": CREDENTIAL_KEY,
        "Send-System-Name": SEND_SYSTEM_NAME,
        "User-Id": USER_ID,
        "User-Type": USER_TYPE,
        "Prompt-Msg-Id": str(uuid.uuid4()),
        "Completion-Msg-Id": str(uuid.uuid4()),
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    try:
        response = requests.post(
            API_BASE_URL,
            headers=headers,
            data=payload,
            timeout=60,
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}


def extract_llm_text(result: dict) -> str:
    if "error" in result:
        return f"[LLM 오류]\n{result['error']}"
    try:
        return result["choices"][0]["message"]["content"]
    except Exception:
        return f"[LLM 응답 파싱 실패]\n{json.dumps(result, ensure_ascii=False, indent=2)}"


# =========================================================
# 3) 공용 유틸
# =========================================================
def make_json_safe(obj):
    """
    datetime/date 등 JSON 직렬화 불가 객체를 문자열 등으로 변환
    """
    if isinstance(obj, datetime):
        return obj.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(obj, date):
        return obj.strftime("%Y-%m-%d")
    if isinstance(obj, dict):
        return {k: make_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [make_json_safe(v) for v in obj]
    if isinstance(obj, tuple):
        return [make_json_safe(v) for v in obj]
    return obj
def normalize_cell_value(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, str):
        v = value.strip()
        if v == "":
            return None
        v2 = v.replace(",", "")
        if re.fullmatch(r"-?\d+(\.\d+)?", v2):
            try:
                if "." in v2:
                    return float(v2)
                return int(v2)
            except Exception:
                return v
        return v

    return value


def clean_header_name(v, idx: int) -> str:
    if v is None:
        return f"COL_{idx}"
    s = str(v).strip()
    if not s:
        return f"COL_{idx}"
    return s


def safe_float(value):
    try:
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, (int, float)):
            if isinstance(value, float) and math.isnan(value):
                return None
            return float(value)
        if isinstance(value, str):
            v = value.strip().replace(",", "")
            if re.fullmatch(r"-?\d+(\.\d+)?", v):
                return float(v)
        return None
    except Exception:
        return None


def format_number(x, digits: int = 2) -> str:
    if x is None:
        return "-"
    if abs(x - int(x)) < 1e-9:
        return f"{int(x):,}"
    return f"{x:,.{digits}f}"


@dataclass
class AnalysisConfig:
    selected_label_column: str | None = None
    selected_time_column: str | None = None
    selected_metric_columns: list[str] = field(default_factory=list)
    exclude_total_rows: bool = True
    apply_merge_candidates: bool = False
    use_first_row_as_header: bool = True


def get_default_analysis_config() -> AnalysisConfig:
    return AnalysisConfig()


@dataclass
class MailQueryParams:
    user_id: str
    password: str
    keyword: str = ""
    recent_days: int = 7
    max_count: int = 20
    host: str = os.getenv("POP3_HOST", "pop3.samsung.net")
    port: int = int(os.getenv("POP3_PORT", "995"))
    use_ssl: bool = os.getenv("POP3_USE_SSL", "1") != "0"


@dataclass
class MailItem:
    index: int
    uid: str
    subject: str
    sender: str
    date_str: str
    date_obj: datetime | None
    body: str


MONTH_TOKEN_PATTERN = re.compile(
    r"^("
    r"\d{6}|"
    r"\d{4}[-/]\d{1,2}|"
    r"\d{2}년\s?\d{1,2}월|"
    r"\d{1,2}월|"
    r"jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec"
    r")$",
    re.IGNORECASE,
)

DATE_TOKEN_PATTERN = re.compile(
    r"^("
    r"\d{4}[-/]\d{1,2}[-/]\d{1,2}|"
    r"\d{4}\.\d{1,2}\.\d{1,2}|"
    r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|"
    r"\d{4}년\s?\d{1,2}월\s?\d{1,2}일"
    r")$",
    re.IGNORECASE,
)

TOTAL_ROW_PATTERN = re.compile(r"(합계|총계|subtotal|sub-total|grand total|total|ttl)", re.IGNORECASE)
ID_HEADER_PATTERN = re.compile(r"(?:^|[_\s-])(id|code|no|part|lot|serial|model)(?:$|[_\s-])", re.IGNORECASE)
PERCENT_HEADER_PATTERN = re.compile(r"(%|비율|rate|ratio|점유율|pct)", re.IGNORECASE)
LABEL_HEADER_PATTERN = re.compile(
    r"(품목|제품|모델|지역|부서|구분|name|item|category|group|라인|고객|site|plant|team)",
    re.IGNORECASE,
)


def get_analysis_thresholds() -> dict:
    return {
        "missing_warning_ratio": 0.3,
        "critical_missing_warning_ratio": 0.5,
        "share_concentration_ratio": 0.6,
        "category_skew_ratio": 0.7,
        "negative_ratio_warning": 0.3,
        "low_variance_ratio": 0.9,
        "extreme_skew_ratio": 3.0,
        "typo_similarity_ratio": 0.84,
        "merge_similarity_ratio": 0.74,
        "rare_text_count": 1,
        "min_text_count_for_typo": 1,
        "min_length_text_normal": 2,
    }


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def normalize_text_token(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def simplify_text_token(value) -> str:
    token = normalize_text_token(value)
    return re.sub(r"[\s\-_./]+", "", token)


def is_missing_value(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_percent_like_value(value) -> bool:
    if isinstance(value, str) and "%" in value:
        return True
    number = safe_float(value)
    return number is not None and 0 <= number <= 1


def is_month_like_value(value) -> bool:
    if value is None:
        return False
    if isinstance(value, (datetime, date)):
        return False
    token = normalize_text_token(value)
    if not token:
        return False
    return bool(MONTH_TOKEN_PATTERN.fullmatch(token))


def is_date_like_value(value) -> bool:
    if value is None:
        return False
    if isinstance(value, (datetime, date)):
        return True
    token = normalize_text_token(value)
    if not token:
        return False
    if is_month_like_value(value):
        return False
    if DATE_TOKEN_PATTERN.fullmatch(token):
        return True
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y"):
        try:
            datetime.strptime(token, fmt)
            return True
        except ValueError:
            continue
    return False


def is_id_like_token(value) -> bool:
    token = normalize_text_token(value)
    if not token:
        return False
    if len(token) < 3:
        return False
    if " " in token:
        return False
    has_digit = any(ch.isdigit() for ch in token)
    has_alpha = any(ch.isalpha() for ch in token)
    return has_digit or ("-" in token and has_alpha)


def detect_total_row_indexes(rows: list[dict], headers: list[str]) -> list[int]:
    total_indexes = []
    scan_headers = headers[: min(4, len(headers))]
    for idx, row in enumerate(rows):
        for header in scan_headers:
            if TOTAL_ROW_PATTERN.search(normalize_text_token(row.get(header))):
                total_indexes.append(idx)
                break
    return total_indexes


def guess_table_topic(headers: list[str]) -> str:
    joined = " ".join(normalize_text_token(h) for h in headers)
    if re.search(r"(매출|금액|이익|원가|revenue|sales|profit|cost|amount)", joined):
        return "financial"
    if re.search(r"(qty|수량|재고|inventory|stock)", joined):
        return "quantity_inventory"
    if re.search(r"(수율|yield|불량|defect|quality|ppm)", joined):
        return "quality"
    if re.search(r"(일정|계획|납기|date|schedule|month|week)", joined):
        return "schedule"
    return "general"


def infer_column_type(header: str, values: list, row_count: int) -> str:
    non_null_values = [v for v in values if not is_missing_value(v)]
    if not non_null_values:
        return "text"

    header_token = normalize_text_token(header)
    numeric_values = [safe_float(v) for v in non_null_values]
    numeric_values = [v for v in numeric_values if v is not None]
    unique_count = len({str(v) for v in non_null_values})
    unique_ratio = unique_count / len(non_null_values) if non_null_values else 0

    month_like_count = sum(1 for v in non_null_values if is_month_like_value(v))
    date_like_count = sum(1 for v in non_null_values if is_date_like_value(v))
    percent_like_count = sum(1 for v in non_null_values if is_percent_like_value(v))
    id_like_count = sum(1 for v in non_null_values if is_id_like_token(v))

    if month_like_count >= max(2, int(len(non_null_values) * 0.6)):
        return "month_like"
    if date_like_count >= max(2, int(len(non_null_values) * 0.6)):
        return "date_like"
    numeric_ratio = len(numeric_values) / len(non_null_values) if non_null_values else 0

    if PERCENT_HEADER_PATTERN.search(header) or percent_like_count >= max(2, int(len(non_null_values) * 0.7)):
        return "percent_like"
    if numeric_values and numeric_ratio >= 0.6:
        return "numeric"
    if ID_HEADER_PATTERN.search(header) or (
        unique_ratio >= 0.85 and id_like_count >= max(2, int(len(non_null_values) * 0.5))
    ):
        return "id_like"
    if unique_ratio >= 0.95 and row_count >= 5 and ID_HEADER_PATTERN.search(f" {header_token} "):
        return "id_like"
    return "text"


def choose_label_column(headers: list[str], rows: list[dict], detected_types: dict[str, str]) -> str | None:
    candidates = []
    for header in headers:
        col_type = detected_types.get(header)
        if col_type == "id_like":
            continue
        if col_type not in {"text", "date_like", "month_like"}:
            continue
        values = [row.get(header) for row in rows if not is_missing_value(row.get(header))]
        if not values:
            continue
        unique_ratio = len({str(v) for v in values}) / len(values)
        score = 0
        if LABEL_HEADER_PATTERN.search(header):
            score += 5
        if 0.2 <= unique_ratio <= 0.95:
            score += 2
        if col_type == "text":
            score += 2
        if len(header) <= 20:
            score += 1
        candidates.append((score, header))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], headers.index(item[1])))
    return candidates[0][1]


def build_value_point(header: str, row_index: int, value: float, row: dict, label_column: str | None) -> dict:
    point = {"row_index": row_index, "value": value}
    if label_column and row.get(label_column) is not None:
        point["label_column"] = label_column
        point["label_value"] = row.get(label_column)
    return point


def build_numeric_profile(header: str, values: list, data_rows: list[dict], data_indexes: list[int], label_column: str | None, top_n: int) -> dict:
    numeric_points = []
    for local_idx, value in enumerate(values):
        number = safe_float(value)
        if number is None:
            continue
        row = data_rows[local_idx]
        row_index = data_indexes[local_idx]
        numeric_points.append(build_value_point(header, row_index, number, row, label_column))

    numeric_vals = [point["value"] for point in numeric_points]
    total = sum(numeric_vals) if numeric_vals else None
    avg = (total / len(numeric_vals)) if numeric_vals else None
    min_v = min(numeric_vals) if numeric_vals else None
    max_v = max(numeric_vals) if numeric_vals else None

    median = None
    if numeric_vals:
        sorted_vals = sorted(numeric_vals)
        n = len(sorted_vals)
        if n % 2 == 1:
            median = sorted_vals[n // 2]
        else:
            median = (sorted_vals[n // 2 - 1] + sorted_vals[n // 2]) / 2

    top_values = sorted(numeric_points, key=lambda point: (-point["value"], point["row_index"]))[:top_n]
    bottom_values = sorted(numeric_points, key=lambda point: (point["value"], point["row_index"]))[:top_n]

    return {
        "count": len(numeric_vals),
        "null_count": sum(1 for v in values if is_missing_value(v)),
        "null_ratio": (sum(1 for v in values if is_missing_value(v)) / len(values)) if values else 0,
        "sum": total,
        "avg": avg,
        "min": min_v,
        "max": max_v,
        "median": median,
        "value_points": numeric_points,
        "top_values": top_values,
        "bottom_values": bottom_values,
    }


def build_text_profile(values: list, top_n: int) -> dict:
    freq = {}
    for value in values:
        if is_missing_value(value):
            continue
        key = str(value)
        freq[key] = freq.get(key, 0) + 1

    top_values = sorted(freq.items(), key=lambda item: (-item[1], item[0]))[:top_n]
    null_count = sum(1 for v in values if is_missing_value(v))
    return {
        "count": len(values) - null_count,
        "null_count": null_count,
        "null_ratio": (null_count / len(values)) if values else 0,
        "unique_count": len(freq),
        "top_values": [{"value": key, "count": count} for key, count in top_values],
        "value_counts": {key: count for key, count in sorted(freq.items(), key=lambda item: (-item[1], item[0]))[:top_n]},
    }


def analyze_categorical_skew(header: str, profile: dict, row_count: int) -> str | None:
    thresholds = get_analysis_thresholds()
    if row_count <= 0 or profile["unique_count"] == 0:
        return None
    if profile["unique_count"] > max(12, int(row_count * 0.5)):
        return None
    top_values = profile.get("top_values") or []
    if not top_values:
        return None
    top_count = top_values[0]["count"]
    ratio = top_count / row_count
    if ratio >= thresholds["category_skew_ratio"]:
        return f"{header} 컬럼은 '{top_values[0]['value']}'가 {ratio:.0%}로 편중되어 있습니다."
    return None


def infer_value_scale_guess(header: str, column_type: str) -> str:
    token = normalize_text_token(header)
    if re.search(r"(qty|수량|count|ea|건수)", token):
        return "count"
    if re.search(r"(amt|amount|cost|price|sales|매출|금액|원가|이익|revenue|profit)", token):
        return "amount"
    if column_type == "percent_like" or re.search(r"(%|비율|rate|yield|ratio)", token):
        return "percent"
    return "unknown"


def choose_category_columns(headers: list[str], detected_types: dict[str, str], column_profiles: dict, label_column: str | None) -> list[str]:
    candidates = []
    for header in headers:
        col_type = detected_types.get(header)
        if col_type != "text":
            continue
        profile = column_profiles.get(header, {})
        unique_count = profile.get("unique_count", 0)
        if unique_count == 0 or unique_count > 12:
            continue
        score = 0
        if header == label_column:
            score += 4
        if LABEL_HEADER_PATTERN.search(header):
            score += 3
        if 2 <= unique_count <= 8:
            score += 2
        candidates.append((score, header))
    candidates.sort(key=lambda item: (-item[0], headers.index(item[1])))
    return [header for _, header in candidates[:2]]


def analyze_numeric_quality(header: str, profile: dict, value_scale_guess: str) -> list[str]:
    thresholds = get_analysis_thresholds()
    warnings = []
    value_points = profile.get("value_points", [])
    numeric_vals = [safe_float(point.get("value")) for point in value_points]
    numeric_vals = [value for value in numeric_vals if value is not None]
    if not numeric_vals:
        return warnings

    negative_ratio = sum(1 for value in numeric_vals if value < 0) / len(numeric_vals)
    if negative_ratio >= thresholds["negative_ratio_warning"]:
        warnings.append(f"{header} 컬럼은 음수 비중이 {negative_ratio:.0%}로 높습니다.")

    most_common_ratio = max(numeric_vals.count(value) for value in set(numeric_vals)) / len(numeric_vals)
    if most_common_ratio >= thresholds["low_variance_ratio"]:
        warnings.append(f"{header} 컬럼은 값이 거의 동일하게 반복됩니다.")

    median = safe_float(profile.get("median"))
    max_value = safe_float(profile.get("max"))
    if median not in (None, 0) and max_value is not None and abs(max_value) / abs(median) >= thresholds["extreme_skew_ratio"]:
        warnings.append(f"{header} 컬럼은 최대값이 중앙값 대비 과도하게 커서 extreme skew 가능성이 있습니다.")

    if profile.get("null_ratio", 0) >= thresholds["missing_warning_ratio"] and value_scale_guess != "unknown":
        warnings.append(f"{header} 핵심 metric의 결측률이 {profile['null_ratio']:.0%}로 높습니다.")
    return warnings


def build_major_metrics(numeric_columns: list[str], column_profiles: dict) -> list[dict]:
    metrics = []
    for header in numeric_columns:
        profile = column_profiles.get(header, {})
        metrics.append({
            "column": header,
            "sum": profile.get("sum"),
            "avg": profile.get("avg"),
            "min": profile.get("min"),
            "max": profile.get("max"),
            "median": profile.get("median"),
            "null_ratio": profile.get("null_ratio"),
            "value_scale_guess": infer_value_scale_guess(header, profile.get("type", "unknown")),
        })
    return metrics


def build_total_row_metrics(total_rows: list[dict], numeric_columns: list[str], label_column: str | None) -> tuple[list[dict], list[dict]]:
    total_row_metrics = []
    detected_totals = []
    for row in total_rows:
        metric_values = {}
        for column in numeric_columns:
            value = safe_float(row.get(column))
            if value is not None:
                metric_values[column] = value
        total_row_metrics.append(metric_values)
        detected_totals.append({
            "label": row.get(label_column) if label_column else None,
            "values": metric_values,
        })
    return total_row_metrics, detected_totals


def build_category_share_entry(category_column: str, metric_column: str, bucket_sums: dict, top_n: int) -> dict:
    thresholds = get_analysis_thresholds()
    total_sum = sum(bucket_sums.values())
    sorted_items = sorted(bucket_sums.items(), key=lambda item: (-item[1], item[0]))
    top_categories = []
    for name, value in sorted_items[:top_n]:
        top_categories.append({
            "category": name,
            "sum": value,
            "share": safe_divide(value, total_sum),
        })
    top_category_share = top_categories[0]["share"] if top_categories else None
    return {
        "category_column": category_column,
        "metric_column": metric_column,
        "top_categories_by_sum": top_categories,
        "top_category_share": top_category_share,
        "concentration_warning": bool(top_category_share is not None and top_category_share >= thresholds["share_concentration_ratio"]),
    }


def build_category_analyses(category_columns: list[str], numeric_columns: list[str], rows: list[dict], top_n: int) -> tuple[list[dict], list[dict], list[dict], list[str]]:
    warnings = []
    share_analyses = []
    top_performers = []
    bottom_performers = []

    for category_column in category_columns:
        for metric_column in numeric_columns:
            bucket_sums = {}
            for row in rows:
                category = row.get(category_column)
                metric_value = safe_float(row.get(metric_column))
                if is_missing_value(category) or metric_value is None:
                    continue
                bucket_key = str(category)
                bucket_sums[bucket_key] = bucket_sums.get(bucket_key, 0.0) + metric_value

            if len(bucket_sums) < 2:
                continue

            share_entry = build_category_share_entry(category_column, metric_column, bucket_sums, top_n)
            share_analyses.append(share_entry)
            if share_entry["concentration_warning"]:
                warnings.append(
                    f"{category_column} 기준 {metric_column} 비중에서 상위 1개 범주가 {share_entry['top_category_share']:.0%}를 차지합니다."
                )

            total_sum = sum(bucket_sums.values())
            sorted_desc = sorted(bucket_sums.items(), key=lambda item: (-item[1], item[0]))
            sorted_asc = sorted(bucket_sums.items(), key=lambda item: (item[1], item[0]))
            top_performers.append({
                "category_column": category_column,
                "metric_column": metric_column,
                "category": sorted_desc[0][0],
                "value": sorted_desc[0][1],
                "share": safe_divide(sorted_desc[0][1], total_sum),
            })
            bottom_performers.append({
                "category_column": category_column,
                "metric_column": metric_column,
                "category": sorted_asc[0][0],
                "value": sorted_asc[0][1],
                "share": safe_divide(sorted_asc[0][1], total_sum),
            })

    top_performers.sort(key=lambda item: abs(item["value"]), reverse=True)
    bottom_performers.sort(key=lambda item: item["value"])
    return share_analyses, top_performers[:top_n], bottom_performers[:top_n], warnings


def build_top_metrics(summary: dict, top_n: int = 3) -> list[dict]:
    metrics = summary.get("major_metrics", [])
    ranked = sorted(
        metrics,
        key=lambda item: (
            0 if item.get("value_scale_guess") in {"amount", "count", "percent"} else 1,
            -(abs(item.get("sum") or item.get("avg") or 0)),
        ),
    )
    return ranked[:top_n]


def build_top_categories(summary: dict, top_n: int = 3) -> list[dict]:
    categories = []
    for share in summary.get("category_shares", []):
        top_category = (share.get("top_categories_by_sum") or [None])[0]
        if not top_category:
            continue
        categories.append({
            "category_column": share.get("category_column"),
            "metric_column": share.get("metric_column"),
            "category": top_category.get("category"),
            "sum": top_category.get("sum"),
            "share": top_category.get("share"),
        })
    categories.sort(key=lambda item: item.get("share") or 0, reverse=True)
    return categories[:top_n]


def build_anomaly_summary(summary: dict) -> list[str]:
    anomaly_notes = []
    trend = summary.get("trend_analysis") or {}
    for change in trend.get("metric_changes", []):
        delta_pct = change.get("delta_pct")
        if delta_pct is not None and abs(delta_pct) >= 0.3:
            direction = "증가" if delta_pct > 0 else "감소"
            anomaly_notes.append(f"{change['metric']}이(가) 직전 대비 {abs(delta_pct):.0%} {direction}")
    for warning in summary.get("warnings", []):
        if "extreme skew" in warning or "음수 비중" in warning or "집중" in warning:
            anomaly_notes.append(warning)
    return dedupe_preserve_order(anomaly_notes)[:5]


def build_data_quality_notes(summary: dict) -> list[str]:
    thresholds = get_analysis_thresholds()
    notes = []
    key_metrics = {metric["column"] for metric in build_top_metrics(summary, top_n=3)}
    for metric in summary.get("major_metrics", []):
        if metric["column"] in key_metrics and (metric.get("null_ratio") or 0) >= thresholds["critical_missing_warning_ratio"]:
            notes.append(f"{metric['column']} 핵심 지표의 결측률이 {metric['null_ratio']:.0%}입니다.")
    notes.extend(summary.get("typo_warnings", []))
    notes.extend(summary.get("label_consistency_warnings", []))
    notes.extend(summary.get("warnings", []))
    return dedupe_preserve_order(notes)[:8]


def build_table_main_points(summary: dict) -> list[str]:
    points = []
    top_metrics = build_top_metrics(summary, top_n=3)
    top_categories = build_top_categories(summary, top_n=2)
    trend = summary.get("trend_analysis") or {}
    top_performers = summary.get("top_performers", [])
    bottom_performers = summary.get("bottom_performers", [])

    for metric in top_metrics[:2]:
        if metric.get("value_scale_guess") == "percent" and metric.get("avg") is not None:
            points.append(f"{metric['column']} 기준 평균은 {metric['avg']:.1%}")
        elif metric.get("sum") is not None:
            points.append(f"{metric['column']} 기준 합계는 {format_number(metric['sum'])}")

    if top_categories:
        top_category = top_categories[0]
        share = top_category.get("share")
        points.append(
            f"{top_category['metric_column']} 기준 {top_category['category_column']}의 최상위 범주는 {top_category['category']}이며 비중은 {'-' if share is None else f'{share:.0%}'}"
        )

    metric_changes = trend.get("metric_changes", [])
    if metric_changes:
        best_change = sorted(metric_changes, key=lambda item: abs(item.get("delta") or 0), reverse=True)[0]
        delta_pct = best_change.get("delta_pct")
        direction = "증가" if (best_change.get("delta") or 0) >= 0 else "감소"
        points.append(
            f"최근 시점 비교에서 {best_change['metric']}이(가) 가장 크게 {direction}했으며 변화율은 {'-' if delta_pct is None else f'{abs(delta_pct):.0%}'}"
        )

    if top_performers:
        top_item = top_performers[0]
        points.append(f"상위 항목은 {top_item['category_column']}={top_item['category']} ({top_item['metric_column']} {format_number(top_item['value'])})")
    if bottom_performers:
        bottom_item = bottom_performers[0]
        points.append(f"하위 항목은 {bottom_item['category_column']}={bottom_item['category']} ({bottom_item['metric_column']} {format_number(bottom_item['value'])})")

    return dedupe_preserve_order(points)[:5]


def choose_canonical_text_value(values: list[str]) -> str:
    def sort_key(value: str):
        return (-len(value.strip()), value != value.strip(), value.lower(), value)
    return sorted(values, key=sort_key)[0]


def detect_case_whitespace_or_separator_reason(a: str, b: str) -> str:
    if a.strip() != a or b.strip() != b:
        return "공백 차이"
    if a.lower() == b.lower() and a != b:
        return "대소문자 차이"
    if simplify_text_token(a) == simplify_text_token(b) and normalize_text_token(a) != normalize_text_token(b):
        return "특수문자/구분자 차이"
    return "유사 표기"


def build_text_similarity_pair(column: str, a: str, b: str, counts: dict) -> dict | None:
    thresholds = get_analysis_thresholds()
    similarity = difflib.SequenceMatcher(None, normalize_text_token(a), normalize_text_token(b)).ratio()
    simplified_match = simplify_text_token(a) == simplify_text_token(b)
    if not simplified_match and similarity < thresholds["typo_similarity_ratio"]:
        return None
    canonical = choose_canonical_text_value([a, b])
    similar_values = [value for value in [a, b] if value != canonical]
    return {
        "column": column,
        "canonical_candidate": canonical,
        "similar_values": similar_values,
        "similarity_score": round(max(similarity, 1.0 if simplified_match else similarity), 3),
        "reason": detect_case_whitespace_or_separator_reason(a, b),
        "counts": {a: counts.get(a, 0), b: counts.get(b, 0)},
    }


def detect_typo_candidates(text_columns: list[str], column_profiles: dict, rows: list[dict]) -> tuple[list[dict], list[dict], list[dict], list[dict], list[str], list[str]]:
    thresholds = get_analysis_thresholds()
    typo_candidates = []
    merge_candidates = []
    anomaly_text_candidates = []
    typo_warnings = []
    consistency_warnings = []

    for column in text_columns:
        profile = column_profiles.get(column, {})
        if profile.get("type") in {"date_like", "month_like"}:
            continue
        values = [str(row.get(column)) for row in rows if not is_missing_value(row.get(column))]
        if len(values) < 2:
            continue
        counts = {}
        for value in values:
            counts[value] = counts.get(value, 0) + 1

        unique_values = sorted(counts.keys())
        seen_pairs = set()
        for i, left in enumerate(unique_values):
            for right in unique_values[i + 1:]:
                pair_key = tuple(sorted((left, right)))
                if pair_key in seen_pairs:
                    continue
                seen_pairs.add(pair_key)

                pair = build_text_similarity_pair(column, left, right, counts)
                if pair:
                    typo_candidates.append(pair)
                    typo_warnings.append(f"{column} 컬럼에 유사 표기 후보가 있습니다: {pair['canonical_candidate']} / {', '.join(pair['similar_values'])}")
                    continue

                merge_ratio = difflib.SequenceMatcher(None, simplify_text_token(left), simplify_text_token(right)).ratio()
                if merge_ratio >= thresholds["merge_similarity_ratio"]:
                    merge_candidates.append({
                        "column": column,
                        "canonical_candidate": choose_canonical_text_value([left, right]),
                        "merge_values": [left, right],
                        "similarity_score": round(merge_ratio, 3),
                        "reason": "병합 검토 후보",
                    })
                    consistency_warnings.append(f"{column} 컬럼에 병합 후보가 있습니다: {left} / {right}")

        for value, count in counts.items():
            token = normalize_text_token(value)
            simplified = simplify_text_token(value)
            if len(token) < thresholds["min_length_text_normal"] and not (len(token) == 1 and token.isalpha()):
                anomaly_text_candidates.append({"column": column, "value": value, "reason": "비정상적으로 짧은 값", "count": count})
            elif len(set(token)) == 1 and len(token) >= 3:
                anomaly_text_candidates.append({"column": column, "value": value, "reason": "단일 문자 반복", "count": count})
            elif count <= thresholds["rare_text_count"] and len(token) > 2:
                alpha = sum(ch.isalpha() for ch in value)
                digit = sum(ch.isdigit() for ch in value)
                if alpha and digit and not re.fullmatch(r"[A-Za-z0-9_-]+", value):
                    anomaly_text_candidates.append({"column": column, "value": value, "reason": "숫자/문자 혼합 패턴 이상", "count": count})
                else:
                    anomaly_text_candidates.append({"column": column, "value": value, "reason": "희귀값", "count": count})
            elif simplified and re.fullmatch(r"[a-z]+[0-9]+[a-z]+", simplified):
                anomaly_text_candidates.append({"column": column, "value": value, "reason": "혼합 패턴 튀는 값", "count": count})

    typo_candidates = typo_candidates[:10]
    merge_candidates = merge_candidates[:10]
    anomaly_text_candidates = anomaly_text_candidates[:10]
    return (
        typo_candidates,
        merge_candidates,
        anomaly_text_candidates,
        dedupe_typo_candidates(typo_candidates),
        dedupe_preserve_order(typo_warnings)[:5],
        dedupe_preserve_order(consistency_warnings)[:5],
    )


def dedupe_typo_candidates(candidates: list[dict]) -> list[dict]:
    seen = set()
    result = []
    for candidate in candidates:
        key = (candidate["column"], candidate["canonical_candidate"], tuple(candidate["similar_values"]))
        if key in seen:
            continue
        seen.add(key)
        result.append(candidate)
    return result


def copy_rows(rows: list[dict]) -> list[dict]:
    return [dict(row) for row in rows]


def apply_merge_candidates_to_rows(
    rows: list[dict],
    typo_candidates: list[dict],
    merge_candidates: list[dict],
    enabled: bool,
) -> list[dict]:
    if not enabled:
        return copy_rows(rows)

    canonical_map = {}
    for candidate in typo_candidates:
        column = candidate.get("column")
        canonical = candidate.get("canonical_candidate")
        if not column or canonical is None:
            continue
        for value in candidate.get("similar_values", []):
            canonical_map[(column, value)] = canonical

    for candidate in merge_candidates:
        column = candidate.get("column")
        canonical = candidate.get("canonical_candidate")
        if not column or canonical is None:
            continue
        for value in candidate.get("merge_values", []):
            if value != canonical:
                canonical_map[(column, value)] = canonical

    merged_rows = []
    for row in rows:
        new_row = dict(row)
        for (column, value), canonical in canonical_map.items():
            if new_row.get(column) == value:
                new_row[column] = canonical
        merged_rows.append(new_row)
    return merged_rows


def build_analysis_candidates(table: dict) -> dict:
    summary = summarize_table(table, config=get_default_analysis_config())
    detected_types = summary.get("detected_column_types", {})
    headers = table.get("headers", [])
    label_candidates = ["자동 추론 사용"]
    for header in headers:
        col_type = detected_types.get(header)
        if col_type in {"text", "date_like", "month_like"}:
            label_candidates.append(header)
        elif col_type == "id_like":
            label_candidates.append(f"{header} (ID-like)")

    time_candidates = ["자동 추론 사용", "없음"]
    for column in summary.get("time_columns", []):
        if column not in time_candidates:
            time_candidates.append(column)

    metric_candidates = summary.get("numeric_columns", [])

    return {
        "summary": summary,
        "label_candidates": label_candidates,
        "time_candidates": time_candidates,
        "metric_candidates": metric_candidates,
        "merge_preview": summary.get("typo_candidates", []) + summary.get("merge_candidates", []),
    }


def normalize_config_value(value: str | None) -> str | None:
    if value in (None, "", "자동 추론 사용", "없음"):
        return None
    return re.sub(r"\s+\(ID-like\)$", "", value)


def build_analysis_config_summary(config: AnalysisConfig, summary: dict) -> dict:
    return {
        "selected_label_column": config.selected_label_column or summary.get("label_column"),
        "selected_time_column": normalize_config_value(config.selected_time_column) if config.selected_time_column is not None else (
            summary.get("time_columns", [None])[0] if summary.get("time_columns") else None
        ),
        "selected_metric_columns": config.selected_metric_columns or summary.get("numeric_columns", []),
        "exclude_total_rows": config.exclude_total_rows,
        "apply_merge_candidates": config.apply_merge_candidates,
        "use_first_row_as_header": config.use_first_row_as_header,
    }


def parse_time_value(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    token = normalize_text_token(value)
    if not token:
        return None
    normalized = token.replace(".", "-").replace("/", "-")
    if re.fullmatch(r"\d{6}", token):
        return datetime.strptime(token + "01", "%Y%m%d")
    if re.fullmatch(r"\d{4}-\d{1,2}", normalized):
        return datetime.strptime(normalized + "-01", "%Y-%m-%d")
    if re.fullmatch(r"\d{1,2}월", token):
        return (0, int(token[:-1]))
    if re.fullmatch(r"\d{2}년\s?\d{1,2}월", token):
        match = re.match(r"(\d{2})년\s?(\d{1,2})월", token)
        if match:
            return (2000 + int(match.group(1)), int(match.group(2)))
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%m-%d-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(token, fmt)
        except ValueError:
            continue
    month_map = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
    if token in month_map:
        return (0, month_map[token])
    return None


def build_metric_change(metric: str, latest: float | None, previous: float | None) -> dict:
    delta = None
    delta_pct = None
    if latest is not None and previous is not None:
        delta = latest - previous
        delta_pct = safe_divide(delta, previous)
    return {
        "metric": metric,
        "latest": latest,
        "previous": previous,
        "delta": delta,
        "delta_pct": delta_pct,
    }


def build_row_time_trend_analysis(time_columns: list[str], numeric_columns: list[str], rows: list[dict]) -> dict | None:
    for time_column in time_columns:
        enriched_rows = []
        for row in rows:
            sort_key = parse_time_value(row.get(time_column))
            if sort_key is not None:
                enriched_rows.append((sort_key, row))
        if len(enriched_rows) < 2:
            continue
        enriched_rows.sort(key=lambda item: item[0])
        previous_row = enriched_rows[-2][1]
        latest_row = enriched_rows[-1][1]
        metric_changes = []
        for metric in numeric_columns:
            latest = safe_float(latest_row.get(metric))
            previous = safe_float(previous_row.get(metric))
            if latest is None or previous is None:
                continue
            metric_changes.append(build_metric_change(metric, latest, previous))
        if metric_changes:
            return {
                "time_axis_type": "row_time",
                "time_column": time_column,
                "latest_period": latest_row.get(time_column),
                "previous_period": previous_row.get(time_column),
                "metric_changes": metric_changes,
            }
    return None


def detect_month_series_columns(headers: list[str]) -> list[str]:
    return [header for header in headers if is_month_like_value(header)]


def build_column_time_trend_analysis(month_series_columns: list[str], rows: list[dict], label_column: str | None) -> dict | None:
    if len(month_series_columns) < 2 or not rows:
        return None
    sortable_columns = []
    for header in month_series_columns:
        sort_key = parse_time_value(header)
        if sort_key is not None:
            sortable_columns.append((sort_key, header))
    if len(sortable_columns) < 2:
        return None
    sortable_columns.sort(key=lambda item: item[0])
    previous_column = sortable_columns[-2][1]
    latest_column = sortable_columns[-1][1]
    metric_changes = []
    for row in rows:
        latest = safe_float(row.get(latest_column))
        previous = safe_float(row.get(previous_column))
        if latest is None or previous is None:
            continue
        metric_name = row.get(label_column) if label_column and row.get(label_column) is not None else f"row_{len(metric_changes)}"
        metric_changes.append(build_metric_change(str(metric_name), latest, previous))
    if not metric_changes:
        return None
    metric_changes.sort(key=lambda item: abs(item.get("delta") or 0), reverse=True)
    return {
        "time_axis_type": "column_time",
        "latest_period": latest_column,
        "previous_period": previous_column,
        "metric_changes": metric_changes[:5],
        "month_series_columns": [header for _, header in sortable_columns],
    }


def build_trend_analysis(summary: dict, rows: list[dict], headers: list[str]) -> dict | None:
    row_time_trend = build_row_time_trend_analysis(summary.get("time_columns", []), summary.get("numeric_columns", []), rows)
    if row_time_trend:
        return row_time_trend
    return build_column_time_trend_analysis(summary.get("month_series_columns", []), rows, summary.get("label_column"))


def build_kpi_brief(summary: dict) -> dict:
    trend_analysis = summary.get("trend_analysis") or {}
    metric_changes = sorted(trend_analysis.get("metric_changes", []), key=lambda item: abs(item.get("delta") or 0), reverse=True)
    top_metric_changes = metric_changes[:3]

    biggest_category = None
    if summary.get("category_shares"):
        top_share = summary["category_shares"][0]
        top_category = (top_share.get("top_categories_by_sum") or [None])[0]
        if top_category:
            biggest_category = {
                "category_column": top_share.get("category_column"),
                "metric_column": top_share.get("metric_column"),
                "category": top_category.get("category"),
                "share": top_category.get("share"),
            }

    trend_summary = None
    if trend_analysis:
        trend_summary = {
            "time_axis_type": trend_analysis.get("time_axis_type"),
            "latest_period": trend_analysis.get("latest_period"),
            "previous_period": trend_analysis.get("previous_period"),
            "top_metric_changes": top_metric_changes,
        }

    return {
        "top_metric_changes": top_metric_changes,
        "biggest_category": biggest_category,
        "biggest_risk": (summary.get("anomaly_summary") or summary.get("data_quality_notes") or [None])[0],
        "missing_data_summary": [warning for warning in summary.get("data_quality_notes", []) if "결측률" in warning][:3],
        "concentration_summary": [warning for warning in summary.get("data_quality_notes", []) if "비중" in warning or "집중" in warning][:3],
        "trend_summary": trend_summary,
    }


def dedupe_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    result = []
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def build_insight_summary(summary: dict) -> dict:
    numeric_highlights = []
    for header in summary["numeric_columns"]:
        profile = summary["column_profiles"].get(header, {})
        if profile.get("count", 0) == 0:
            continue
        highlight = {
            "column": header,
            "avg": profile.get("avg"),
            "median": profile.get("median"),
            "min": profile.get("min"),
            "max": profile.get("max"),
            "top_values": profile.get("top_values", [])[:3],
            "bottom_values": profile.get("bottom_values", [])[:3],
        }
        numeric_highlights.append(highlight)

    missing_data_warnings = [w for w in summary["warnings"] if "결측률" in w]
    skew_warnings = [w for w in summary["warnings"] if "편중" in w]
    total_row_warnings = [w for w in summary["warnings"] if "합계행" in w or "총계행" in w]

    return {
        "table_topic_guess": summary.get("table_topic_guess", "general"),
        "table_main_points": summary.get("table_main_points", []),
        "top_metrics": summary.get("top_metrics", []),
        "top_categories": summary.get("top_categories", []),
        "trend_summary": (summary.get("kpi_brief") or {}).get("trend_summary"),
        "anomaly_summary": summary.get("anomaly_summary", []),
        "data_quality_notes": summary.get("data_quality_notes", []),
        "numeric_highlights": numeric_highlights,
        "missing_data_warnings": missing_data_warnings,
        "skew_warnings": skew_warnings,
        "total_row_warnings": total_row_warnings,
        "major_metrics": summary.get("major_metrics", []),
        "category_shares": summary.get("category_shares", []),
        "top_performers": summary.get("top_performers", []),
        "bottom_performers": summary.get("bottom_performers", []),
        "trend_analysis": summary.get("trend_analysis"),
        "kpi_brief": summary.get("kpi_brief"),
        "typo_candidates": summary.get("typo_candidates", []),
        "merge_candidates": summary.get("merge_candidates", []),
        "anomaly_text_candidates": summary.get("anomaly_text_candidates", []),
        "detected_column_types": summary.get("detected_column_types", {}),
        "time_columns": summary.get("time_columns", []),
    }


# =========================================================
# 4) 현재 Excel 선택 영역 읽기
# =========================================================

def get_current_excel_selection() -> dict:
    if win32com is None or pythoncom is None:
        raise RuntimeError("pywin32(win32com, pythoncom) 환경이 없어 Excel 선택영역을 읽을 수 없습니다.")

    pythoncom.CoInitialize()
    try:
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel = win32com.client.GetObject(Class="Excel.Application")

        wb = excel.ActiveWorkbook
        if wb is None:
            raise RuntimeError("활성 Workbook이 없습니다.")

        ws = excel.ActiveSheet
        sel = excel.Selection
        if sel is None:
            raise RuntimeError("현재 선택된 범위가 없습니다.")

        row_count = sel.Rows.Count
        col_count = sel.Columns.Count

        if row_count < 2:
            raise RuntimeError("선택 범위는 최소 2행 이상이어야 합니다. 첫 행은 헤더로 사용됩니다.")

        values = sel.Value

        if row_count == 1 and col_count == 1:
            values = [[values]]
        elif row_count == 1:
            values = [list(values)]
        elif col_count == 1:
            values = [[v] for v in values]
        else:
            values = [list(row) for row in values]

        normalized_values = []
        for raw_row in values:
            normalized_values.append([normalize_cell_value(cell_value) for cell_value in raw_row])

        default_table = selection_to_table(
            {
                "headers": [],
                "rows": [],
                "raw_values": normalized_values,
                "address": sel.Address,
            },
            config=get_default_analysis_config(),
        )

        return {
            "workbook_name": wb.Name,
            "sheet_name": ws.Name,
            "address": sel.Address,
            "headers": default_table["headers"],
            "rows": default_table["rows"],
            "raw_values": normalized_values,
        }
    finally:
        pythoncom.CoUninitialize()

def build_headers_from_raw_row(header_row: list) -> list[str]:
    headers = []
    used = {}
    for idx, value in enumerate(header_row, start=1):
        name = clean_header_name(value, idx)
        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1
        headers.append(name)
    return headers


def selection_to_table(selection_data: dict, config: AnalysisConfig | None = None) -> dict:
    config = config or get_default_analysis_config()
    raw_values = selection_data.get("raw_values") or []
    if not raw_values:
        return {
            "headers": selection_data.get("headers", []),
            "rows": selection_data.get("rows", []),
            "table_range": {"address": selection_data["address"]},
        }

    if config.use_first_row_as_header:
        headers = build_headers_from_raw_row(raw_values[0])
        data_rows_raw = raw_values[1:]
    else:
        col_count = max(len(row) for row in raw_values) if raw_values else 0
        headers = [f"COL_{idx}" for idx in range(1, col_count + 1)]
        data_rows_raw = raw_values

    rows = []
    for raw_row in data_rows_raw:
        item = {}
        has_any = False
        for idx, header in enumerate(headers):
            value = raw_row[idx] if idx < len(raw_row) else None
            item[header] = value
            if value is not None:
                has_any = True
        if has_any:
            rows.append(item)

    return {
        "headers": headers,
        "rows": rows,
        "table_range": {"address": selection_data["address"]},
    }


# =========================================================
# 5) 표 요약
# =========================================================

def summarize_table(table: dict, top_n: int = 5, config: AnalysisConfig | None = None) -> dict:
    config = config or get_default_analysis_config()
    headers = table["headers"]
    rows = table["rows"]
    total_row_indexes = detect_total_row_indexes(rows, headers)
    if config.exclude_total_rows:
        data_rows = [row for idx, row in enumerate(rows) if idx not in total_row_indexes]
        data_indexes = [idx for idx in range(len(rows)) if idx not in total_row_indexes]
    else:
        data_rows = list(rows)
        data_indexes = list(range(len(rows)))
    total_rows = [row for idx, row in enumerate(rows) if idx in total_row_indexes]
    table_topic_guess = guess_table_topic(headers)
    thresholds = get_analysis_thresholds()

    summary = {
        "row_count": len(rows),
        "data_row_count": len(data_rows),
        "column_count": len(headers),
        "headers": headers,
        "numeric_columns": [],
        "text_columns": [],
        "time_columns": [],
        "null_counts": {},
        "null_ratios": {},
        "column_profiles": {},
        "detected_column_types": {},
        "warnings": [],
        "total_rows_count": len(total_row_indexes),
        "total_row_indexes": total_row_indexes,
        "table_topic_guess": table_topic_guess,
        "label_column": None,
        "applied_time_column": None,
        "applied_metric_columns": [],
        "analysis_config": {},
        "major_metrics": [],
        "category_shares": [],
        "top_performers": [],
        "bottom_performers": [],
        "trend_analysis": None,
        "kpi_brief": None,
        "month_series_columns": detect_month_series_columns(headers),
        "total_row_metrics": [],
        "detected_totals": [],
        "table_main_points": [],
        "top_metrics": [],
        "top_categories": [],
        "anomaly_summary": [],
        "data_quality_notes": [],
        "typo_candidates": [],
        "merge_candidates": [],
        "anomaly_text_candidates": [],
        "typo_warnings": [],
        "label_consistency_warnings": [],
    }

    detected_types = {}
    analysis_rows = data_rows if data_rows else rows
    analysis_indexes = data_indexes if data_rows else list(range(len(rows)))
    label_column = None

    col_values = {h: [] for h in headers}
    for row in analysis_rows:
        for h in headers:
            col_values[h].append(row.get(h))

    for h in headers:
        detected_types[h] = infer_column_type(h, col_values[h], len(analysis_rows))
    summary["detected_column_types"] = detected_types
    summary["time_columns"] = [h for h, col_type in detected_types.items() if col_type in {"date_like", "month_like"}]
    label_column = normalize_config_value(config.selected_label_column) or choose_label_column(headers, analysis_rows, detected_types)
    summary["label_column"] = label_column

    for h in headers:
        values = col_values[h]
        null_count = sum(1 for v in values if is_missing_value(v))
        summary["null_counts"][h] = null_count
        summary["null_ratios"][h] = (null_count / len(values)) if values else 0

        col_type = detected_types[h]
        if col_type in {"numeric", "percent_like"}:
            summary["numeric_columns"].append(h)
            summary["column_profiles"][h] = {
                "type": col_type,
                **build_numeric_profile(h, values, analysis_rows, analysis_indexes, label_column, top_n),
            }
        else:
            summary["text_columns"].append(h)
            summary["column_profiles"][h] = {
                "type": col_type,
                **build_text_profile(values, top_n),
            }

        if summary["null_ratios"][h] >= thresholds["missing_warning_ratio"]:
            summary["warnings"].append(f"{h} 컬럼의 결측률이 {summary['null_ratios'][h]:.0%}로 높습니다.")

        if col_type in {"text", "id_like", "date_like", "month_like"}:
            skew_warning = analyze_categorical_skew(h, summary["column_profiles"][h], len(analysis_rows))
            if skew_warning:
                summary["warnings"].append(skew_warning)

    summary["time_columns"] = (
        [normalize_config_value(config.selected_time_column)]
        if normalize_config_value(config.selected_time_column)
        else summary["time_columns"]
    )
    summary["applied_time_column"] = summary["time_columns"][0] if summary["time_columns"] else None
    if config.selected_time_column == "없음":
        summary["time_columns"] = []
        summary["applied_time_column"] = None

    if config.selected_metric_columns:
        selected_metrics = [column for column in config.selected_metric_columns if column in summary["numeric_columns"]]
        summary["numeric_columns"] = selected_metrics
    summary["applied_metric_columns"] = list(summary["numeric_columns"])

    summary["major_metrics"] = build_major_metrics(summary["numeric_columns"], summary["column_profiles"])
    for metric in summary["major_metrics"]:
        profile = summary["column_profiles"].get(metric["column"], {})
        summary["warnings"].extend(analyze_numeric_quality(metric["column"], profile, metric["value_scale_guess"]))

    category_columns = choose_category_columns(headers, detected_types, summary["column_profiles"], label_column)
    category_shares, top_performers, bottom_performers, category_warnings = build_category_analyses(
        category_columns=category_columns,
        numeric_columns=summary["numeric_columns"],
        rows=analysis_rows,
        top_n=top_n,
    )
    summary["category_shares"] = category_shares
    summary["top_performers"] = top_performers
    summary["bottom_performers"] = bottom_performers
    summary["warnings"].extend(category_warnings)

    (
        typo_candidates,
        merge_candidates,
        anomaly_text_candidates,
        deduped_typo_candidates,
        typo_warnings,
        consistency_warnings,
    ) = detect_typo_candidates(summary["text_columns"], summary["column_profiles"], analysis_rows)
    summary["typo_candidates"] = deduped_typo_candidates
    summary["merge_candidates"] = merge_candidates
    summary["anomaly_text_candidates"] = anomaly_text_candidates
    summary["typo_warnings"] = typo_warnings
    summary["label_consistency_warnings"] = consistency_warnings

    effective_rows = apply_merge_candidates_to_rows(
        analysis_rows,
        summary["typo_candidates"],
        summary["merge_candidates"],
        config.apply_merge_candidates,
    )
    if config.apply_merge_candidates:
        category_shares, top_performers, bottom_performers, category_warnings = build_category_analyses(
            category_columns=category_columns,
            numeric_columns=summary["numeric_columns"],
            rows=effective_rows,
            top_n=top_n,
        )
        summary["category_shares"] = category_shares
        summary["top_performers"] = top_performers
        summary["bottom_performers"] = bottom_performers
        summary["warnings"].extend(category_warnings)

    summary["total_row_metrics"], summary["detected_totals"] = build_total_row_metrics(
        total_rows=total_rows,
        numeric_columns=summary["numeric_columns"],
        label_column=label_column,
    )
    summary["trend_analysis"] = build_trend_analysis(summary, effective_rows, headers)
    summary["top_metrics"] = build_top_metrics(summary, top_n=3)
    summary["top_categories"] = build_top_categories(summary, top_n=3)
    summary["anomaly_summary"] = build_anomaly_summary(summary)

    if total_row_indexes:
        summary["warnings"].append(
            f"합계행/총계행 후보 {len(total_row_indexes)}개가 감지되어 일반 분포 계산에서 제외했습니다. row_indexes={total_row_indexes}"
        )

    summary["warnings"] = dedupe_preserve_order(summary["warnings"])
    summary["data_quality_notes"] = build_data_quality_notes(summary)
    summary["table_main_points"] = build_table_main_points(summary)
    summary["analysis_config"] = build_analysis_config_summary(config, summary)
    summary["kpi_brief"] = build_kpi_brief(summary)
    return summary


def build_preview_rows(table: dict, limit: int = 8) -> list:
    return table["rows"][:limit]


def build_basic_summary_text(selection_data: dict, table: dict, summary: dict) -> str:
    lines = []
    lines.append("=" * 80)
    lines.append(f"[Workbook] {selection_data['workbook_name']}")
    lines.append(f"[Sheet] {selection_data['sheet_name']}")
    lines.append(f"[선택 범위] {selection_data['address']}")
    lines.append(f"[행 수] {summary['row_count']}")
    lines.append(f"[분석 대상 행 수] {summary.get('data_row_count', summary['row_count'])}")
    lines.append(f"[열 수] {summary['column_count']}")
    applied_config = summary.get("analysis_config", {})
    lines.append(f"[적용 label column] {applied_config.get('selected_label_column') or '-'}")
    lines.append(f"[적용 time column] {applied_config.get('selected_time_column') or '-'}")
    lines.append(f"[적용 metric columns] {', '.join(applied_config.get('selected_metric_columns', [])) or '-'}")
    lines.append(f"[합계행 제외] {'ON' if applied_config.get('exclude_total_rows', True) else 'OFF'}")
    lines.append(f"[merge 후보 반영] {'ON' if applied_config.get('apply_merge_candidates') else 'OFF'}")
    lines.append(f"[표 주제 추정] {summary.get('table_topic_guess', 'general')}")
    lines.append(f"[헤더] {', '.join(summary['headers'])}")
    if summary.get("label_column"):
        lines.append(f"[대표 라벨 컬럼] {summary['label_column']}")
    if summary.get("time_columns"):
        lines.append(f"[시간 컬럼 후보] {', '.join(summary['time_columns'])}")
    if summary.get("month_series_columns"):
        lines.append(f"[월 시계열 헤더] {', '.join(summary['month_series_columns'])}")
    if summary.get("total_rows_count"):
        lines.append(
            f"[합계행 후보] {summary['total_rows_count']}개 (indexes={summary.get('total_row_indexes', [])})"
        )
    lines.append("-" * 80)

    lines.append("[숫자형 컬럼]")
    if not summary["numeric_columns"]:
        lines.append("  - 없음")
    else:
        for col in summary["numeric_columns"]:
            p = summary["column_profiles"][col]
            lines.append(
                f"  - {col}: count={p['count']}, null={p['null_count']}, "
                f"null_ratio={p.get('null_ratio', 0):.0%}, "
                f"sum={format_number(p['sum'])}, avg={format_number(p['avg'])}, "
                f"min={format_number(p['min'])}, max={format_number(p['max'])}"
            )

    lines.append("-" * 80)
    lines.append("[문자형 컬럼]")
    if not summary["text_columns"]:
        lines.append("  - 없음")
    else:
        for col in summary["text_columns"]:
            p = summary["column_profiles"][col]
            top_vals = ", ".join(f"{x['value']}({x['count']})" for x in p["top_values"][:5])
            lines.append(
                f"  - {col}: count={p['count']}, null={p['null_count']}, "
                f"null_ratio={p.get('null_ratio', 0):.0%}, unique={p['unique_count']}, top={top_vals}"
            )

    if summary.get("table_main_points"):
        lines.append("-" * 80)
        lines.append("[현재 표에서 볼 점]")
        for point in summary.get("table_main_points", [])[:5]:
            lines.append(f"  - {point}")

    if summary.get("data_quality_notes"):
        lines.append("-" * 80)
        lines.append("[주의 메모]")
        for note in summary["data_quality_notes"][:6]:
            lines.append(f"  - {note}")

    if summary.get("trend_analysis"):
        trend = summary["trend_analysis"]
        lines.append("-" * 80)
        lines.append("[추이 분석]")
        lines.append(
            f"  - axis={trend.get('time_axis_type')}, latest={trend.get('latest_period')}, previous={trend.get('previous_period')}"
        )
        for change in trend.get("metric_changes", [])[:3]:
            delta_pct = change.get("delta_pct")
            lines.append(
                f"  - {change['metric']}: delta={format_number(change.get('delta'))}, "
                f"delta_pct={'-' if delta_pct is None else f'{delta_pct * 100:.1f}%'}"
            )

    lines.append("=" * 80)
    return "\n".join(lines)


# =========================================================
# 6) 프롬프트
# =========================================================

def build_llm_prompt(
    selection_data: dict,
    table: dict,
    summary: dict,
    insight_summary: dict,
    preview_rows: list,
    mode: str,
) -> str:
    prompt_map = {
        "summary": """
요구사항:
1. 아래 형식 유지:
   [한줄 요약]
   [핵심 포인트]
   [현재 표에서 볼 점]
   [주의 메모]
2. 핵심 포인트는 현재 표 내용 위주로 쓰고, 증감/비중/top category를 우선 활용
3. 결측/오타/표기 흔들림은 [주의 메모]에서만 짧게 언급
""",
        "report": """
요구사항:
1. 리더/임원 보고용 문체로 작성
2. 형식:
   - 현황
   - 핵심 해석
   - 시사점
   - 확인 필요 사항
3. 본문은 현재 수치와 구조 중심으로 작성
4. 확인 필요 사항에서만 데이터 품질/오타/병합 후보/결측을 언급
""",
        "risk": """
요구사항:
1. 결측만 보지 말고 급격한 증감, 과도한 편중, 이상치, 오타/표기 불일치로 인한 집계 왜곡 가능성을 함께 설명
2. 반드시 '집계 왜곡 가능성' 항목을 포함
3. 마지막에 '확인 필요 항목 3개'를 반드시 포함
"""
    }

    payload = {
        "workbook_name": selection_data["workbook_name"],
        "sheet_name": selection_data["sheet_name"],
        "table_range": table["table_range"],
        "analysis_config": summary.get("analysis_config", {}),
        "table_main_points": summary.get("table_main_points", []),
        "kpi_brief": summary.get("kpi_brief"),
        "trend_summary": (summary.get("kpi_brief") or {}).get("trend_summary"),
        "top_categories": summary.get("top_categories", []),
        "trend_analysis": summary.get("trend_analysis"),
        "top_performers": summary.get("top_performers", []),
        "bottom_performers": summary.get("bottom_performers", []),
        "typo_candidates": summary.get("typo_candidates", []),
        "merge_candidates": summary.get("merge_candidates", []),
        "anomaly_text_candidates": summary.get("anomaly_text_candidates", []),
        "data_quality_notes": summary.get("data_quality_notes", []),
        "insight_summary": insight_summary,
        "summary": summary,
        "preview_rows": preview_rows,
    }

    safe_payload = make_json_safe(payload)

    return f"""
    다음은 현재 사용자가 Excel에서 선택한 표 영역을 읽어 요약한 데이터입니다.
    이 정보를 바탕으로 한국어로 알기 쉽게 해석해 주세요.
    우선순위는 1) table_main_points 2) kpi_brief 3) trend_summary / top_categories / top_performers 4) typo_candidates / merge_candidates / anomaly_text_candidates 5) data_quality_notes 6) summary 7) preview_rows 입니다.
    사용자가 지정한 label/time/metric 기준을 자동 추론보다 우선 신뢰하세요.
    merge 후보 반영 여부와 total row 제외 여부를 analysis_config 기준으로 해석하세요.
    현재 표의 핵심 내용부터 설명하세요.
    데이터 품질 이슈보다 본 표의 주요 수치와 구조를 먼저 설명하세요.
    결측 경고만 반복하지 말고, 결측률만으로 답변을 채우지 마세요.
    오타/표기 흔들림은 집계 왜곡 가능성 관점에서만 짚으세요.
    숫자는 반드시 제공된 계산 결과 기준으로 설명하세요.
    직접 재계산하려고 하지 말고, 증감은 delta/delta_pct를 우선 사용하세요.
    비중은 share 값을 우선 사용하세요.
    합계행/총계행은 참고용 총계로만 사용하고 일반 행 통계와 혼합하지 마세요.
    모르는 것은 반드시 '추정'이라고 표시하고, 없는 사실을 상상하지 마세요.

    {prompt_map.get(mode, prompt_map["summary"])}

    분석 데이터(JSON):
    {json.dumps(safe_payload, ensure_ascii=False, indent=2)}
    """.strip()


def normalize_mail_text(text: str) -> str:
    cleaned = (text or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    cleaned = re.sub(r"[ \t]+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned


def decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    try:
        parts = []
        for chunk, encoding in decode_header(value):
            if isinstance(chunk, bytes):
                parts.append(chunk.decode(encoding or "utf-8", errors="replace"))
            else:
                parts.append(str(chunk))
        return "".join(parts).strip()
    except Exception:
        return str(value).strip()


def clean_text(text: str) -> str:
    cleaned = normalize_mail_text(html_lib.unescape(text or ""))
    cleaned = re.sub(r"[ \t]{2,}", " ", cleaned)
    return cleaned


def remove_structural_lines(text: str) -> str:
    lines = []
    structural_pattern = re.compile(r"^(from|sent|to|cc|subject|date)\s*:", re.IGNORECASE)
    separator_pattern = re.compile(r"^[-_]{3,}$")
    for line in clean_text(text).splitlines():
        stripped = line.strip()
        if structural_pattern.match(stripped):
            continue
        if separator_pattern.match(stripped):
            continue
        lines.append(line)
    return normalize_mail_text("\n".join(lines))


def split_sentences(text: str) -> list[str]:
    normalized = clean_text(text)
    return [piece.strip() for piece in re.split(r"(?<=[.!?。다])\s+|\n+", normalized) if piece.strip()]


def summarize_text(text: str, max_sentences: int = 3) -> str:
    sentences = split_sentences(text)
    if not sentences:
        return ""
    return " ".join(sentences[:max_sentences])


def html_to_text_basic(html_text: str) -> str:
    text = html_text or ""
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = html_lib.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    return normalize_mail_text(text)


def trim_mail_body(text: str, max_len: int = 12000) -> str:
    trimmed = remove_structural_lines(text)
    split_patterns = [
        r"(?m)^On .+wrote:$",
        r"(?m)^From:\s",
        r"(?m)^Sent:\s",
        r"(?m)^-----Original Message-----$",
        r"(?m)^보낸 사람\s*:",
        r"(?m)^원본 메시지\s*$",
    ]
    for pattern in split_patterns:
        match = re.search(pattern, trimmed)
        if match:
            trimmed = trimmed[:match.start()].rstrip()
            break
    if len(trimmed) > max_len:
        trimmed = trimmed[:max_len].rstrip() + "\n...[truncated]"
    return normalize_mail_text(trimmed)


def extract_body_from_message(msg) -> str:
    plain_parts = []
    html_parts = []
    try:
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                disposition = (part.get("Content-Disposition") or "").lower()
                if "attachment" in disposition:
                    continue
                try:
                    payload = part.get_payload(decode=True)
                    charset = part.get_content_charset() or "utf-8"
                    decoded = payload.decode(charset, errors="replace") if payload is not None else ""
                except Exception:
                    decoded = part.get_payload() if isinstance(part.get_payload(), str) else ""
                if content_type == "text/plain":
                    plain_parts.append(decoded)
                elif content_type == "text/html":
                    html_parts.append(decoded)
        else:
            payload = msg.get_payload(decode=True)
            charset = msg.get_content_charset() or "utf-8"
            decoded = payload.decode(charset, errors="replace") if payload is not None else ""
            if msg.get_content_type() == "text/html":
                html_parts.append(decoded)
            else:
                plain_parts.append(decoded)
    except Exception:
        return ""

    if plain_parts:
        return trim_mail_body("\n\n".join(part for part in plain_parts if part))
    if html_parts:
        return trim_mail_body(html_to_text_basic("\n".join(part for part in html_parts if part)))
    return ""


def normalize_mail_search_keyword(keyword: str) -> str:
    return re.sub(r"\s+", " ", (keyword or "").strip()).lower()


def filter_mail_items(items: list[MailItem], keyword: str) -> list[MailItem]:
    normalized_keyword = normalize_mail_search_keyword(keyword)
    if not normalized_keyword:
        return list(items)

    filtered = []
    for item in items:
        haystack = normalize_mail_search_keyword(" ".join([item.subject, item.sender, item.body]))
        if normalized_keyword in haystack:
            filtered.append(item)
    return filtered


def normalize_mail_datetime(value: datetime | None) -> datetime | None:
    if value is None:
        return None
    if value.tzinfo is None:
        return value
    try:
        return value.astimezone().replace(tzinfo=None)
    except Exception:
        return value.replace(tzinfo=None)


def pop3_connect(user: str, password: str, host: str | None = None, port: int | None = None, use_ssl: bool = True):
    if not user.strip() or not password:
        raise RuntimeError("메일함 조회를 위해 사용자 ID와 비밀번호를 입력하세요.")

    host = host or os.getenv("POP3_HOST", "pop3.samsung.net")
    port = port or int(os.getenv("POP3_PORT", "995"))
    try:
        if use_ssl:
            server = poplib.POP3_SSL(host, port, timeout=20)
        else:
            server = poplib.POP3(host, port, timeout=20)
        server.user(user)
        server.pass_(password)
        return server
    except poplib.error_proto as e:
        raise RuntimeError(f"POP3 로그인 실패: {e}") from e
    except Exception as e:
        raise RuntimeError(f"POP3 연결 실패: {e}") from e


def build_mail_item_from_message(index: int, raw_message: bytes) -> MailItem | None:
    try:
        msg = BytesParser(policy=policy.default).parsebytes(raw_message)
        subject = decode_mime_header(msg.get("Subject")) or "(제목 없음)"
        sender_name, sender_addr = parseaddr(decode_mime_header(msg.get("From")))
        sender = sender_name or sender_addr or "(발신자 없음)"
        raw_date = msg.get("Date")
        try:
            date_obj = normalize_mail_datetime(parsedate_to_datetime(raw_date)) if raw_date else None
        except Exception:
            date_obj = None
        date_str = date_obj.strftime("%Y-%m-%d %H:%M") if date_obj else (raw_date or "-")
        body = extract_body_from_message(msg) or "(본문 없음)"
        uid = msg.get("Message-ID") or f"index-{index}"
        return MailItem(
            index=index,
            uid=str(uid),
            subject=subject,
            sender=sender,
            date_str=date_str,
            date_obj=date_obj,
            body=body,
        )
    except Exception:
        return None


def fetch_recent_mails(params: MailQueryParams) -> list[MailItem]:
    server = pop3_connect(
        user=params.user_id,
        password=params.password,
        host=params.host,
        port=params.port,
        use_ssl=params.use_ssl,
    )
    scanned_items: list[MailItem] = []
    try:
        message_count = len(server.list()[1])
        if message_count <= 0:
            return []

        cutoff = None
        if params.recent_days > 0:
            cutoff = datetime.now() - timedelta(days=params.recent_days)

        scan_limit = min(message_count, max(params.max_count * 5, 60))
        start_index = message_count
        end_index = max(message_count - scan_limit + 1, 1)

        for index in range(start_index, end_index - 1, -1):
            try:
                _, lines, _ = server.retr(index)
                raw_message = b"\n".join(lines)
                item = build_mail_item_from_message(index, raw_message)
                if not item:
                    continue
                item_date = normalize_mail_datetime(item.date_obj)
                if cutoff and item_date and item_date < cutoff:
                    continue
                scanned_items.append(item)
                if not params.keyword and len(scanned_items) >= params.max_count:
                    break
            except Exception:
                continue
    finally:
        try:
            server.quit()
        except Exception:
            pass

    filtered = filter_mail_items(scanned_items, params.keyword)
    return filtered[:params.max_count]


MAIL_TYPE_KEYWORDS = {
    "승인/검토 요청": ["검토 부탁", "검토 요청", "승인 부탁", "승인 요청", "review", "approve"],
    "요청": ["부탁드립니다", "요청드립니다", "회신 부탁", "전달 부탁", "지원 부탁"],
    "일정": ["일정", "회의", "미팅", "금요일", "내일", "오늘", "까지", "오전", "오후"],
    "이슈/장애": ["이슈", "장애", "오류", "문제", "긴급", "ERR", "error"],
    "보고": ["보고드립니다", "보고 드립니다", "실적", "현황 공유", "결과 공유"],
    "공유": ["공유드립니다", "전달드립니다", "참고 부탁", "안내드립니다"],
    "문의": ["문의", "질문", "확인 가능", "확인 부탁드립니다", "알려주세요"],
}

REQUEST_LINE_PATTERN = re.compile(r"(부탁드립니다|요청드립니다|회신 부탁|검토 부탁|확인 부탁|공유 부탁|전달 부탁)")
CHECK_LINE_PATTERN = re.compile(r"(확인 필요|검토 필요|체크 필요|확인 부탁|검토 부탁|알려주세요)")
DEADLINE_PATTERN = re.compile(r"((\d{1,2}월\s?\d{1,2}일|\d{1,2}/\d{1,2}|\d{4}-\d{1,2}-\d{1,2}|오늘|내일|모레|금주|이번 주|차주|다음 주).{0,12}(까지|전까지|마감|오전|오후))")
SCHEDULE_PATTERN = re.compile(r"(\d{1,2}월\s?\d{1,2}일|\d{1,2}/\d{1,2}|\d{4}-\d{1,2}-\d{1,2}|오늘|내일|모레|금주|이번 주|차주|다음 주|오전\s?\d{1,2}시|오후\s?\d{1,2}시|회의|미팅)")
OWNER_PATTERN = re.compile(r"([가-힣A-Za-z0-9]+(?:팀|실|파트|센터|부서|책임|님))")


def split_mail_lines(text: str) -> list[str]:
    return [line.strip() for line in normalize_mail_text(text).splitlines() if line.strip()]


def infer_mail_type_hint(subject: str, body: str) -> dict:
    combined = f"{subject}\n{body}".lower()
    scores = {}
    for mail_type, keywords in MAIL_TYPE_KEYWORDS.items():
        score = sum(1 for keyword in keywords if keyword.lower() in combined)
        if score > 0:
            scores[mail_type] = score
    if not scores:
        return {"mail_type": "기타", "reason": "뚜렷한 유형 키워드가 명시되지 않음"}
    best_type = sorted(scores.items(), key=lambda item: (-item[1], item[0]))[0][0]
    return {"mail_type": best_type, "reason": f"키워드 기반 힌트: {best_type}"}


def extract_mail_rule_hints(subject: str, body: str) -> dict:
    lines = split_mail_lines(body)
    requests = []
    check_items = []
    schedule_mentions = []
    deadline_mentions = []
    owner_candidates = []

    for line in lines:
        if REQUEST_LINE_PATTERN.search(line):
            requests.append(line)
        if CHECK_LINE_PATTERN.search(line):
            check_items.append(line)
        deadline_mentions.extend(match[0] for match in DEADLINE_PATTERN.findall(line))
        schedule_mentions.extend(match for match in SCHEDULE_PATTERN.findall(line))
        owner_candidates.extend(match for match in OWNER_PATTERN.findall(line))

    hint = infer_mail_type_hint(subject, body)
    return {
        "mail_type_hint": hint["mail_type"],
        "mail_type_reason_hint": hint["reason"],
        "requests": dedupe_preserve_order(requests)[:5],
        "check_items": dedupe_preserve_order(check_items)[:5],
        "schedule_mentions": dedupe_preserve_order(schedule_mentions)[:5],
        "deadline_mentions": dedupe_preserve_order(deadline_mentions)[:5],
        "owner_candidates": dedupe_preserve_order(owner_candidates)[:5],
    }


def parse_json_object_from_text(text: str) -> dict | None:
    if not text:
        return None
    stripped = text.strip()
    candidates = [stripped]
    start = stripped.find("{")
    end = stripped.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidates.append(stripped[start:end + 1])
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            continue
    return None


def normalize_mail_structure_result(structured_info: dict | None, subject: str, body: str, hints: dict) -> dict:
    structured_info = structured_info or {}
    return {
        "mail_type": structured_info.get("mail_type") or hints.get("mail_type_hint") or "기타",
        "mail_type_reason": structured_info.get("mail_type_reason") or hints.get("mail_type_reason_hint") or "명시되지 않음",
        "summary_one_line": structured_info.get("summary_one_line") or "명시되지 않음",
        "key_points": (structured_info.get("key_points") or [])[:3],
        "requests": dedupe_preserve_order((structured_info.get("requests") or []) + hints.get("requests", []))[:5],
        "check_items": dedupe_preserve_order((structured_info.get("check_items") or []) + hints.get("check_items", []))[:5],
        "schedule_mentions": dedupe_preserve_order((structured_info.get("schedule_mentions") or []) + hints.get("schedule_mentions", []))[:5],
        "deadline_mentions": dedupe_preserve_order((structured_info.get("deadline_mentions") or []) + hints.get("deadline_mentions", []))[:5],
        "owner_candidates": dedupe_preserve_order((structured_info.get("owner_candidates") or []) + hints.get("owner_candidates", []))[:5],
        "risk_notes": (structured_info.get("risk_notes") or [])[:5],
        "subject": subject.strip() if subject else "",
        "body": body,
        "rule_hints": hints,
    }


def format_mail_structure_result(structured_info: dict) -> str:
    def format_list(title: str, items: list[str], bullet: str = "-") -> list[str]:
        if items:
            return [f"[{title}]"] + [f"{bullet} {item}" for item in items]
        return [f"[{title}]", "- 명시되지 않음"]

    lines = ["===== 메일 구조화 분석 ====="]
    lines.extend(format_list("메일 유형", [f"{structured_info.get('mail_type', '기타')} ({structured_info.get('mail_type_reason', '명시되지 않음')})"]))
    lines.extend(format_list("요청사항", structured_info.get("requests", [])))
    lines.extend(format_list("확인 필요", structured_info.get("check_items", [])))
    lines.extend(format_list("일정", structured_info.get("schedule_mentions", [])))
    lines.extend(format_list("마감", structured_info.get("deadline_mentions", [])))
    lines.extend(format_list("관련자 후보", structured_info.get("owner_candidates", [])))
    return "\n".join(lines)


def build_mail_structure_prompt(subject: str, body: str) -> str:
    normalized_body = normalize_mail_text(body)
    hints = extract_mail_rule_hints(subject, normalized_body)
    payload = {
        "subject": subject.strip() if subject else "",
        "body": normalized_body,
        "rule_hints": hints,
    }
    return f"""
다음 메일을 실무 기준으로 구조화 분석하세요.

규칙:
- 메일 본문에 없는 사실을 만들지 마세요.
- 요청/일정/기한을 우선 정리하세요.
- 애매한 것은 '명시되지 않음' 또는 '추정'으로 처리하세요.
- JSON만 출력하세요.
- 아래 스키마를 따르세요.

{{
  "mail_type": "요청|보고|공유|일정|승인/검토 요청|문의|이슈/장애|기타",
  "mail_type_reason": "...",
  "summary_one_line": "...",
  "key_points": ["...", "...", "..."],
  "requests": ["..."],
  "check_items": ["..."],
  "schedule_mentions": ["..."],
  "deadline_mentions": ["..."],
  "owner_candidates": ["..."],
  "risk_notes": ["..."]
}}

메일 데이터(JSON):
{json.dumps(payload, ensure_ascii=False, indent=2)}
    """.strip()


def build_mail_summary_prompt(subject: str, body: str, structured_info: dict | None = None) -> str:
    payload = {
        "subject": subject.strip() if subject else "",
        "body": normalize_mail_text(body),
        "structured_info": structured_info or {},
    }
    return f"""
다음 메일을 한국어로 간결하게 요약하세요.

규칙:
- 핵심만 간결하게 정리하세요.
- 요청사항이 있으면 분리해서 적으세요.
- 일정/기한 언급이 있으면 짚으세요.
- 없는 사실을 추정하지 마세요.
- 구조화 분석 결과가 있으면 그것을 우선 사용하세요.
- 아래 형식을 유지하세요.

===== 메일 핵심 요약 =====
[메일 유형]
...
[한줄 요약]
...
[핵심 포인트]
1. ...
2. ...
3. ...
[요청사항]
- ...
[확인 필요]
- ...
[일정/기한]
- ...

메일 데이터(JSON):
{json.dumps(payload, ensure_ascii=False, indent=2)}
    """.strip()


def build_mail_reply_prompt(subject: str, body: str, structured_info: dict | None = None) -> str:
    payload = {
        "subject": subject.strip() if subject else "",
        "body": normalize_mail_text(body),
        "structured_info": structured_info or {},
    }
    return f"""
다음 메일에 대한 한국어 업무 메일 답장 초안을 작성하세요.

규칙:
- 정중한 한국어 업무 메일 톤을 유지하세요.
- 지나치게 장황하지 않게 작성하세요.
- 원문에 없는 사실을 추가하지 마세요.
- 원문 근거 없는 확답은 금지합니다.
- 불확실한 내용은 '확인 후 회신드리겠습니다'처럼 중립적으로 표현하세요.
- 요청 수락/보류/확인 중 모두 가능한 중립적 표현을 허용합니다.
- mail_type, request list, check_items, deadline/schedule mentions가 있으면 반영하세요.
- 아래 형식을 유지하세요.

[메일 유형]
...
[답장 초안]
안녕하세요...

메일 데이터(JSON):
{json.dumps(payload, ensure_ascii=False, indent=2)}
    """.strip()


def build_mail_inspection_prompt(subject: str, body: str, structured_info: dict | None = None) -> str:
    payload = {
        "subject": subject.strip() if subject else "",
        "body": normalize_mail_text(body),
        "structured_info": structured_info or {},
    }
    return f"""
다음 메일 문구를 한국어 업무 메일 관점에서 표현 점검하세요.

규칙:
- 원문에 없는 사실을 추가하지 마세요.
- 어색한 표현, 중복, 과도한 표현, 불명확한 요청 문구를 우선 짚으세요.
- 수정 제안은 간결하게 작성하세요.
- 아래 형식을 유지하세요.

===== 메일 표현 점검 =====
[총평]
...
[개선 포인트]
1. ...
2. ...
3. ...
[추천 수정 방향]
- ...

메일 데이터(JSON):
{json.dumps(payload, ensure_ascii=False, indent=2)}
    """.strip()


def structure_mail_content(subject: str, body: str) -> dict:
    normalized_body = normalize_mail_text(body)
    if not normalized_body:
        raise RuntimeError("메일 본문이 비어 있습니다.")

    system_prompt = (
        "당신은 한국어 업무 메일을 실무 관점에서 구조화 분석하는 비서입니다. "
        "없는 사실은 만들지 말고, 요청/확인 필요/일정/기한/관련자를 우선 정리하세요. "
        "반드시 JSON만 출력하세요."
    )
    hints = extract_mail_rule_hints(subject, normalized_body)
    result = call_gpt_oss(
        prompt=build_mail_structure_prompt(subject, normalized_body),
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=1200,
    )
    llm_text = extract_llm_text(result)
    parsed = parse_json_object_from_text(llm_text)
    return normalize_mail_structure_result(parsed, subject, normalized_body, hints)


def summarize_mail_content(subject: str, body: str, structured_info: dict | None = None) -> str:
    normalized_body = normalize_mail_text(body)
    if not normalized_body:
        raise RuntimeError("메일 본문이 비어 있습니다.")

    system_prompt = (
        "당신은 한국어 업무 메일을 빠르게 읽고 실무자가 바로 행동할 수 있게 요약해주는 비서입니다. "
        "없는 사실은 만들지 말고, 요청사항과 일정이 있으면 분리해서 적으세요."
    )
    result = call_gpt_oss(
        prompt=build_mail_summary_prompt(subject, normalized_body, structured_info=structured_info),
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=1200,
    )
    return extract_llm_text(result)


def generate_mail_reply(subject: str, body: str, structured_info: dict | None = None) -> str:
    normalized_body = normalize_mail_text(body)
    if not normalized_body:
        raise RuntimeError("메일 본문이 비어 있습니다.")

    system_prompt = (
        "당신은 한국어 업무 메일 답장 초안을 작성하는 비서입니다. "
        "정중하고 간결하게 쓰되, 근거 없는 확답과 사실 추가는 금지합니다."
    )
    result = call_gpt_oss(
        prompt=build_mail_reply_prompt(subject, normalized_body, structured_info=structured_info),
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=1200,
    )
    return extract_llm_text(result)


def inspect_mail_expression(subject: str, body: str, structured_info: dict | None = None) -> dict:
    normalized_body = normalize_mail_text(body)
    if not normalized_body:
        raise RuntimeError("메일 본문이 비어 있습니다.")

    system_prompt = (
        "당신은 한국어 업무 메일 문장을 다듬는 에디터입니다. "
        "어색한 표현, 중복, 불명확한 요청을 짚되 과장하지 마세요."
    )
    result = call_gpt_oss(
        prompt=build_mail_inspection_prompt(subject, normalized_body, structured_info=structured_info),
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=1000,
    )
    result_text = extract_llm_text(result)
    inspection_notes = []
    for line in result_text.splitlines():
        line = line.strip()
        if not line or line.startswith("[") or line.startswith("="):
            continue
        if re.match(r"^\d+\.\s+", line):
            inspection_notes.append(re.sub(r"^\d+\.\s+", "", line))
        elif line.startswith("- "):
            inspection_notes.append(line[2:].strip())
    return {
        "result_text": result_text,
        "inspection_notes": dedupe_preserve_order(inspection_notes)[:5],
    }


def html_escape(text: str) -> str:
    if text is None:
        return ""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def nl2br(text: str) -> str:
    return html_escape(text).replace("\n", "<br>\n")


def render_info_list(items: list[str], empty_text: str = "명시되지 않음") -> str:
    if not items:
        return f"<p class='muted'>{html_escape(empty_text)}</p>"
    return "<ul>" + "".join(f"<li>{html_escape(item)}</li>" for item in items) + "</ul>"


def render_section_card(title: str, body_html: str) -> str:
    return f"""
    <section class="card">
      <h2>{html_escape(title)}</h2>
      <div class="card-body">
        {body_html}
      </div>
    </section>
    """.strip()


def render_highlight_blocks(items: list[str], empty_text: str = "핵심 내용 없음") -> str:
    if not items:
        return f"<div class='highlight-empty'>{html_escape(empty_text)}</div>"
    blocks = []
    for idx, item in enumerate(items, start=1):
        blocks.append(
            f"""
            <div class="highlight-block">
              <div class="highlight-index">{idx:02d}</div>
              <div class="highlight-text">{html_escape(item)}</div>
            </div>
            """.strip()
        )
    return "<div class='highlight-list'>" + "".join(blocks) + "</div>"


def derive_summary_one_line(structured_info: dict, summary_text: str | None, subject: str) -> str:
    summary_one_line = structured_info.get("summary_one_line")
    if summary_one_line and summary_one_line != "명시되지 않음":
        return summary_one_line
    if summary_text:
        for line in summary_text.splitlines():
            line = line.strip()
            if line and not line.startswith("[") and not line.startswith("="):
                return line
    return subject.strip() if subject else "메일 핵심 내용 요약"


def extract_reply_direction(reply_text: str | None) -> list[str]:
    if not reply_text:
        return []
    lines = []
    for line in reply_text.splitlines():
        line = line.strip()
        if not line or line.startswith("[") or line.startswith("="):
            continue
        lines.append(line)
    return lines[:3]


def build_mail_html_context(subject: str, body: str, structured_info: dict | None = None, summary_info: str | None = None, inspection_info: dict | None = None, reply_info: str | None = None) -> dict:
    normalized_body = normalize_mail_text(body)
    structured_info = structured_info or normalize_mail_structure_result(None, subject, normalized_body, extract_mail_rule_hints(subject, normalized_body))
    title = subject.strip() if subject else structured_info.get("summary_one_line") or "메일 브리핑"
    summary_one_line = derive_summary_one_line(structured_info, summary_info, subject)

    return {
        "title": title,
        "mail_type": structured_info.get("mail_type", "기타"),
        "mail_type_reason": structured_info.get("mail_type_reason", "명시되지 않음"),
        "summary_one_line": summary_one_line,
        "key_points": structured_info.get("key_points") or [],
        "requests": structured_info.get("requests") or [],
        "check_items": structured_info.get("check_items") or [],
        "schedule_mentions": structured_info.get("schedule_mentions") or [],
        "deadline_mentions": structured_info.get("deadline_mentions") or [],
        "stakeholder_candidates": structured_info.get("owner_candidates") or [],
        "risk_notes": structured_info.get("risk_notes") or [],
        "reply_suggestions": extract_reply_direction(reply_info),
        "inspection_notes": (inspection_info or {}).get("inspection_notes", []) if inspection_info else [],
        "summary_info": summary_info or "",
        "main_action": (structured_info.get("requests") or structured_info.get("check_items") or ["명시되지 않음"])[0],
        "raw_body": normalized_body,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def build_html_style_label(style: str) -> tuple[str, str]:
    mapping = {
        "briefing": ("브리핑형", "핵심과 액션을 빠르게 보는 브리핑 스타일"),
        "meeting": ("회의자료형", "회의 안건과 액션 중심의 자료형 스타일"),
        "news": ("신문형", "헤드라인과 섹션 중심의 공유문 스타일"),
    }
    return mapping.get(style, mapping["briefing"])


def build_mail_presentation_html(context: dict, style: str = "briefing") -> str:
    style_name, style_desc = build_html_style_label(style)
    requests_html = render_info_list(context.get("requests", []), empty_text="요청사항 없음")
    checks_html = render_info_list(context.get("check_items", []), empty_text="확인 필요 없음")
    schedules_html = render_info_list(context.get("schedule_mentions", []), empty_text="일정 언급 없음")
    deadlines_html = render_info_list(context.get("deadline_mentions", []), empty_text="기한 언급 없음")
    stakeholders_html = render_info_list(context.get("stakeholder_candidates", []), empty_text="관련자 명시 없음")
    risks_html = render_info_list(context.get("risk_notes", []), empty_text="리스크 메모 없음")
    reply_html = render_info_list(context.get("reply_suggestions", []), empty_text="추천 회신 방향 없음")
    inspection_html = render_info_list(context.get("inspection_notes", []), empty_text="표현 점검 메모 없음")
    key_points_cards = render_highlight_blocks(context.get("key_points", []), empty_text="핵심 포인트가 명시되지 않았습니다.")

    common_css = """
    :root {
      --bg: #eef3f8;
      --surface: #ffffff;
      --surface-soft: #f7fafc;
      --line: #d8e1eb;
      --text: #14263a;
      --muted: #5f7082;
      --accent: #164c7e;
      --accent-soft: #e7f0fa;
      --accent-2: #0b6e4f;
      --shadow: 0 18px 40px rgba(15, 23, 42, 0.08);
    }
    * { box-sizing: border-box; }
    body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; background: linear-gradient(180deg, #f5f8fb 0%, #edf2f7 100%); color: var(--text); margin: 0; }
    .wrap { max-width: 1180px; margin: 0 auto; padding: 34px 24px 52px; }
    .hero { position: relative; overflow: hidden; background: linear-gradient(135deg, #183b63 0%, #0e5d87 100%); color: #f8fbff; border-radius: 24px; padding: 30px 32px 28px; box-shadow: var(--shadow); }
    .hero::after { content: ''; position: absolute; inset: auto -60px -60px auto; width: 220px; height: 220px; background: rgba(255,255,255,0.08); border-radius: 999px; }
    .eyebrow { font-size: 12px; font-weight: 800; letter-spacing: 0.16em; text-transform: uppercase; opacity: 0.86; }
    h1 { margin: 12px 0 10px; font-size: 34px; line-height: 1.2; letter-spacing: -0.02em; }
    .subtitle { max-width: 780px; color: rgba(245, 250, 255, 0.88); font-size: 15px; line-height: 1.6; }
    .hero-meta { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 16px; }
    .tag { display: inline-flex; align-items: center; padding: 7px 12px; border-radius: 999px; background: rgba(255,255,255,0.14); color: #f5fbff; font-size: 12px; font-weight: 700; }
    .summary-card { margin-top: 18px; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.18); border-radius: 18px; padding: 18px 20px; }
    .summary-label { font-size: 12px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.82; }
    .summary-line { margin-top: 8px; font-size: 22px; line-height: 1.45; font-weight: 800; }
    .summary-sub { margin-top: 10px; color: rgba(246, 250, 255, 0.9); font-size: 14px; }
    .grid { display: grid; gap: 18px; margin-top: 22px; }
    .grid.two { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .grid.three { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .card { background: var(--surface); border: 1px solid var(--line); border-radius: 20px; padding: 20px 22px; box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05); }
    .card.accent { border-color: #bfd1e3; background: linear-gradient(180deg, #ffffff 0%, #f8fbfd 100%); }
    .card.soft { background: var(--surface-soft); }
    h2 { margin: 0 0 14px; font-size: 18px; line-height: 1.3; color: #102a43; }
    .section-kicker { display: inline-block; margin-bottom: 8px; color: var(--accent); font-size: 11px; font-weight: 800; letter-spacing: 0.14em; text-transform: uppercase; }
    .card-body p, .card-body li { font-size: 14px; line-height: 1.72; color: #22384d; }
    .card-body ul { padding-left: 20px; margin: 0; }
    .muted { color: var(--muted); }
    .highlight-list { display: grid; gap: 12px; }
    .highlight-block { display: grid; grid-template-columns: 44px 1fr; gap: 12px; align-items: start; padding: 14px; border: 1px solid #d9e5f0; border-radius: 16px; background: #f8fbff; }
    .highlight-index { width: 44px; height: 44px; display: flex; align-items: center; justify-content: center; border-radius: 12px; background: #173f6a; color: #fff; font-size: 14px; font-weight: 800; }
    .highlight-text { font-size: 15px; line-height: 1.6; font-weight: 700; color: #1b3651; }
    .highlight-empty { color: var(--muted); font-size: 14px; padding: 16px; border: 1px dashed #cad7e3; border-radius: 14px; background: #fbfdff; }
    .quick-facts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; margin-top: 20px; }
    .fact { background: rgba(255,255,255,0.14); border: 1px solid rgba(255,255,255,0.18); border-radius: 16px; padding: 14px 16px; }
    .fact-label { font-size: 11px; letter-spacing: 0.12em; text-transform: uppercase; opacity: 0.82; }
    .fact-value { margin-top: 8px; font-size: 17px; line-height: 1.45; font-weight: 700; }
    .quote { background: #fbfcfe; border: 1px solid #d8e3ed; border-left: 6px solid #376996; padding: 18px 18px; border-radius: 16px; }
    .footer { margin-top: 22px; padding-top: 16px; border-top: 1px solid #d8e1eb; color: #6c7f90; font-size: 12px; line-height: 1.7; }
    .note-strip { margin-top: 18px; padding: 14px 16px; border-radius: 14px; background: #f2f7fb; color: #31516d; font-size: 13px; line-height: 1.7; border: 1px solid #d7e4ef; }
    @media (max-width: 900px) { .grid.two, .grid.three, .quick-facts { grid-template-columns: 1fr; } .wrap { padding: 20px 14px 34px; } .hero { padding: 24px 20px; } h1 { font-size: 28px; } }
    """

    header_meta = f"""
    <div class="hero-meta">
      <div class="tag">스타일: {html_escape(style_name)}</div>
      <div class="tag">메일 유형: {html_escape(context['mail_type'])}</div>
      <div class="tag">핵심 액션: {html_escape(context.get('main_action', '명시되지 않음'))}</div>
    </div>
    """
    quick_facts = f"""
    <div class="quick-facts">
      <div class="fact"><div class="fact-label">요청사항</div><div class="fact-value">{len(context.get('requests', []))}건</div></div>
      <div class="fact"><div class="fact-label">확인 필요</div><div class="fact-value">{len(context.get('check_items', []))}건</div></div>
      <div class="fact"><div class="fact-label">일정/기한</div><div class="fact-value">{len(context.get('schedule_mentions', [])) + len(context.get('deadline_mentions', []))}건</div></div>
    </div>
    """

    if style == "meeting":
        body_html = f"""
        <div class="wrap">
          <section class="hero">
            <div class="eyebrow">Meeting Material</div>
            <h1>{html_escape(context['title'])}</h1>
            <div class="subtitle">{html_escape(style_desc)}</div>
            {header_meta}
            <div class="summary-card">
              <div class="summary-label">Meeting Summary</div>
              <div class="summary-line">{html_escape(context['summary_one_line'])}</div>
              <div class="summary-sub">{html_escape(context.get('mail_type_reason', '명시되지 않음'))}</div>
            </div>
            {quick_facts}
          </section>
          <div class="grid two">
            <section class="card accent"><div class="section-kicker">Overview</div><h2>개요</h2><div class="card-body">{key_points_cards}</div></section>
            {render_section_card("주요 액션 아이템", requests_html)}
            {render_section_card("확인 필요 사항", checks_html)}
            {render_section_card("일정 / 기한", schedules_html + deadlines_html)}
            {render_section_card("관련자 후보", stakeholders_html)}
            {render_section_card("리스크 / 표현 점검", risks_html + inspection_html)}
          </div>
          <div class="note-strip">추천 회신 방향: {html_escape(" / ".join(context.get("reply_suggestions", [])) if context.get("reply_suggestions") else "확인 후 회신드리겠습니다 수준의 보수적 회신이 적합합니다.")}</div>
          {render_section_card("원문 발췌", f"<div class='quote'>{nl2br(context['raw_body'])}</div>")}
          <div class="footer">생성 시각: {html_escape(context['generated_at'])}<br>분석 기준: 구조화 분석 결과와 요약 정보를 바탕으로 Python 템플릿에서 조립<br>주의: 원문에 없는 사실은 포함하지 않는 원칙으로 생성</div>
        </div>
        """
    elif style == "news":
        body_html = f"""
        <div class="wrap">
          <section class="hero">
            <div class="eyebrow">News Style Brief</div>
            <h1>{html_escape(context['title'])}</h1>
            <div class="subtitle">{html_escape(style_desc)}</div>
            {header_meta}
            <div class="summary-card">
              <div class="summary-label">Headline Summary</div>
              <div class="summary-line">{html_escape(context['summary_one_line'])}</div>
              <div class="summary-sub">핵심 내용을 빠르게 배포하는 공유문 형식입니다.</div>
            </div>
          </section>
          <div class="grid three">
            <section class="card accent"><div class="section-kicker">Headline</div><h2>핵심 포인트</h2><div class="card-body">{key_points_cards}</div></section>
            {render_section_card("요청 / 액션", requests_html)}
            {render_section_card("일정 / 기한", schedules_html + deadlines_html)}
          </div>
          <div class="grid two">
            {render_section_card("확인 필요 / 리스크", checks_html + risks_html + inspection_html)}
            {render_section_card("관련자 / 회신 방향", stakeholders_html + reply_html)}
          </div>
          <div class="note-strip">브리핑 메모: {html_escape(context.get('main_action', '핵심 액션 명시 없음'))}</div>
          {render_section_card("참고 원문", f"<div class='quote'>{nl2br(context['raw_body'])}</div>")}
          <div class="footer">생성 시각: {html_escape(context['generated_at'])}<br>분석 기준: 메일 구조화 결과를 중심으로 시각적 위계를 강화한 뉴스레터형 요약</div>
        </div>
        """
    else:
        body_html = f"""
        <div class="wrap">
          <section class="hero">
            <div class="eyebrow">Executive Briefing</div>
            <h1>{html_escape(context['title'])}</h1>
            <div class="subtitle">{html_escape(style_desc)}</div>
            {header_meta}
            <div class="summary-card">
              <div class="summary-label">One-line Brief</div>
              <div class="summary-line">{html_escape(context['summary_one_line'])}</div>
              <div class="summary-sub">메일 유형: {html_escape(context['mail_type'])} / 관련자: {html_escape(", ".join(context.get("stakeholder_candidates", [])) or "명시되지 않음")}</div>
            </div>
            {quick_facts}
          </section>
          <div class="grid two">
            <section class="card accent">
              <div class="section-kicker">Key Brief</div>
              <h2>핵심 포인트</h2>
              <div class="card-body">{key_points_cards}</div>
            </section>
            <section class="card soft">
              <div class="section-kicker">Primary Action</div>
              <h2>요청사항</h2>
              <div class="card-body">{requests_html}</div>
            </section>
            {render_section_card("확인 필요", checks_html)}
            {render_section_card("일정 / 기한", schedules_html + deadlines_html)}
            {render_section_card("리스크 / 표현 점검", risks_html + inspection_html)}
            {render_section_card("관련자 / 회신 방향", stakeholders_html + reply_html)}
          </div>
          {render_section_card("참고 원문", f"<div class='quote'>{nl2br(context['raw_body'])}</div>")}
          <div class="footer">생성 시각: {html_escape(context['generated_at'])}<br>분석 기준: 메일 구조화 결과, 요약, 표현 점검 메모를 바탕으로 임원 브리핑형 레이아웃 구성<br>주의: 일정/기한/요청은 원문에 명시된 내용만 반영</div>
        </div>
        """

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html_escape(context['title'])}</title>
  <style>
    {common_css}
  </style>
</head>
<body>
{body_html}
</body>
</html>"""


MAIL_HTML_STYLE_MAP = {
    "브리핑형": "briefing",
    "회의자료형": "meeting",
    "신문형": "news",
}


def run_mail_analysis(
    mode: str,
    subject: str,
    body: str,
    html_style: str = "briefing",
    structured_info: dict | None = None,
    summary_text: str | None = None,
    inspection_info: dict | None = None,
) -> dict:
    normalized_body = normalize_mail_text(body)
    if not normalized_body:
        raise RuntimeError("메일 본문이 비어 있습니다.")

    structured_info = structured_info or structure_mail_content(subject, normalized_body)
    if mode == "structure":
        return {
            "mode": mode,
            "structured_info": structured_info,
            "result_text": format_mail_structure_result(structured_info),
        }
    if mode == "summary":
        summary_text = summary_text or summarize_mail_content(subject, normalized_body, structured_info=structured_info)
        return {
            "mode": mode,
            "structured_info": structured_info,
            "result_text": summary_text,
        }
    if mode == "reply":
        reply_text = generate_mail_reply(subject, normalized_body, structured_info=structured_info)
        return {
            "mode": mode,
            "structured_info": structured_info,
            "result_text": reply_text,
        }
    if mode == "inspect":
        inspection_info = inspection_info or inspect_mail_expression(subject, normalized_body, structured_info=structured_info)
        return {
            "mode": mode,
            "structured_info": structured_info,
            "inspection_info": inspection_info,
            "result_text": inspection_info["result_text"],
        }
    if mode == "html":
        summary_text = summary_text or summarize_mail_content(subject, normalized_body, structured_info=structured_info)
        inspection_info = inspection_info or {"inspection_notes": []}
        html_context = build_mail_html_context(
            subject,
            normalized_body,
            structured_info=structured_info,
            summary_info=summary_text,
            inspection_info=inspection_info,
        )
        html_text = build_mail_presentation_html(html_context, style=html_style)
        return {
            "mode": mode,
            "structured_info": structured_info,
            "summary_text": summary_text,
            "inspection_info": inspection_info,
            "html_context": html_context,
            "html_text": html_text,
            "result_text": html_text,
        }
    raise RuntimeError(f"지원하지 않는 메일 분석 모드입니다: {mode}")


def analyze_mail_structure(subject: str, body: str) -> str:
    return run_mail_analysis("structure", subject, body)["result_text"]


def analyze_mail_summary(subject: str, body: str) -> str:
    return run_mail_analysis("summary", subject, body)["result_text"]


def analyze_mail_reply(subject: str, body: str) -> str:
    return run_mail_analysis("reply", subject, body)["result_text"]


def analyze_mail_html(subject: str, body: str, html_style: str = "briefing") -> str:
    return run_mail_analysis("html", subject, body, html_style=html_style)["result_text"]


def analyze_mail_inspect(subject: str, body: str) -> str:
    return run_mail_analysis("inspect", subject, body)["result_text"]


def analyze_selection_data(selection_data: dict, mode: str, config: AnalysisConfig | None = None) -> dict:
    config = config or get_default_analysis_config()
    table = selection_to_table(selection_data, config=config)

    if not table["rows"]:
        raise RuntimeError("선택 범위에 데이터 행이 없습니다.")

    summary = summarize_table(table, config=config)
    if not summary.get("numeric_columns"):
        raise RuntimeError("분석할 metric column이 없습니다. 숫자형 컬럼을 1개 이상 선택하세요.")

    insight_summary = build_insight_summary(summary)
    preview_rows = build_preview_rows(table, limit=8)
    basic_summary = build_basic_summary_text(selection_data, table, summary)

    system_prompt = (
        "당신은 엑셀 표를 읽고 실무자가 바로 이해할 수 있게 요약해주는 데이터 분석 도우미입니다. "
        "추정은 추정이라고 명시하고, 과장하지 말고, 데이터 품질 문제도 솔직히 짚어주세요."
    )

    user_prompt = build_llm_prompt(
        selection_data=selection_data,
        table=table,
        summary=summary,
        insight_summary=insight_summary,
        preview_rows=preview_rows,
        mode=mode,
    )

    result = call_gpt_oss(
        prompt=user_prompt,
        system_prompt=system_prompt,
        temperature=0.2,
        max_tokens=9000,
    )
    llm_text = extract_llm_text(result)

    return {
        "table": table,
        "summary": summary,
        "insight_summary": insight_summary,
        "preview_rows": preview_rows,
        "basic_summary": basic_summary,
        "llm_text": llm_text,
        "result_text": f"{basic_summary}\n\n[LLM 결과]\n{llm_text}",
    }


def run_selection_analysis(mode: str, config: AnalysisConfig | None = None, selection_data: dict | None = None) -> str:
    selection_data = selection_data or get_current_excel_selection()
    analysis_result = analyze_selection_data(selection_data=selection_data, mode=mode, config=config)
    return analysis_result["result_text"]


def get_mock_tables() -> list[dict]:
    return [
        {
            "name": "label_variation_categories",
            "headers": ["지역", "제품명", "매출", "수량"],
            "rows": [
                {"지역": "서울", "제품명": "Mobile", "매출": 2100, "수량": 80},
                {"지역": "서울 ", "제품명": "MOBILE", "매출": 1900, "수량": 74},
                {"지역": "서울시", "제품명": "Mobile", "매출": 450, "수량": 18},
                {"지역": "부산", "제품명": "LPDDR5X", "매출": 1600, "수량": 55},
                {"지역": "부산", "제품명": "LPDDR-5X", "매출": 1580, "수량": 53},
                {"지역": "총계", "제품명": None, "매출": 7630, "수량": 280},
            ],
            "table_range": {"address": "A1:D7"},
        },
        {
            "name": "current_metrics_focus",
            "headers": ["월", "제품군", "매출", "출하량", "점유율"],
            "rows": [
                {"월": "2026-01", "제품군": "Flagship", "매출": 3200, "출하량": 120, "점유율": 0.42},
                {"월": "2026-02", "제품군": "Flagship", "매출": 3550, "출하량": 135, "점유율": 0.45},
                {"월": "2026-03", "제품군": "Flagship", "매출": 4100, "출하량": 149, "점유율": 0.51},
                {"월": "총계", "제품군": None, "매출": 10850, "출하량": 404, "점유율": 0.46},
            ],
            "table_range": {"address": "A1:E5"},
        },
        {
            "name": "rare_text_anomalies",
            "headers": ["라인", "상태", "코멘트", "불량수"],
            "rows": [
                {"라인": "A", "상태": "정상", "코멘트": "OK", "불량수": 2},
                {"라인": "A", "상태": "정상", "코멘트": "정상완료", "불량수": 1},
                {"라인": "B", "상태": "정상", "코멘트": "aaaa", "불량수": 0},
                {"라인": "B", "상태": "점검", "코멘트": "ERR#12X", "불량수": 7},
                {"라인": "C", "상태": "점검", "코멘트": "?", "불량수": 6},
                {"라인": "총계", "상태": None, "코멘트": None, "불량수": 16},
            ],
            "table_range": {"address": "A1:D7"},
        },
        {
            "name": "monthly_trend_columns",
            "headers": ["제품명", "202601", "202602", "202603"],
            "rows": [
                {"제품명": "A", "202601": 120, "202602": 135, "202603": 180},
                {"제품명": "B", "202601": 90, "202602": 88, "202603": 76},
                {"제품명": "C", "202601": 45, "202602": 50, "202603": 70},
            ],
            "table_range": {"address": "A1:D4"},
        },
    ]


def run_mock_analysis_tests():
    print("[Mock analysis tests]")
    for table in get_mock_tables():
        summary = summarize_table(table)
        insight_summary = build_insight_summary(summary)
        print(f"\n=== {table['name']} ===")
        print("[table_main_points]")
        print(json.dumps(make_json_safe(summary.get("table_main_points")), ensure_ascii=False, indent=2))
        print("[typo_candidates]")
        print(json.dumps(make_json_safe(summary.get("typo_candidates")), ensure_ascii=False, indent=2))
        print("[merge_candidates]")
        print(json.dumps(make_json_safe(summary.get("merge_candidates")), ensure_ascii=False, indent=2))
        print("[anomaly_text_candidates]")
        print(json.dumps(make_json_safe(summary.get("anomaly_text_candidates")), ensure_ascii=False, indent=2))
        print("[data_quality_notes]")
        print(json.dumps(make_json_safe(summary.get("data_quality_notes")), ensure_ascii=False, indent=2))
        summary_mode_payload = {
            "table_main_points": summary.get("table_main_points"),
            "kpi_brief": summary.get("kpi_brief"),
            "top_categories": summary.get("top_categories"),
            "top_performers": summary.get("top_performers"),
            "typo_candidates": summary.get("typo_candidates"),
            "merge_candidates": summary.get("merge_candidates"),
            "anomaly_text_candidates": summary.get("anomaly_text_candidates"),
            "data_quality_notes": summary.get("data_quality_notes"),
        }
        print("[summary_mode_input]")
        print(json.dumps(make_json_safe(summary_mode_payload), ensure_ascii=False, indent=2))
        print("[insight_summary]")
        print(json.dumps(make_json_safe(insight_summary), ensure_ascii=False, indent=2))


def build_mock_selection_data_from_table(table: dict, use_first_row_as_header: bool = True) -> dict:
    headers = table["headers"]
    raw_values = []
    if use_first_row_as_header:
        raw_values.append(headers)
        for row in table["rows"]:
            raw_values.append([row.get(header) for header in headers])
    else:
        for row in table["rows"]:
            raw_values.append([row.get(header) for header in headers])

    selection_data = {
        "workbook_name": "mock.xlsx",
        "sheet_name": table.get("name", "MockSheet"),
        "address": table.get("table_range", {}).get("address", "A1"),
        "raw_values": raw_values,
        "headers": headers if use_first_row_as_header else [f"COL_{idx}" for idx in range(1, len(headers) + 1)],
        "rows": table["rows"],
    }
    return selection_data


def run_mock_config_tests():
    print("[Mock config tests]")

    table = get_mock_tables()[0]
    selection_data = build_mock_selection_data_from_table(table)
    configs = [
        ("label_by_region", AnalysisConfig(selected_label_column="지역", selected_metric_columns=["매출", "수량"])),
        ("label_by_product", AnalysisConfig(selected_label_column="제품명", selected_metric_columns=["매출", "수량"])),
        ("total_exclude_on", AnalysisConfig(exclude_total_rows=True, selected_metric_columns=["매출", "수량"])),
        ("total_exclude_off", AnalysisConfig(exclude_total_rows=False, selected_metric_columns=["매출", "수량"])),
        ("merge_off", AnalysisConfig(apply_merge_candidates=False, selected_metric_columns=["매출", "수량"])),
        ("merge_on", AnalysisConfig(apply_merge_candidates=True, selected_metric_columns=["매출", "수량"])),
    ]

    header_on_selection = build_mock_selection_data_from_table(get_mock_tables()[2], use_first_row_as_header=True)
    header_off_selection = build_mock_selection_data_from_table(get_mock_tables()[2], use_first_row_as_header=False)
    configs.extend([
        ("header_on", AnalysisConfig(use_first_row_as_header=True, selected_metric_columns=["불량수"])),
        ("header_off", AnalysisConfig(use_first_row_as_header=False)),
    ])

    for name, config in configs:
        current_selection = header_off_selection if name == "header_off" else header_on_selection if name == "header_on" else selection_data
        table_result = selection_to_table(current_selection, config=config)
        summary = summarize_table(table_result, config=config)
        print(f"\n=== {name} ===")
        print("[analysis_config]")
        print(json.dumps(make_json_safe(summary.get("analysis_config")), ensure_ascii=False, indent=2))
        print("[selected]")
        selected_info = {
            "label": summary.get("analysis_config", {}).get("selected_label_column"),
            "time": summary.get("analysis_config", {}).get("selected_time_column"),
            "metrics": summary.get("analysis_config", {}).get("selected_metric_columns"),
        }
        print(json.dumps(make_json_safe(selected_info), ensure_ascii=False, indent=2))
        print("[category_shares]")
        print(json.dumps(make_json_safe(summary.get("category_shares")), ensure_ascii=False, indent=2))
        print("[trend_analysis]")
        print(json.dumps(make_json_safe(summary.get("trend_analysis")), ensure_ascii=False, indent=2))
        print("[table_main_points]")
        print(json.dumps(make_json_safe(summary.get("table_main_points")), ensure_ascii=False, indent=2))


MAIL_SAMPLE_CASES = [
    {
        "name": "deadline_request",
        "subject": "3월 운영 계획 검토 요청",
        "body": """
안녕하세요.

첨부드린 3월 운영 계획안 검토 부탁드립니다.
가능하시면 3월 15일(금) 오전까지 의견 회신 부탁드리며,
특히 인력 배치와 예산 항목 중심으로 확인해주시면 됩니다.

추가 확인이 필요한 부분이 있으면 알려주세요.
감사합니다.
        """,
    },
    {
        "name": "meeting_followup",
        "subject": "",
        "body": """
금일 회의 내용 정리드립니다.

1. 샘플 출하는 다음 주 초 목표로 진행
2. 고객 요청사항은 내부 검토 후 회신 예정
3. 일정 변경 가능성이 있어 금요일에 다시 공유 필요

확인 부탁드립니다.
        """,
    },
    {
        "name": "issue_review_request",
        "subject": "고객 장애 이슈 검토 요청",
        "body": """
안녕하세요.

어제 저녁부터 고객사에서 로그인 오류가 반복 발생하고 있습니다.
플랫폼개발팀 확인 부탁드리며, 원인 파악 후 오늘 오후 3시 전까지 1차 공유 가능 여부 회신 부탁드립니다.
필요 시 홍길동 책임도 함께 검토 부탁드립니다.

감사합니다.
        """,
    },
]


MAIL_FETCH_MOCK_ITEMS = [
    MailItem(
        index=101,
        uid="mock-101",
        subject="3월 운영 계획 검토 요청",
        sender="운영기획팀",
        date_str="2026-03-10 09:00",
        date_obj=datetime(2026, 3, 10, 9, 0),
        body="첨부드린 3월 운영 계획안 검토 부탁드립니다. 3월 15일까지 의견 회신 부탁드립니다.",
    ),
    MailItem(
        index=102,
        uid="mock-102",
        subject="샘플 출하 일정 공유",
        sender="SCM팀",
        date_str="2026-03-11 14:30",
        date_obj=datetime(2026, 3, 11, 14, 30),
        body="다음 주 초 샘플 출하 예정입니다. 일정 변경 가능성은 금요일에 다시 공유드리겠습니다.",
    ),
    MailItem(
        index=103,
        uid="mock-103",
        subject="고객 장애 이슈 검토 요청",
        sender="플랫폼개발팀",
        date_str="2026-03-12 08:15",
        date_obj=datetime(2026, 3, 12, 8, 15),
        body="고객사 로그인 오류 관련 원인 파악 후 오늘 오후 3시 전까지 1차 공유 부탁드립니다.",
    ),
]


def test_mail_prompts():
    print("[Mail prompt tests]")
    for case in MAIL_SAMPLE_CASES:
        print(f"\n=== {case['name']} / summary ===")
        print(build_mail_summary_prompt(case["subject"], case["body"]))
        print(f"\n=== {case['name']} / reply ===")
        print(build_mail_reply_prompt(case["subject"], case["body"]))


def test_mail_structure_prompts():
    print("[Mail structure prompt tests]")
    for case in MAIL_SAMPLE_CASES:
        print(f"\n=== {case['name']} / structure ===")
        print(build_mail_structure_prompt(case["subject"], case["body"]))


def test_mail_summary_prompts():
    print("[Mail summary prompt tests]")
    for case in MAIL_SAMPLE_CASES:
        structured_info = normalize_mail_structure_result(
            None,
            case["subject"],
            normalize_mail_text(case["body"]),
            extract_mail_rule_hints(case["subject"], case["body"]),
        )
        print(f"\n=== {case['name']} / summary ===")
        print(build_mail_summary_prompt(case["subject"], case["body"], structured_info=structured_info))


def test_mail_html_context():
    print("[Mail HTML context tests]")
    for case in MAIL_SAMPLE_CASES:
        body = normalize_mail_text(case["body"])
        structured_info = normalize_mail_structure_result(
            None,
            case["subject"],
            body,
            extract_mail_rule_hints(case["subject"], body),
        )
        context = build_mail_html_context(case["subject"], body, structured_info=structured_info)
        print(f"\n=== {case['name']} / context ===")
        print(json.dumps(context, ensure_ascii=False, indent=2))


def test_mail_html_render():
    print("[Mail HTML render tests]")
    for case in MAIL_SAMPLE_CASES:
        body = normalize_mail_text(case["body"])
        structured_info = normalize_mail_structure_result(
            None,
            case["subject"],
            body,
            extract_mail_rule_hints(case["subject"], body),
        )
        context = build_mail_html_context(case["subject"], body, structured_info=structured_info)
        for style in ("briefing", "meeting", "news"):
            print(f"\n=== {case['name']} / {style} ===")
            print(build_mail_presentation_html(context, style=style))


def test_mail_html_render_styles():
    print("[Mail HTML style comparison tests]")
    case = MAIL_SAMPLE_CASES[0]
    body = normalize_mail_text(case["body"])
    structured_info = normalize_mail_structure_result(
        None,
        case["subject"],
        body,
        extract_mail_rule_hints(case["subject"], body),
    )
    context = build_mail_html_context(case["subject"], body, structured_info=structured_info)
    for style in ("briefing", "meeting", "news"):
        html = build_mail_presentation_html(context, style=style)
        print(f"\n=== style={style} / length={len(html)} ===")
        print("\n".join(html.splitlines()[:30]))


def test_result_type_state():
    print("[Mail result type state tests]")
    for result_type in ("structure", "summary", "reply", "inspect", "html", None):
        html_enabled = result_type == "html"
        print({
            "result_type": result_type,
            "label": {
                "structure": "구조화 분석",
                "summary": "핵심 요약",
                "reply": "자동 답장문구",
                "inspect": "표현 점검",
                "html": "발표용 HTML",
                None: "없음",
            }[result_type],
            "html_actions_enabled": html_enabled,
        })


def test_mail_filtering():
    print("[Mail filtering tests]")
    for keyword in ("검토", "플랫폼", "출하", "없음"):
        matched = filter_mail_items(MAIL_FETCH_MOCK_ITEMS, keyword)
        print(f"\n=== keyword={keyword} ===")
        for item in matched:
            print(f"- {item.date_str} | {item.sender} | {item.subject}")
        if not matched:
            print("- no match")


def test_trim_mail_body():
    print("[Mail trim body test]")
    source = """
안녕하세요.

본문 내용입니다.
확인 부탁드립니다.

-----Original Message-----
From: test@example.com
Sent: Monday, March 10, 2026 9:00 AM
Subject: 이전 메일

이전 메일 내용
    """
    print(trim_mail_body(source))


def test_html_to_text_basic():
    print("[Mail html_to_text_basic test]")
    source = "<html><body><p>안녕하세요.</p><p><b>검토 부탁드립니다.</b><br>3월 15일까지 회신 부탁드립니다.</p></body></html>"
    print(html_to_text_basic(source))


# =========================================================
# 7) Tkinter UI
# =========================================================

class ExcelLLMApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel / 메일 LLM 분석기")
        self.root.geometry("1480x940")

        self.excel_status_var = tk.StringVar(value="준비됨")
        self.mail_status_var = tk.StringVar(value="준비됨")
        self.mode_var = tk.StringVar(value="summary")
        self.selection_info_var = tk.StringVar(value="선택영역을 아직 불러오지 않았습니다.")
        self.headers_var = tk.StringVar(value="-")
        self.auto_label_var = tk.StringVar(value="-")
        self.auto_time_var = tk.StringVar(value="-")
        self.auto_metrics_var = tk.StringVar(value="-")
        self.auto_total_rows_var = tk.StringVar(value="-")
        self.mail_fetch_user_var = tk.StringVar(value="")
        self.mail_fetch_password_var = tk.StringVar(value="")
        self.mail_search_var = tk.StringVar(value="")
        self.mail_recent_days_var = tk.StringVar(value="7")
        self.mail_max_count_var = tk.StringVar(value="20")
        self.selected_mail_info_var = tk.StringVar(value="선택 메일: 없음")
        self.mail_result_type_var = tk.StringVar(value="현재 결과: 없음")
        self.mail_result_meta_var = tk.StringVar(value="메일 메타: 없음")

        self.selection_data: dict | None = None
        self.auto_summary: dict | None = None
        self.metric_vars: dict[str, tk.BooleanVar] = {}
        self.metric_checkbuttons: list[ttk.Checkbutton] = []

        self.label_var = tk.StringVar(value="자동 추론 사용")
        self.time_var = tk.StringVar(value="자동 추론 사용")
        self.exclude_total_rows_var = tk.BooleanVar(value=True)
        self.apply_merge_candidates_var = tk.BooleanVar(value=False)
        self.use_first_row_header_var = tk.BooleanVar(value=True)
        self.mail_structured_info: dict | None = None
        self.mail_inspection_info: dict | None = None
        self.mail_last_html: str = ""
        self.mail_html_style_var = tk.StringVar(value="브리핑형")
        self.mail_items: list[MailItem] = []
        self.mail_selected_item: MailItem | None = None
        self.current_result_type: str | None = None
        self.current_mail_meta: dict[str, str] = {}

        self._build_ui()

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)
        self.main_notebook = notebook

        self.excel_tab = ttk.Frame(notebook)
        self.mail_tab = ttk.Frame(notebook)
        self._build_mail_tab(self.mail_tab)
        self._build_excel_tab(self.excel_tab)
        notebook.add(self.mail_tab, text="메일 분석")
        notebook.add(self.excel_tab, text="Excel 분석")
        notebook.select(self.mail_tab)

    def _build_excel_tab(self, parent: ttk.Frame):
        top = ttk.Frame(parent, padding=12)
        top.pack(fill="x")

        title = ttk.Label(top, text="현재 Excel 선택영역 요약 / 해석", font=("맑은 고딕", 15, "bold"))
        title.pack(anchor="w")

        desc = ttk.Label(
            top,
            text="Excel에서 범위를 선택한 뒤 버튼을 누르세요. 첫 행은 헤더, 아래 행은 데이터로 인식합니다.",
            font=("맑은 고딕", 10),
        )
        desc.pack(anchor="w", pady=(4, 10))

        mode_frame = ttk.LabelFrame(top, text="분석 유형", padding=10)
        mode_frame.pack(fill="x", pady=(0, 10))

        ttk.Radiobutton(mode_frame, text="일반 요약", variable=self.mode_var, value="summary").pack(side="left", padx=8)
        ttk.Radiobutton(mode_frame, text="보고용 문구", variable=self.mode_var, value="report").pack(side="left", padx=8)
        ttk.Radiobutton(mode_frame, text="주의점 분석", variable=self.mode_var, value="risk").pack(side="left", padx=8)

        btn_frame = ttk.Frame(top)
        btn_frame.pack(fill="x", pady=(0, 10))

        self.btn_load = ttk.Button(btn_frame, text="선택영역 불러오기", command=self.on_load_selection)
        self.btn_load.pack(side="left", padx=(0, 8))

        self.btn_run = ttk.Button(btn_frame, text="현재 설정으로 분석 실행", command=self.on_run)
        self.btn_run.pack(side="left", padx=(0, 8))

        self.btn_copy = ttk.Button(btn_frame, text="결과 복사", command=self.on_copy)
        self.btn_copy.pack(side="left", padx=(0, 8))

        self.btn_clear = ttk.Button(btn_frame, text="결과 지우기", command=self.on_clear)
        self.btn_clear.pack(side="left", padx=(0, 8))

        self.btn_help = ttk.Button(btn_frame, text="사용 방법", command=self.on_help)
        self.btn_help.pack(side="left")

        status_frame = ttk.Frame(top)
        status_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(status_frame, text="상태:", font=("맑은 고딕", 10, "bold")).pack(side="left")
        ttk.Label(status_frame, textvariable=self.excel_status_var).pack(side="left", padx=(6, 0))

        self._build_config_panel(parent)
        self._build_preview_panel(parent)
        self._build_result_panel(parent)

    def _build_config_panel(self, parent: ttk.Frame):
        config_frame = ttk.LabelFrame(parent, text="분석 설정", padding=12)
        config_frame.pack(fill="x", padx=12, pady=(0, 10))

        info_frame = ttk.Frame(config_frame)
        info_frame.pack(fill="x")

        ttk.Label(info_frame, textvariable=self.selection_info_var).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        ttk.Label(info_frame, text="Headers:", font=("맑은 고딕", 9, "bold")).grid(row=1, column=0, sticky="nw")
        ttk.Label(info_frame, textvariable=self.headers_var, wraplength=980).grid(row=1, column=1, columnspan=3, sticky="w")
        ttk.Label(info_frame, text="자동 label:", font=("맑은 고딕", 9, "bold")).grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, textvariable=self.auto_label_var).grid(row=2, column=1, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, text="자동 time:", font=("맑은 고딕", 9, "bold")).grid(row=2, column=2, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, textvariable=self.auto_time_var).grid(row=2, column=3, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, text="자동 metrics:", font=("맑은 고딕", 9, "bold")).grid(row=3, column=0, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, textvariable=self.auto_metrics_var, wraplength=400).grid(row=3, column=1, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, text="감지 total rows:", font=("맑은 고딕", 9, "bold")).grid(row=3, column=2, sticky="w", pady=(4, 0))
        ttk.Label(info_frame, textvariable=self.auto_total_rows_var).grid(row=3, column=3, sticky="w", pady=(4, 0))

        control_frame = ttk.Frame(config_frame)
        control_frame.pack(fill="x", pady=(10, 0))

        ttk.Label(control_frame, text="Label column").grid(row=0, column=0, sticky="w")
        self.label_combo = ttk.Combobox(control_frame, textvariable=self.label_var, state="readonly", width=28)
        self.label_combo.grid(row=0, column=1, sticky="w", padx=(6, 18))

        ttk.Label(control_frame, text="Time column").grid(row=0, column=2, sticky="w")
        self.time_combo = ttk.Combobox(control_frame, textvariable=self.time_var, state="readonly", width=24)
        self.time_combo.grid(row=0, column=3, sticky="w", padx=(6, 0))

        option_frame = ttk.Frame(config_frame)
        option_frame.pack(fill="x", pady=(10, 0))
        ttk.Checkbutton(
            option_frame,
            text="합계/총계 행 제외 후 분석",
            variable=self.exclude_total_rows_var,
            command=self.on_config_option_changed,
        ).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(option_frame, text="자동 병합 후보 반영", variable=self.apply_merge_candidates_var).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(
            option_frame,
            text="첫 행을 헤더로 사용",
            variable=self.use_first_row_header_var,
            command=self.on_config_option_changed,
        ).pack(side="left")

        metric_frame = ttk.LabelFrame(config_frame, text="Metric columns", padding=8)
        metric_frame.pack(fill="x", pady=(10, 0))
        self.metric_checks_frame = ttk.Frame(metric_frame)
        self.metric_checks_frame.pack(fill="x")

        merge_frame = ttk.LabelFrame(config_frame, text="Merge / Typo 후보", padding=8)
        merge_frame.pack(fill="x", pady=(10, 0))
        self.merge_text = tk.Text(merge_frame, height=5, wrap="word", font=("Consolas", 9))
        self.merge_text.pack(fill="x")
        self.merge_text.configure(state="disabled")

    def _build_preview_panel(self, parent: ttk.Frame):
        preview_frame = ttk.LabelFrame(parent, text="선택영역 미리보기", padding=12)
        preview_frame.pack(fill="both", expand=False, padx=12, pady=(0, 10))

        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=8)
        self.preview_tree.pack(side="left", fill="both", expand=True)
        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        preview_scroll.pack(side="right", fill="y")
        self.preview_tree.configure(yscrollcommand=preview_scroll.set)

    def _build_result_panel(self, parent: ttk.Frame):
        result_frame = ttk.Frame(parent, padding=(12, 0, 12, 12))
        result_frame.pack(fill="both", expand=True)

        self.txt = tk.Text(result_frame, wrap="word", font=("Consolas", 10))
        self.txt.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.txt.yview)
        scroll.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=scroll.set)

    def _build_mail_tab(self, parent: ttk.Frame):
        container = ttk.Frame(parent, padding=12)
        container.pack(fill="both", expand=True)

        title = ttk.Label(container, text="메일 분석", font=("맑은 고딕", 16, "bold"))
        title.pack(anchor="w")
        desc = ttk.Label(
            container,
            text="메일함 조회 또는 수동 입력 후 구조화 분석, 요약, 답장, 표현 점검, 발표용 HTML 생성을 실행하세요.",
            font=("맑은 고딕", 10),
        )
        desc.pack(anchor="w", pady=(4, 10))

        paned = ttk.PanedWindow(container, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left_panel = ttk.Frame(paned, padding=(0, 0, 10, 0))
        right_panel = ttk.Frame(paned, padding=(10, 0, 0, 0))
        paned.add(left_panel, weight=1)
        paned.add(right_panel, weight=2)

        self._build_mail_left_panel(left_panel)
        self._build_mail_right_panel(right_panel)

    def _build_mail_left_panel(self, parent: ttk.Frame):
        self._build_mail_fetch_section(parent)
        self._build_mail_input_section(parent)
        self._build_mail_action_section(parent)

    def _build_mail_right_panel(self, parent: ttk.Frame):
        self._build_mail_result_section(parent)

    def _build_mail_fetch_section(self, parent: ttk.Frame):
        query_frame = ttk.LabelFrame(parent, text="메일함 조회", padding=10)
        query_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(query_frame, text="사용자 ID", width=10).grid(row=0, column=0, sticky="w")
        self.mail_fetch_user_entry = ttk.Entry(query_frame, textvariable=self.mail_fetch_user_var, width=22)
        self.mail_fetch_user_entry.grid(row=0, column=1, sticky="we", padx=(0, 10))
        ttk.Label(query_frame, text="비밀번호", width=10).grid(row=0, column=2, sticky="w")
        self.mail_fetch_password_entry = ttk.Entry(query_frame, textvariable=self.mail_fetch_password_var, width=18, show="*")
        self.mail_fetch_password_entry.grid(row=0, column=3, sticky="we")

        ttk.Label(query_frame, text="검색어", width=10).grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.mail_search_entry = ttk.Entry(query_frame, textvariable=self.mail_search_var)
        self.mail_search_entry.grid(row=1, column=1, sticky="we", padx=(0, 10), pady=(8, 0))
        ttk.Label(query_frame, text="최근 N일", width=10).grid(row=1, column=2, sticky="w", pady=(8, 0))
        self.mail_recent_days_entry = ttk.Entry(query_frame, textvariable=self.mail_recent_days_var, width=8)
        self.mail_recent_days_entry.grid(row=1, column=3, sticky="w", pady=(8, 0))

        ttk.Label(query_frame, text="최대 개수", width=10).grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.mail_max_count_entry = ttk.Entry(query_frame, textvariable=self.mail_max_count_var, width=8)
        self.mail_max_count_entry.grid(row=2, column=1, sticky="w", pady=(8, 0))

        fetch_btn_frame = ttk.Frame(query_frame)
        fetch_btn_frame.grid(row=2, column=2, columnspan=2, sticky="e", pady=(8, 0))
        self.mail_btn_fetch_list = ttk.Button(fetch_btn_frame, text="메일 조회", command=self.on_mail_fetch_list)
        self.mail_btn_fetch_list.pack(side="left", padx=(0, 6))
        self.mail_btn_load_selected = ttk.Button(fetch_btn_frame, text="선택 메일 불러오기", command=self.on_mail_load_selected)
        self.mail_btn_load_selected.pack(side="left", padx=(0, 6))
        self.mail_btn_clear_login = ttk.Button(fetch_btn_frame, text="로그인 정보 지우기", command=self.on_mail_clear_login)
        self.mail_btn_clear_login.pack(side="left")

        query_frame.columnconfigure(1, weight=1)
        query_frame.columnconfigure(3, weight=1)

        list_frame = ttk.LabelFrame(parent, text="조회된 메일 목록", padding=8)
        list_frame.pack(fill="x", pady=(0, 10))
        columns = ("date", "sender", "subject")
        self.mail_list_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=6)
        self.mail_list_tree.heading("date", text="날짜")
        self.mail_list_tree.heading("sender", text="발신자")
        self.mail_list_tree.heading("subject", text="제목")
        self.mail_list_tree.column("date", width=118, anchor="w")
        self.mail_list_tree.column("sender", width=150, anchor="w")
        self.mail_list_tree.column("subject", width=300, anchor="w")
        self.mail_list_tree.pack(side="left", fill="x", expand=True)
        mail_list_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.mail_list_tree.yview)
        mail_list_scroll.pack(side="right", fill="y")
        self.mail_list_tree.configure(yscrollcommand=mail_list_scroll.set)
        self.mail_list_tree.bind("<<TreeviewSelect>>", self.on_mail_result_select)
        self.mail_list_tree.bind("<Double-1>", self.on_mail_load_selected)

        ttk.Label(parent, textvariable=self.selected_mail_info_var, foreground="#35506b", wraplength=460).pack(anchor="w", pady=(0, 8))

    def _build_mail_input_section(self, parent: ttk.Frame):
        input_frame = ttk.LabelFrame(parent, text="메일 입력", padding=10)
        input_frame.pack(fill="both", expand=True, pady=(0, 10))

        subject_frame = ttk.Frame(input_frame)
        subject_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(subject_frame, text="메일 제목", width=10).pack(side="left")
        self.mail_subject_entry = ttk.Entry(subject_frame)
        self.mail_subject_entry.pack(side="left", fill="x", expand=True)

        option_frame = ttk.Frame(input_frame)
        option_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(option_frame, text="HTML 스타일", width=10).pack(side="left")
        self.mail_html_style_combo = ttk.Combobox(
            option_frame,
            textvariable=self.mail_html_style_var,
            state="readonly",
            width=14,
            values=list(MAIL_HTML_STYLE_MAP.keys()),
        )
        self.mail_html_style_combo.pack(side="left")

        body_frame = ttk.LabelFrame(input_frame, text="메일 본문", padding=6)
        body_frame.pack(fill="both", expand=True)
        self.mail_body_text = scrolledtext.ScrolledText(
            body_frame,
            wrap="word",
            font=("Consolas", 10),
            height=18,
            spacing1=2,
            spacing3=4,
        )
        self.mail_body_text.pack(fill="both", expand=True)

    def _build_mail_action_section(self, parent: ttk.Frame):
        action_frame = ttk.LabelFrame(parent, text="분석 실행", padding=10)
        action_frame.pack(fill="x")

        top_row = ttk.Frame(action_frame)
        top_row.pack(fill="x", pady=(0, 6))
        self.mail_btn_structure = ttk.Button(top_row, text="구조화 분석", command=self.on_mail_structure)
        self.mail_btn_structure.pack(side="left", padx=(0, 6))
        self.mail_btn_summary = ttk.Button(top_row, text="핵심 요약", command=self.on_mail_summary)
        self.mail_btn_summary.pack(side="left", padx=(0, 6))
        self.mail_btn_reply = ttk.Button(top_row, text="자동 답장문구 생성", command=self.on_mail_reply)
        self.mail_btn_reply.pack(side="left", padx=(0, 6))

        bottom_row = ttk.Frame(action_frame)
        bottom_row.pack(fill="x")
        self.mail_btn_inspect = ttk.Button(bottom_row, text="표현 점검", command=self.on_mail_inspect)
        self.mail_btn_inspect.pack(side="left", padx=(0, 6))
        self.mail_btn_html = ttk.Button(bottom_row, text="발표용 HTML 생성", command=self.on_mail_html_generate)
        self.mail_btn_html.pack(side="left", padx=(0, 6))
        self.mail_btn_clear = ttk.Button(bottom_row, text="입력 지우기", command=self.on_mail_clear)
        self.mail_btn_clear.pack(side="left")

    def _build_mail_result_section(self, parent: ttk.Frame):
        result_wrap = ttk.LabelFrame(parent, text="결과 워크스페이스", padding=10)
        result_wrap.pack(fill="both", expand=True)

        header = ttk.Frame(result_wrap)
        header.pack(fill="x", pady=(0, 8))
        left_meta = ttk.Frame(header)
        left_meta.pack(side="left", fill="x", expand=True)
        ttk.Label(left_meta, textvariable=self.mail_result_type_var, font=("맑은 고딕", 11, "bold")).pack(anchor="w")
        ttk.Label(left_meta, textvariable=self.mail_result_meta_var, foreground="#47617c", wraplength=760).pack(anchor="w", pady=(2, 0))

        result_btn_frame = ttk.Frame(header)
        result_btn_frame.pack(side="right")
        self.mail_btn_copy = ttk.Button(result_btn_frame, text="결과 복사", command=self.on_mail_copy)
        self.mail_btn_copy.pack(side="left", padx=(0, 6))
        self.mail_btn_copy_html = ttk.Button(result_btn_frame, text="HTML만 복사", command=self.on_mail_copy_html)
        self.mail_btn_copy_html.pack(side="left", padx=(0, 6))
        self.mail_btn_save_html = ttk.Button(result_btn_frame, text="HTML 저장", command=self.on_mail_save_html)
        self.mail_btn_save_html.pack(side="left", padx=(0, 6))
        self.mail_btn_clear_result = ttk.Button(result_btn_frame, text="결과 지우기", command=self.on_mail_clear_result)
        self.mail_btn_clear_result.pack(side="left")

        status_frame = ttk.Frame(result_wrap)
        status_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(status_frame, text="상태:", font=("맑은 고딕", 10, "bold")).pack(side="left")
        ttk.Label(status_frame, textvariable=self.mail_status_var).pack(side="left", padx=(6, 0))

        self.mail_result_text = scrolledtext.ScrolledText(
            result_wrap,
            wrap="word",
            font=("Consolas", 11),
            height=34,
            spacing1=2,
            spacing3=5,
            padx=12,
            pady=12,
        )
        self.mail_result_text.pack(fill="both", expand=True)
        self._update_mail_result_action_buttons()

    def _render_metric_checkboxes(self, metrics: list[str], default_selected: list[str]):
        for checkbutton in self.metric_checkbuttons:
            checkbutton.destroy()
        self.metric_checkbuttons.clear()
        self.metric_vars.clear()

        for idx, metric in enumerate(metrics):
            var = tk.BooleanVar(value=(metric in default_selected))
            self.metric_vars[metric] = var
            check = ttk.Checkbutton(self.metric_checks_frame, text=metric, variable=var)
            check.grid(row=idx // 4, column=idx % 4, sticky="w", padx=(0, 12), pady=2)
            self.metric_checkbuttons.append(check)

    def _render_merge_preview(self, summary: dict):
        lines = []
        for candidate in summary.get("typo_candidates", []):
            lines.append(
                f"[typo] {candidate['column']}: {candidate['canonical_candidate']} <= {', '.join(candidate['similar_values'])} ({candidate['reason']})"
            )
        for candidate in summary.get("merge_candidates", []):
            lines.append(
                f"[merge] {candidate['column']}: {candidate['canonical_candidate']} <= {', '.join(candidate['merge_values'])}"
            )
        if not lines:
            lines = ["후보 없음"]

        self.merge_text.configure(state="normal")
        self.merge_text.delete("1.0", "end")
        self.merge_text.insert("1.0", "\n".join(lines))
        self.merge_text.configure(state="disabled")

    def _render_preview_tree(self, selection_data: dict):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.preview_tree["columns"] = ()

        raw_values = selection_data.get("raw_values") or []
        if not raw_values:
            return

        col_count = max(len(row) for row in raw_values)
        columns = [f"C{idx}" for idx in range(1, col_count + 1)]
        self.preview_tree["columns"] = columns
        for idx, column in enumerate(columns):
            self.preview_tree.heading(column, text=f"COL_{idx + 1}")
            self.preview_tree.column(column, width=120, anchor="w")

        for row in raw_values[:8]:
            values = row + [None] * (col_count - len(row))
            self.preview_tree.insert("", "end", values=values)

    def _build_config_from_ui(self) -> AnalysisConfig:
        selected_metrics = [metric for metric, var in self.metric_vars.items() if var.get()]
        if not selected_metrics:
            raise RuntimeError("metric column이 하나도 선택되지 않았습니다.")

        return AnalysisConfig(
            selected_label_column=normalize_config_value(self.label_var.get()),
            selected_time_column=None if self.time_var.get() == "자동 추론 사용" else self.time_var.get(),
            selected_metric_columns=selected_metrics,
            exclude_total_rows=self.exclude_total_rows_var.get(),
            apply_merge_candidates=self.apply_merge_candidates_var.get(),
            use_first_row_as_header=self.use_first_row_header_var.get(),
        )

    def _refresh_analysis_settings(self, selection_data: dict):
        previous_label = self.label_var.get()
        previous_time = self.time_var.get()
        previous_metric_selection = {metric: var.get() for metric, var in self.metric_vars.items()}

        preview_config = AnalysisConfig(
            use_first_row_as_header=self.use_first_row_header_var.get(),
            exclude_total_rows=self.exclude_total_rows_var.get(),
        )
        table = selection_to_table(selection_data, config=preview_config)
        candidates = build_analysis_candidates(table)
        summary = candidates["summary"]
        self.auto_summary = summary
        self.selection_info_var.set(
            f"Workbook={selection_data['workbook_name']} | Sheet={selection_data['sheet_name']} | Range={selection_data['address']}"
        )
        self.headers_var.set(", ".join(table["headers"]) if table["headers"] else "-")
        self.auto_label_var.set(summary.get("label_column") or "-")
        self.auto_time_var.set(summary.get("time_columns", ["-"])[0] if summary.get("time_columns") else "-")
        self.auto_metrics_var.set(", ".join(summary.get("numeric_columns", [])) or "-")
        self.auto_total_rows_var.set(str(summary.get("total_rows_count", 0)))

        self.label_combo["values"] = candidates["label_candidates"]
        self.time_combo["values"] = candidates["time_candidates"]
        self.label_var.set(previous_label if previous_label in candidates["label_candidates"] else "자동 추론 사용")
        self.time_var.set(previous_time if previous_time in candidates["time_candidates"] else "자동 추론 사용")
        default_metrics = [
            metric for metric in candidates["metric_candidates"]
            if previous_metric_selection.get(metric, metric in summary.get("numeric_columns", []))
        ]
        self._render_metric_checkboxes(candidates["metric_candidates"], default_metrics)
        self._render_merge_preview(summary)
        self._render_preview_tree(selection_data)

    def on_config_option_changed(self):
        if self.selection_data:
            self._refresh_analysis_settings(self.selection_data)

    def on_load_selection(self):
        try:
            self.set_status("선택영역을 읽는 중...")
            selection_data = get_current_excel_selection()
            self.selection_data = selection_data
            self._refresh_analysis_settings(selection_data)
            self.set_status("선택영역과 분석 설정을 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.set_status("오류 발생")

    def set_status(self, text: str):
        self.excel_status_var.set(text)
        self.root.update_idletasks()

    def append_text(self, text: str):
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", text)
        self.txt.see("1.0")

    def set_mail_status(self, text: str):
        self.mail_status_var.set(text)
        self.root.update_idletasks()

    def append_mail_result(self, text: str):
        self.mail_result_text.delete("1.0", "end")
        self.mail_result_text.insert("1.0", text)
        self.mail_result_text.see("1.0")

    def _result_type_label(self, result_type: str | None) -> str:
        mapping = {
            "structure": "구조화 분석",
            "summary": "핵심 요약",
            "reply": "자동 답장문구",
            "inspect": "표현 점검",
            "html": "발표용 HTML",
            None: "없음",
        }
        return mapping.get(result_type, str(result_type))

    def _update_mail_result_action_buttons(self):
        html_state = "normal" if self.current_result_type == "html" and self.mail_last_html else "disabled"
        if hasattr(self, "mail_btn_copy_html"):
            self.mail_btn_copy_html.config(state=html_state)
        if hasattr(self, "mail_btn_save_html"):
            self.mail_btn_save_html.config(state=html_state)

    def _refresh_mail_result_meta(self):
        meta_parts = []
        if self.current_mail_meta.get("date"):
            meta_parts.append(f"날짜: {self.current_mail_meta['date']}")
        if self.current_mail_meta.get("sender"):
            meta_parts.append(f"발신자: {self.current_mail_meta['sender']}")
        if self.current_mail_meta.get("subject"):
            meta_parts.append(f"제목: {self.current_mail_meta['subject']}")
        self.mail_result_meta_var.set(" | ".join(meta_parts) if meta_parts else "메일 메타: 없음")

    def _set_mail_result(self, text: str, result_type: str, mail_meta: dict | None = None):
        self.current_result_type = result_type
        if mail_meta:
            self.current_mail_meta = dict(mail_meta)
        self.mail_result_type_var.set(f"현재 결과: {self._result_type_label(result_type)}")
        self._refresh_mail_result_meta()
        self.append_mail_result(text)
        self._update_mail_result_action_buttons()

    def _set_mail_fetch_controls_state(self, state: str):
        self.mail_btn_fetch_list.config(state=state)
        self.mail_btn_load_selected.config(state=state)
        self.mail_btn_clear_login.config(state=state)

    def _clear_mail_list(self):
        for item_id in self.mail_list_tree.get_children():
            self.mail_list_tree.delete(item_id)
        self.mail_items = []
        self.mail_selected_item = None
        self.selected_mail_info_var.set("선택 메일: 없음")

    def _render_mail_items(self, items: list[MailItem]):
        self._clear_mail_list()
        self.mail_items = list(items)
        for item in items:
            self.mail_list_tree.insert(
                "",
                "end",
                iid=str(item.index),
                values=(item.date_str, item.sender, item.subject),
            )
        if items:
            first_id = str(items[0].index)
            self.mail_list_tree.selection_set(first_id)
            self.mail_list_tree.focus(first_id)
            self.on_mail_result_select()

    def _get_mail_query_params(self) -> MailQueryParams:
        user_id = self.mail_fetch_user_var.get().strip()
        password = self.mail_fetch_password_var.get()
        keyword = self.mail_search_var.get().strip()
        try:
            recent_days = max(1, int(self.mail_recent_days_var.get().strip() or "7"))
        except ValueError as e:
            raise RuntimeError("조회 기간은 숫자로 입력하세요.") from e
        try:
            max_count = max(1, min(100, int(self.mail_max_count_var.get().strip() or "20")))
        except ValueError as e:
            raise RuntimeError("최대 조회 개수는 숫자로 입력하세요.") from e
        return MailQueryParams(
            user_id=user_id,
            password=password,
            keyword=keyword,
            recent_days=recent_days,
            max_count=max_count,
        )

    def _find_mail_item_by_tree_selection(self) -> MailItem | None:
        selected = self.mail_list_tree.selection()
        if not selected:
            return None
        selected_id = selected[0]
        for item in self.mail_items:
            if str(item.index) == selected_id:
                return item
        return None

    def _build_current_mail_meta(self, subject: str, sender: str = "", date_str: str = "") -> dict[str, str]:
        return {
            "subject": subject.strip() if subject else "",
            "sender": sender.strip() if sender else "",
            "date": date_str.strip() if date_str else "",
        }

    def _apply_mail_item_to_inputs(self, item: MailItem):
        current_subject = self.mail_subject_entry.get().strip()
        current_body = self.mail_body_text.get("1.0", "end").strip()
        if current_subject or current_body:
            overwrite = messagebox.askyesno(
                "메일 불러오기",
                "현재 입력된 제목/본문을 선택한 메일 내용으로 덮어쓸까요?",
            )
            if not overwrite:
                return

        self.mail_subject_entry.delete(0, "end")
        self.mail_subject_entry.insert(0, item.subject)
        self.mail_body_text.delete("1.0", "end")
        self.mail_body_text.insert("1.0", item.body)
        self.mail_structured_info = None
        self.mail_inspection_info = None
        self.mail_last_html = ""
        self.current_mail_meta = self._build_current_mail_meta(item.subject, item.sender, item.date_str)
        self.selected_mail_info_var.set(f"선택 메일: [{item.date_str}] {item.sender} | {item.subject}")
        self._refresh_mail_result_meta()
        self.set_mail_status("메일 불러오기 완료")

    def on_run(self):
        if not self.selection_data:
            self.on_load_selection()
            if not self.selection_data:
                return

        self.btn_run.config(state="disabled")
        self.set_status("분석 중... 설정 반영 후 LLM 호출 중")

        def worker():
            try:
                mode = self.mode_var.get()
                config = self._build_config_from_ui()
                result_text = run_selection_analysis(mode, config=config, selection_data=self.selection_data)
                self.root.after(0, lambda: self.append_text(result_text))
                self.root.after(0, lambda: self.set_status("완료"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.root.after(0, lambda: self.set_status("오류 발생"))
            finally:
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _get_mail_inputs(self) -> tuple[str, str]:
        subject = self.mail_subject_entry.get().strip()
        body = normalize_mail_text(self.mail_body_text.get("1.0", "end"))
        if not body:
            messagebox.showwarning("입력 필요", "메일 본문을 입력하세요.")
            raise RuntimeError("메일 본문이 비어 있습니다.")
        return subject, body

    def _run_mail_action(self, action_name: str, mode: str):
        try:
            subject, body = self._get_mail_inputs()
        except RuntimeError:
            return

        self.mail_btn_structure.config(state="disabled")
        self.mail_btn_summary.config(state="disabled")
        self.mail_btn_reply.config(state="disabled")
        self.mail_btn_inspect.config(state="disabled")
        self.mail_btn_html.config(state="disabled")
        self.set_mail_status(action_name)
        mail_meta = self.current_mail_meta or self._build_current_mail_meta(subject)

        def worker():
            try:
                analysis_result = run_mail_analysis(
                    mode,
                    subject,
                    body,
                    structured_info=self.mail_structured_info,
                    inspection_info=self.mail_inspection_info,
                )
                self.mail_structured_info = analysis_result.get("structured_info")
                self.mail_inspection_info = analysis_result.get("inspection_info", self.mail_inspection_info)
                if mode != "html":
                    self.mail_last_html = ""
                self.root.after(0, lambda: self._set_mail_result(analysis_result["result_text"], mode, mail_meta=mail_meta))
                self.root.after(0, lambda: self.set_mail_status("완료"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.root.after(0, lambda: self.set_mail_status("오류 발생"))
            finally:
                self.root.after(0, lambda: self.mail_btn_structure.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_summary.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_reply.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_inspect.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_html.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def on_mail_structure(self):
        self._run_mail_action("구조화 분석 중...", "structure")

    def on_mail_summary(self):
        self._run_mail_action("요약 생성 중...", "summary")

    def on_mail_reply(self):
        self._run_mail_action("답장문구 생성 중...", "reply")

    def on_mail_inspect(self):
        self._run_mail_action("표현 점검 중...", "inspect")

    def on_mail_html_generate(self):
        try:
            subject, body = self._get_mail_inputs()
        except RuntimeError:
            return

        style_label = self.mail_html_style_var.get()
        html_style = MAIL_HTML_STYLE_MAP.get(style_label, "briefing")

        self.mail_btn_structure.config(state="disabled")
        self.mail_btn_summary.config(state="disabled")
        self.mail_btn_reply.config(state="disabled")
        self.mail_btn_inspect.config(state="disabled")
        self.mail_btn_html.config(state="disabled")
        self.set_mail_status("HTML 생성 중...")
        mail_meta = self.current_mail_meta or self._build_current_mail_meta(subject)

        def worker():
            try:
                analysis_result = run_mail_analysis(
                    "html",
                    subject,
                    body,
                    html_style=html_style,
                    structured_info=self.mail_structured_info,
                    inspection_info=self.mail_inspection_info,
                )
                self.mail_structured_info = analysis_result.get("structured_info")
                self.mail_inspection_info = analysis_result.get("inspection_info", self.mail_inspection_info)
                self.mail_last_html = analysis_result.get("html_text", "")
                self.root.after(0, lambda: self._set_mail_result(analysis_result["result_text"], "html", mail_meta=mail_meta))
                self.root.after(0, lambda: self.set_mail_status("완료"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.root.after(0, lambda: self.set_mail_status("오류 발생"))
            finally:
                self.root.after(0, lambda: self.mail_btn_structure.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_summary.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_reply.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_inspect.config(state="normal"))
                self.root.after(0, lambda: self.mail_btn_html.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def on_mail_fetch_list(self):
        try:
            query_params = self._get_mail_query_params()
        except RuntimeError as e:
            messagebox.showwarning("입력 확인", str(e))
            return

        self._set_mail_fetch_controls_state("disabled")
        self.set_mail_status("메일함 조회 중...")

        def worker():
            try:
                items = fetch_recent_mails(query_params)
                if not items:
                    self.root.after(0, self._clear_mail_list)
                    self.root.after(0, lambda: self.set_mail_status("조회 결과 없음"))
                    self.root.after(0, lambda: messagebox.showinfo("조회 결과", "조건에 맞는 메일이 없습니다."))
                    return
                self.root.after(0, lambda: self._render_mail_items(items))
                self.root.after(0, lambda: self.set_mail_status("조회 완료"))
            except Exception as e:
                self.root.after(0, self._clear_mail_list)
                error_text = str(e)
                status_text = "로그인 실패" if "로그인 실패" in error_text else "조회 실패"
                self.root.after(0, lambda: messagebox.showerror("메일 조회 오류", error_text))
                self.root.after(0, lambda: self.set_mail_status(status_text))
            finally:
                self.root.after(0, lambda: self._set_mail_fetch_controls_state("normal"))

        threading.Thread(target=worker, daemon=True).start()

    def on_mail_result_select(self, event=None):
        item = self._find_mail_item_by_tree_selection()
        self.mail_selected_item = item
        if not item:
            self.selected_mail_info_var.set("선택 메일: 없음")
            return
        self.selected_mail_info_var.set(f"선택 메일: [{item.date_str}] {item.sender} | {item.subject}")

    def on_mail_load_selected(self, event=None):
        item = self._find_mail_item_by_tree_selection()
        if not item:
            messagebox.showinfo("안내", "불러올 메일을 목록에서 먼저 선택하세요.")
            return
        self._apply_mail_item_to_inputs(item)

    def on_mail_clear_login(self):
        self.mail_fetch_password_var.set("")
        self.mail_fetch_user_var.set("")
        self.set_mail_status("로그인 정보를 지웠습니다.")

    def on_mail_clear(self):
        self.mail_subject_entry.delete(0, "end")
        self.mail_body_text.delete("1.0", "end")
        self.mail_structured_info = None
        self.mail_inspection_info = None
        self.mail_last_html = ""
        self.current_result_type = None
        self.current_mail_meta = {}
        self.selected_mail_info_var.set("선택 메일: 없음")
        self.mail_result_type_var.set("현재 결과: 없음")
        self.mail_result_meta_var.set("메일 메타: 없음")
        self.on_mail_clear_result()
        self.set_mail_status("준비됨")

    def on_mail_copy(self):
        content = self.mail_result_text.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("안내", "복사할 결과가 없습니다.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.set_mail_status("결과를 클립보드에 복사했습니다.")

    def on_mail_clear_result(self):
        self.mail_result_text.delete("1.0", "end")
        self.current_result_type = None
        self.mail_last_html = ""
        self.mail_result_type_var.set("현재 결과: 없음")
        self._update_mail_result_action_buttons()

    def on_mail_copy_html(self):
        if not self.mail_last_html:
            messagebox.showinfo("안내", "복사할 HTML 결과가 없습니다.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(self.mail_last_html)
        self.set_mail_status("HTML 복사 완료")

    def on_mail_save_html(self):
        if not self.mail_last_html:
            messagebox.showinfo("안내", "저장할 HTML 결과가 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            title="HTML 저장",
            defaultextension=".html",
            filetypes=[("HTML Files", "*.html"), ("All Files", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fp:
                fp.write(self.mail_last_html)
            self.set_mail_status("HTML 저장 완료")
            messagebox.showinfo("저장 완료", f"HTML 파일을 저장했습니다.\n{path}")
        except Exception as e:
            messagebox.showerror("저장 실패", str(e))
            self.set_mail_status("오류 발생")

    def on_copy(self):
        content = self.txt.get("1.0", "end").strip()
        if not content:
            messagebox.showinfo("안내", "복사할 결과가 없습니다.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.set_status("결과를 클립보드에 복사했습니다.")

    def on_clear(self):
        self.txt.delete("1.0", "end")
        self.set_status("결과를 지웠습니다.")

    def on_help(self):
        msg = (
            "1. Excel 데스크톱 앱에서 파일을 연다.\n"
            "2. 요약할 범위를 드래그해서 선택한다.\n"
            "3. '선택영역 불러오기'로 미리보기와 자동 추론 결과를 확인한다.\n"
            "4. label/time/metric/total row/merge/header 옵션을 필요시 수정한다.\n"
            "5. 분석 유형을 고른 뒤 '현재 설정으로 분석 실행'을 누른다.\n\n"
            "주의:\n"
            "- Windows + Excel 데스크톱 환경 기준\n"
            "- Excel Online / LibreOffice는 지원하지 않음\n"
            "- 최소 2행 이상 선택 필요"
        )
        messagebox.showinfo("사용 방법", msg)


def main():
    if os.getenv("RUN_MOCK_TESTS") == "1" or "--mock-test" in sys.argv:
        run_mock_analysis_tests()
        return
    if os.getenv("RUN_MOCK_CONFIG_TESTS") == "1" or "--mock-config-test" in sys.argv:
        run_mock_config_tests()
        return
    if os.getenv("RUN_MAIL_PROMPT_TESTS") == "1" or "--mail-prompt-test" in sys.argv:
        test_mail_prompts()
        return
    if os.getenv("RUN_MAIL_STRUCTURE_PROMPT_TESTS") == "1" or "--mail-structure-prompt-test" in sys.argv:
        test_mail_structure_prompts()
        return
    if os.getenv("RUN_MAIL_SUMMARY_PROMPT_TESTS") == "1" or "--mail-summary-prompt-test" in sys.argv:
        test_mail_summary_prompts()
        return
    if os.getenv("RUN_MAIL_HTML_CONTEXT_TESTS") == "1" or "--mail-html-context-test" in sys.argv:
        test_mail_html_context()
        return
    if os.getenv("RUN_MAIL_HTML_RENDER_TESTS") == "1" or "--mail-html-render-test" in sys.argv:
        test_mail_html_render()
        return
    if os.getenv("RUN_MAIL_HTML_RENDER_STYLE_TESTS") == "1" or "--mail-html-style-test" in sys.argv:
        test_mail_html_render_styles()
        return
    if os.getenv("RUN_MAIL_FILTER_TESTS") == "1" or "--mail-filter-test" in sys.argv:
        test_mail_filtering()
        return
    if os.getenv("RUN_MAIL_TRIM_TESTS") == "1" or "--mail-trim-test" in sys.argv:
        test_trim_mail_body()
        return
    if os.getenv("RUN_MAIL_HTML_TO_TEXT_TESTS") == "1" or "--mail-html-to-text-test" in sys.argv:
        test_html_to_text_basic()
        return
    if os.getenv("RUN_MAIL_RESULT_TYPE_TESTS") == "1" or "--mail-result-type-test" in sys.argv:
        test_result_type_state()
        return
    root = tk.Tk()
    ExcelLLMApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
