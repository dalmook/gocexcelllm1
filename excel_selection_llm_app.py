from __future__ import annotations

import os
import re
import json
import sys
import uuid
import math
import difflib
from datetime import datetime, date
import threading
import tkinter as tk
from tkinter import ttk, messagebox

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

        headers_raw = values[0]
        headers = []
        used = {}

        for i, h in enumerate(headers_raw, start=1):
            name = clean_header_name(normalize_cell_value(h), i)
            if name in used:
                used[name] += 1
                name = f"{name}_{used[name]}"
            else:
                used[name] = 1
            headers.append(name)

        rows = []
        for raw_row in values[1:]:
            item = {}
            has_any = False
            for idx, cell_value in enumerate(raw_row):
                v = normalize_cell_value(cell_value)
                item[headers[idx]] = v
                if v is not None:
                    has_any = True
            if has_any:
                rows.append(item)

        return {
            "workbook_name": wb.Name,
            "sheet_name": ws.Name,
            "address": sel.Address,
            "headers": headers,
            "rows": rows,
        }
    finally:
        pythoncom.CoUninitialize()

def selection_to_table(selection_data: dict) -> dict:
    return {
        "headers": selection_data["headers"],
        "rows": selection_data["rows"],
        "table_range": {
            "address": selection_data["address"]
        }
    }


# =========================================================
# 5) 표 요약
# =========================================================

def summarize_table(table: dict, top_n: int = 5) -> dict:
    headers = table["headers"]
    rows = table["rows"]
    total_row_indexes = detect_total_row_indexes(rows, headers)
    data_rows = [row for idx, row in enumerate(rows) if idx not in total_row_indexes]
    data_indexes = [idx for idx in range(len(rows)) if idx not in total_row_indexes]
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
    label_column = choose_label_column(headers, analysis_rows, detected_types)
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

    summary["total_row_metrics"], summary["detected_totals"] = build_total_row_metrics(
        total_rows=total_rows,
        numeric_columns=summary["numeric_columns"],
        label_column=label_column,
    )
    summary["trend_analysis"] = build_trend_analysis(summary, analysis_rows, headers)
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


def run_selection_analysis(mode: str) -> str:
    selection_data = get_current_excel_selection()
    table = selection_to_table(selection_data)

    if not table["rows"]:
        raise RuntimeError("선택 범위에 데이터 행이 없습니다.")

    summary = summarize_table(table)
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

    return f"{basic_summary}\n\n[LLM 결과]\n{llm_text}"


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


# =========================================================
# 7) Tkinter UI
# =========================================================

class ExcelLLMApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel 선택영역 LLM 요약기")
        self.root.geometry("1100x760")

        self.status_var = tk.StringVar(value="준비됨")
        self.mode_var = tk.StringVar(value="summary")

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=12)
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

        self.btn_run = ttk.Button(btn_frame, text="선택영역 분석 실행", command=self.on_run)
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
        ttk.Label(status_frame, textvariable=self.status_var).pack(side="left", padx=(6, 0))

        result_frame = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        result_frame.pack(fill="both", expand=True)

        self.txt = tk.Text(result_frame, wrap="word", font=("Consolas", 10))
        self.txt.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(result_frame, orient="vertical", command=self.txt.yview)
        scroll.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=scroll.set)

    def set_status(self, text: str):
        self.status_var.set(text)
        self.root.update_idletasks()

    def append_text(self, text: str):
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", text)
        self.txt.see("1.0")

    def on_run(self):
        self.btn_run.config(state="disabled")
        self.set_status("분석 중... Excel 선택영역을 읽고 LLM 호출 중")

        def worker():
            try:
                mode = self.mode_var.get()
                result_text = run_selection_analysis(mode)
                self.root.after(0, lambda: self.append_text(result_text))
                self.root.after(0, lambda: self.set_status("완료"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.root.after(0, lambda: self.set_status("오류 발생"))
            finally:
                self.root.after(0, lambda: self.btn_run.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

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
            "3. 첫 행은 헤더, 아래 행은 데이터여야 한다.\n"
            "4. 이 앱에서 분석 유형을 고른 뒤 '선택영역 분석 실행'을 누른다.\n\n"
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
    root = tk.Tk()
    ExcelLLMApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
