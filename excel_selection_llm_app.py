from __future__ import annotations

import os
import re
import json
import sys
import uuid
import math
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


def normalize_text_token(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


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
        if col_type in {"id_like", "text"}:
            values = [row.get(header) for row in rows if not is_missing_value(row.get(header))]
            if not values:
                continue
            unique_ratio = len({str(v) for v in values}) / len(values)
            score = 0
            if col_type == "id_like":
                score += 3
            if 0.3 <= unique_ratio <= 1.0:
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
    if row_count <= 0 or profile["unique_count"] == 0:
        return None
    if profile["unique_count"] > max(12, int(row_count * 0.5)):
        return None
    top_values = profile.get("top_values") or []
    if not top_values:
        return None
    top_count = top_values[0]["count"]
    ratio = top_count / row_count
    if ratio >= 0.7:
        return f"{header} 컬럼은 '{top_values[0]['value']}'가 {ratio:.0%}로 편중되어 있습니다."
    return None


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
        "numeric_highlights": numeric_highlights,
        "missing_data_warnings": missing_data_warnings,
        "skew_warnings": skew_warnings,
        "total_row_warnings": total_row_warnings,
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
    table_topic_guess = guess_table_topic(headers)

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

        if summary["null_ratios"][h] >= 0.3:
            summary["warnings"].append(f"{h} 컬럼의 결측률이 {summary['null_ratios'][h]:.0%}로 높습니다.")

        if col_type in {"text", "id_like", "date_like", "month_like"}:
            skew_warning = analyze_categorical_skew(h, summary["column_profiles"][h], len(analysis_rows))
            if skew_warning:
                summary["warnings"].append(skew_warning)

    if total_row_indexes:
        summary["warnings"].append(
            f"합계행/총계행 후보 {len(total_row_indexes)}개가 감지되어 일반 분포 계산에서 제외했습니다. row_indexes={total_row_indexes}"
        )

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
    if summary.get("time_columns"):
        lines.append(f"[시간 컬럼 후보] {', '.join(summary['time_columns'])}")
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

    if summary.get("warnings"):
        lines.append("-" * 80)
        lines.append("[주의 사항]")
        for warning in summary["warnings"][:8]:
            lines.append(f"  - {warning}")

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
1. 표가 무엇을 나타내는지 한 줄로 추정
2. 핵심 포인트 3개
3. 이상치/누락/주의점이 있으면 지적
4. 실무자 보고용 한줄 코멘트
""",
        "report": """
요구사항:
1. 임원/리더 보고용으로 자연스럽게 5~8줄 내외로 정리
2. 핵심 수치가 있으면 언급
3. 과장 금지, 데이터 품질 이슈가 있으면 함께 언급
4. 마지막에 '권장 액션' 2개 제시
""",
        "risk": """
요구사항:
1. 데이터의 문제점/주의점/이상치 중심으로 분석
2. 누락값, 편중, 튀는 수치, 해석상 한계를 우선 설명
3. 실무자가 확인해야 할 체크포인트 3개 제시
"""
    }

    payload = {
        "workbook_name": selection_data["workbook_name"],
        "sheet_name": selection_data["sheet_name"],
        "table_range": table["table_range"],
        "insight_summary": insight_summary,
        "summary": summary,
        "preview_rows": preview_rows,
    }

    safe_payload = make_json_safe(payload)

    return f"""
    다음은 현재 사용자가 Excel에서 선택한 표 영역을 읽어 요약한 데이터입니다.
    이 정보를 바탕으로 한국어로 알기 쉽게 해석해 주세요.
    우선순위는 1) insight_summary 2) summary 3) preview_rows 입니다.
    모르는 것은 반드시 '추정'이라고 표시하고, 없는 사실을 만들지 마세요.
    숫자 근거를 우선 사용하세요.
    합계행/총계행은 일반 데이터와 구분해서 해석하세요.
    결측률/편중/이상치가 있으면 먼저 언급하세요.

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
            "name": "financial_monthly",
            "headers": ["월", "매출", "원가", "이익률", "LOT_ID"],
            "rows": [
                {"월": "2026-01", "매출": 1200, "원가": 900, "이익률": 0.25, "LOT_ID": "LOT-001"},
                {"월": "2026-02", "매출": 1500, "원가": 1100, "이익률": 0.27, "LOT_ID": "LOT-002"},
                {"월": "2026-03", "매출": None, "원가": 1000, "이익률": 0.0, "LOT_ID": "LOT-003"},
                {"월": "총계", "매출": 2700, "원가": 3000, "이익률": 0.18, "LOT_ID": "TOTAL"},
            ],
            "table_range": {"address": "A1:E5"},
        },
        {
            "name": "quality_daily",
            "headers": ["일자", "라인", "수율", "불량코드", "비고"],
            "rows": [
                {"일자": "2026/03/01", "라인": "A", "수율": 0.98, "불량코드": "D01", "비고": None},
                {"일자": "2026/03/02", "라인": "A", "수율": 0.97, "불량코드": "D01", "비고": None},
                {"일자": "2026/03/03", "라인": "A", "수율": 0.96, "불량코드": "D02", "비고": "점검"},
                {"일자": "2026/03/04", "라인": "B", "수율": 0.89, "불량코드": "D01", "비고": None},
                {"일자": "합계", "라인": None, "수율": 0.95, "불량코드": None, "비고": None},
            ],
            "table_range": {"address": "A1:E6"},
        },
    ]


def run_mock_analysis_tests():
    print("[Mock analysis tests]")
    for table in get_mock_tables():
        summary = summarize_table(table)
        insight_summary = build_insight_summary(summary)
        print(f"\n=== {table['name']} ===")
        print("[summary]")
        print(json.dumps(make_json_safe(summary), ensure_ascii=False, indent=2))
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
