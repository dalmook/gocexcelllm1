from __future__ import annotations

import os
import re
import json
import uuid
import math
from datetime import datetime, date
import threading
import tkinter as tk
from tkinter import ttk, messagebox

import requests
import win32com.client
import pythoncom


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


# =========================================================
# 4) 현재 Excel 선택 영역 읽기
# =========================================================

def get_current_excel_selection() -> dict:
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

    summary = {
        "row_count": len(rows),
        "column_count": len(headers),
        "headers": headers,
        "numeric_columns": [],
        "text_columns": [],
        "null_counts": {},
        "column_profiles": {},
    }

    col_values = {h: [] for h in headers}
    for row in rows:
        for h in headers:
            col_values[h].append(row.get(h))

    for h in headers:
        values = col_values[h]
        null_count = sum(1 for v in values if v is None)
        summary["null_counts"][h] = null_count

        numeric_vals = [safe_float(v) for v in values]
        numeric_vals = [v for v in numeric_vals if v is not None]

        if len(values) > 0 and len(numeric_vals) >= max(2, int(len(values) * 0.5)):
            summary["numeric_columns"].append(h)
            total = sum(numeric_vals)
            avg = total / len(numeric_vals) if numeric_vals else None
            min_v = min(numeric_vals) if numeric_vals else None
            max_v = max(numeric_vals) if numeric_vals else None

            sorted_vals = sorted(numeric_vals)
            median = None
            if sorted_vals:
                n = len(sorted_vals)
                if n % 2 == 1:
                    median = sorted_vals[n // 2]
                else:
                    median = (sorted_vals[n // 2 - 1] + sorted_vals[n // 2]) / 2

            summary["column_profiles"][h] = {
                "type": "numeric",
                "count": len(numeric_vals),
                "null_count": null_count,
                "sum": total,
                "avg": avg,
                "min": min_v,
                "max": max_v,
                "median": median,
            }
        else:
            summary["text_columns"].append(h)
            freq = {}
            for v in values:
                if v is None:
                    continue
                key = str(v)
                freq[key] = freq.get(key, 0) + 1

            top_values = sorted(freq.items(), key=lambda x: (-x[1], x[0]))[:top_n]
            summary["column_profiles"][h] = {
                "type": "text",
                "count": len(values) - null_count,
                "null_count": null_count,
                "unique_count": len(freq),
                "top_values": [{"value": k, "count": v} for k, v in top_values],
            }

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
    lines.append(f"[열 수] {summary['column_count']}")
    lines.append(f"[헤더] {', '.join(summary['headers'])}")
    lines.append("-" * 80)

    lines.append("[숫자형 컬럼]")
    if not summary["numeric_columns"]:
        lines.append("  - 없음")
    else:
        for col in summary["numeric_columns"]:
            p = summary["column_profiles"][col]
            lines.append(
                f"  - {col}: count={p['count']}, null={p['null_count']}, "
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
                f"unique={p['unique_count']}, top={top_vals}"
            )

    lines.append("=" * 80)
    return "\n".join(lines)


# =========================================================
# 6) 프롬프트
# =========================================================

def build_llm_prompt(selection_data: dict, table: dict, summary: dict, preview_rows: list, mode: str) -> str:
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
        "summary": summary,
        "preview_rows": preview_rows,
    }

    safe_payload = make_json_safe(payload)
    json.dumps(safe_payload, ensure_ascii=False, indent=2)

    return f"""
    다음은 현재 사용자가 Excel에서 선택한 표 영역을 읽어 요약한 데이터입니다.
    이 정보를 바탕으로 한국어로 알기 쉽게 해석해 주세요.
    모르는 것은 추정이라고 표시하고, 없는 사실을 만들지 마세요.

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
    root = tk.Tk()
    app = ExcelLLMApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
