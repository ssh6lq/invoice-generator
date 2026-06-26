"""
server.py — 청구서 자동 작성 (FastAPI)

Streamlit 대신 FastAPI로 다시 만든 버전. 사내망 다중 사용자용.
  · 백엔드는 무상태(stateless): 각 요청마다 계산만 하고, 화면 상태는 브라우저(JS)가 보관.
  · 핵심 로직은 기존 모듈을 그대로 재사용한다:
      receipt_parser / excel_filler / overtime_filler

실행:
    .venv\\Scripts\\uvicorn server:app --host 0.0.0.0 --port 8000
브라우저: http://localhost:8000  (사내망: http://<이 PC IP>:8000)
"""

import io
import os
import json
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()  # 같은 폴더의 .env에서 AI 연결 설정(주소·모델·토큰)을 읽는다.
except Exception:  # noqa: BLE001
    pass

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from langchain_core.messages import HumanMessage, SystemMessage

from receipt_parser import Receipt, SYSTEM_PROMPT, _image_to_data_url
from excel_filler import (
    fill_workbook, get_dropdown_options, get_support_limits, _to_date,
)
from overtime_filler import parse_attendance, fill_overtime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXPENSE_TPL = os.path.join(APP_DIR, "비용청구양식.xlsm")
DEFAULT_OVERTIME_TPL = os.path.join(APP_DIR, "연장근무(수당)신청서_양식.xlsx")
STATIC_DIR = os.path.join(APP_DIR, "static")

XLSM_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# AI 모델 연결 설정 — 관리자가 .env(또는 환경변수)로만 설정한다.
# 사내 사용자에게 주소·모델·토큰이 노출되지 않도록 화면/응답에는 절대 내보내지 않는다.
ENV_PROVIDER = os.getenv("RECEIPT_PROVIDER", "로컬 서버")
ENV_BASE_URL = os.getenv("RECEIPT_BASE_URL",
                         "https://vllm-qwen.proxy.ainexus.ktcloud.com/v1")
ENV_MODEL = os.getenv("RECEIPT_MODEL", "/models/Qwen3.6-35B-A3B-FP8")
# Bearer 토큰: OpenAI 호환 클라이언트는 api_key 값을 'Authorization: Bearer <값>'으로 보낸다.
ENV_API_KEY = (os.getenv("RECEIPT_API_KEY")
               or os.getenv("RECEIPT_BEARER")
               or os.getenv("OPENAI_API_KEY", ""))

app = FastAPI(title="청구서 자동 작성")


@app.middleware("http")
async def no_cache_static(request, call_next):
    """정적 파일(html/js/css)을 브라우저가 캐시해 옛 버전을 쓰는 문제 방지."""
    resp = await call_next(request)
    path = request.url.path
    if path == "/" or path.endswith((".html", ".js", ".css")):
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


# ======================================================================
# 공통 로직 (Streamlit 앱에서 그대로 가져온 순수 함수들)
# ======================================================================
def _build_receipt_llm(provider, model, api_key, base_url):
    """영수증 파싱용 LLM 생성. provider='로컬 서버'면 OpenAI 호환 엔드포인트에 붙는다."""
    from langchain_openai import ChatOpenAI
    if provider == "로컬 서버":
        if model.startswith("openai:"):
            model = model[len("openai:"):]
        llm = ChatOpenAI(
            model=model, base_url=base_url, api_key=(api_key or "EMPTY"),
            temperature=0.1, max_retries=5,
            model_kwargs={"extra_body": {
                "chat_template_kwargs": {"enable_thinking": False},
            }},
        )
        return llm.with_structured_output(Receipt, method="json_schema")
    kwargs = {"model": model, "temperature": 0.0}
    if api_key:
        kwargs["api_key"] = api_key
    llm = ChatOpenAI(**kwargs)
    return llm.with_structured_output(Receipt)


def _parse_receipts(payload, llm):
    """payload=list[(filename, bytes)] -> list[dict]. 한 장씩 파싱한다."""
    results = []
    for fname, content in payload:
        rec = {"filename": fname, "date": None, "store": None,
               "amount": None, "time": None, "region": None, "error": None}
        try:
            data_url = _image_to_data_url(content, fname)
            messages = [
                SystemMessage(content=SYSTEM_PROMPT),
                HumanMessage(content=[
                    {"type": "image_url", "image_url": {"url": data_url}},
                    {"type": "text", "text": "이 영수증 이미지를 자세히 보고, 인쇄된 "
                     "글자를 오타 없이 한 글자도 틀리지 않게 그대로 읽어 정보를 추출해줘. "
                     "특히 상호명과 금액 숫자를 정확히."},
                ]),
            ]
            r = llm.invoke(messages)
            rec.update(date=r.date, store=r.store, amount=r.amount,
                       time=r.time, region=r.region)
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        results.append(rec)
    return results


def _to_int(v):
    """'12,000' / 12000 / '' -> int or None."""
    if v is None or v == "":
        return None
    try:
        return int(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _capped_claim(purpose, claim_amount, receipt_amount, limits):
    """목적별 지원한도를 적용한 청구금액. 한도가 없으면 원래 값."""
    base = claim_amount if claim_amount not in (None, "") else receipt_amount
    base = _to_int(base)
    if base is None:
        return None
    lim = limits.get(purpose) if purpose else None
    return base if lim is None else min(base, int(lim))


def _welfare_alloc(rows, budget):
    """복지비 모드: 영수일자 빠른 순으로 누적해 한도까지 청구금액을 채운다.
    rows 원래 순서에 맞춘 청구금액 리스트를 반환."""
    order = sorted(range(len(rows)), key=lambda i: (
        _to_date(rows[i].get("date")) is None,
        _to_date(rows[i].get("date")) or datetime.max.date(),
    ))
    remaining = int(budget) if budget else 0
    give = {}
    for i in order:
        a = _to_int(rows[i].get("amount")) or 0
        g = min(a, remaining)
        remaining -= g
        give[i] = g
    return [give[i] for i in range(len(rows))]


def _template_bytes(upload_bytes, default_path):
    """업로드 양식이 있으면 그것을, 없으면 번들 기본 양식 bytes를 반환."""
    if upload_bytes:
        return upload_bytes
    if os.path.exists(default_path):
        with open(default_path, "rb") as f:
            return f.read()
    return None


# ======================================================================
# 비용청구 API
# ======================================================================
@app.get("/api/config")
def get_config():
    """사이드바용 정보. AI 연결 주소·모델·토큰은 절대 내보내지 않는다(노출 방지)."""
    return {
        "expense_template": os.path.basename(DEFAULT_EXPENSE_TPL),
        "overtime_template": os.path.basename(DEFAULT_OVERTIME_TPL),
        "ai_ready": bool(ENV_BASE_URL or ENV_API_KEY),  # 연결 설정 존재 여부만
    }


@app.post("/api/expense/options")
async def expense_options(template: UploadFile = File(None)):
    """양식에서 목적/결제방식 드롭다운 + 목적별 한도를 읽어 반환.
    양식을 올리지 않으면 번들 기본 양식을 사용한다."""
    custom = await template.read() if template else None
    tpl = _template_bytes(custom, DEFAULT_EXPENSE_TPL)
    name = (template.filename if template
            else os.path.basename(DEFAULT_EXPENSE_TPL))
    if tpl is None:
        return {"purpose": [], "payment": [], "limits": {},
                "template_name": None, "is_default": True}
    try:
        opts = get_dropdown_options(tpl)
        limits = get_support_limits(tpl)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"양식을 읽지 못했습니다: {e}")
    return {
        "purpose": opts.get("purpose", []),
        "payment": opts.get("payment", []),
        "limits": limits,
        "template_name": name,
        "is_default": template is None,
    }


@app.post("/api/expense/analyze")
async def expense_analyze(
    images: list[UploadFile] = File(...),
    provider: str = Form(None),
    model: str = Form(None),
    api_key: str = Form(None),
    base_url: str = Form(None),
):
    """영수증 이미지들을 GPT 비전으로 분석해 행 목록을 반환한다.
    기본은 서버 .env 설정(사내 모델)을 쓰고, 사용자가 '고급 설정'에서 OpenAI나 직접 로컬을
    고른 경우에만 해당 값으로 덮어쓴다. 값을 안 보내면 .env 기본값이 그대로 쓰인다."""
    payload = [(f.filename, await f.read()) for f in images]
    if not payload:
        raise HTTPException(400, "이미지가 없습니다.")
    llm = _build_receipt_llm(
        provider or ENV_PROVIDER, model or ENV_MODEL,
        api_key or ENV_API_KEY, base_url or ENV_BASE_URL,
    )
    try:
        results = _parse_receipts(payload, llm)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"분석 실패: {e}")
    # 청구금액 기본값 = 영수금액
    rows = [{
        "date": r["date"], "store": r["store"], "purpose": "",
        "amount": r["amount"], "claim": r["amount"], "payment": "",
        "time": r["time"], "region": r["region"], "note": "",
        "filename": r["filename"], "error": r["error"],
    } for r in results]
    return {"rows": rows}


@app.post("/api/expense/generate")
async def expense_generate(
    payload: str = Form(...),
    template: UploadFile = File(None),
):
    """편집된 행으로 비용청구서(.xlsm)를 채워 다운로드로 반환한다.
    payload(JSON): {rows:[...], basic:{dept,name,title}, welfare_budget:int}
    """
    data = json.loads(payload)
    rows = data.get("rows", [])
    basic = data.get("basic", {})
    welfare_budget = _to_int(data.get("welfare_budget")) or 0

    tpl = _template_bytes(await template.read() if template else None,
                          DEFAULT_EXPENSE_TPL)
    if tpl is None:
        raise HTTPException(400, "양식을 찾지 못했습니다.")

    try:
        limits = get_support_limits(tpl)
    except Exception:  # noqa: BLE001
        limits = {}

    # 비어 있지 않은 행만
    rows = [r for r in rows if (r.get("store") or r.get("date"))]
    if not rows:
        raise HTTPException(400, "채울 데이터가 없습니다.")

    alloc = _welfare_alloc(rows, welfare_budget) if welfare_budget else None
    records = []
    for i, r in enumerate(rows):
        claim = (alloc[i] if welfare_budget
                 else _capped_claim(r.get("purpose"), r.get("claim"),
                                    r.get("amount"), limits))
        records.append({
            "date": r.get("date"), "store": r.get("store"),
            "purpose": r.get("purpose"), "amount": _to_int(r.get("amount")),
            "payment": r.get("payment"), "time": r.get("time"),
            "claim_amount": claim, "region": r.get("region"),
            "participants": basic.get("name", ""), "note": r.get("note"),
        })
    try:
        buf, _start, n = fill_workbook(
            tpl, records, append=False,
            basic_info={"dept": basic.get("dept"), "name": basic.get("name"),
                        "title": basic.get("title")},
        )
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"생성 실패: {e}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"비용청구양식_작성완료_{stamp}.xlsm"
    return _download(buf.getvalue(), fname, XLSM_MIME, count=n)


# ======================================================================
# 연장근무 API
# ======================================================================
@app.post("/api/overtime/parse")
async def overtime_parse(attendance: UploadFile = File(...)):
    """근태현황(.xlsx)을 읽어 연장근무 대상일 목록을 반환한다."""
    try:
        name, year, month, records = parse_attendance(await attendance.read())
    except Exception as e:  # noqa: BLE001
        raise HTTPException(400, f"근태현황을 읽지 못했습니다: {e}")
    rows = [{
        "day": r["day"],
        "clock_in": _sec_to_hhmm(r["clock_in"]),
        "clock_out": _sec_to_hhmm(r["clock_out"]),
        "work_start": _sec_to_hhmm(r["clock_in"] + 9 * 3600),  # 근무시작 = 출근+9h
        "work_end": _sec_to_hhmm(r["clock_out"]),              # 근무종료 = 퇴근
        "approved_ot": _sec_to_hhmm(r["approved_ot"]),
        "payoff": "X", "hours": "", "note": "",   # 기본: 대체휴무 미지급
    } for r in records]
    return {"name": name, "year": year, "month": month, "rows": rows}


@app.post("/api/overtime/generate")
async def overtime_generate(
    attendance: UploadFile = File(...),
    payload: str = Form(...),
    template: UploadFile = File(None),
):
    """연장근무신청서(.xlsx)를 채워 다운로드로 반환한다.
    payload(JSON): {extras:{day:{payoff,hours,note}}, dept_position:str}"""
    data = json.loads(payload)
    extras_in = data.get("extras", {})
    dept_position = data.get("dept_position") or None
    extras = {int(k): v for k, v in extras_in.items()}

    tpl = _template_bytes(await template.read() if template else None,
                          DEFAULT_OVERTIME_TPL)
    if tpl is None:
        raise HTTPException(400, "양식을 찾지 못했습니다.")
    att = await attendance.read()
    try:
        buf, n = fill_overtime(tpl, att, extras=extras,
                               dept_position=dept_position)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"생성 실패: {e}")
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    fname = f"연장근무신청서_작성완료_{stamp}.xlsx"
    return _download(buf.getvalue(), fname, XLSX_MIME, count=n)


# ======================================================================
# 헬퍼 / 정적 파일
# ======================================================================
def _sec_to_hhmm(sec):
    if sec is None:
        return ""
    h, m = divmod(int(sec) // 60, 60)
    return f"{h:02d}:{m:02d}"


def _download(data: bytes, filename: str, mime: str, count=None):
    from urllib.parse import quote
    headers = {
        "Content-Disposition":
            f"attachment; filename*=UTF-8''{quote(filename)}",
    }
    if count is not None:
        headers["X-Record-Count"] = str(count)
        headers["Access-Control-Expose-Headers"] = "X-Record-Count"
    return StreamingResponse(io.BytesIO(data), media_type=mime, headers=headers)


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


app.mount("/", StaticFiles(directory=STATIC_DIR), name="static")
