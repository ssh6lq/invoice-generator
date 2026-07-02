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
import secrets
import traceback
from datetime import datetime

try:
    from dotenv import load_dotenv
    load_dotenv()  # 같은 폴더의 .env에서 AI 연결 설정(주소·모델·토큰)을 읽는다.
except Exception:  # noqa: BLE001
    pass

from fastapi import Depends, FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from fastapi.staticfiles import StaticFiles

from langchain_core.messages import HumanMessage, SystemMessage

from receipt_parser import Receipt, SYSTEM_PROMPT, _image_to_data_url
from excel_filler import (
    fill_workbook, get_dropdown_options, get_support_limits, _to_date,
    validate_claims, sort_for_claim, build_claim_xlsx, get_note_examples,
    MEAL_PURPOSES,
)
from overtime_filler import parse_attendance, fill_overtime
import feedback_store
import submission_store

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXPENSE_TPL = os.path.join(APP_DIR, "비용청구양식.xlsm")
DEFAULT_OVERTIME_TPL = os.path.join(APP_DIR, "초과근무(수당)신청서_양식.xlsx")
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

# 문의/이슈 관리자 페이지 보호(HTTP Basic) — .env의 ADMIN_PASSWORD로만 설정한다.
ADMIN_USER = os.getenv("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "")
_admin_security = HTTPBasic()


def require_admin(credentials: HTTPBasicCredentials = Depends(_admin_security)):
    """관리자 페이지/API용 인증. 타이밍 공격 방지를 위해 compare_digest로 비교한다."""
    if not ADMIN_PASSWORD:
        raise HTTPException(500, "관리자 비밀번호가 설정되지 않았습니다. 서버 .env에 ADMIN_PASSWORD를 설정하세요.")
    user_ok = secrets.compare_digest(credentials.username, ADMIN_USER)
    pass_ok = secrets.compare_digest(credentials.password, ADMIN_PASSWORD)
    if not (user_ok and pass_ok):
        raise HTTPException(401, "인증에 실패했습니다.", headers={"WWW-Authenticate": "Basic"})
    return True


app = FastAPI(title="청구서 자동 작성")
feedback_store.init_db()
submission_store.init_db()


@app.middleware("http")
async def no_cache_static(request, call_next):
    """정적 파일(html/js/css)을 브라우저가 캐시해 옛 버전을 쓰는 문제 방지."""
    resp = await call_next(request)
    path = request.url.path
    if path == "/" or path.endswith((".html", ".js", ".css", ".ico")):
        resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


# ======================================================================
# 공통 로직 (Streamlit 앱에서 그대로 가져온 순수 함수들)
# ======================================================================
def _supports_custom_temperature(model: str) -> bool:
    """gpt-5 계열·o1/o3/o4 추론형 모델은 temperature 기본값(1)만 허용한다.
    이런 모델엔 temperature를 아예 보내면 안 되므로(400 에러) False를 돌려준다."""
    m = (model or "").lower().removeprefix("openai:")
    return not m.startswith(("gpt-5", "o1", "o3", "o4"))


def _build_receipt_llm(provider, model, api_key, base_url):
    """영수증 파싱용 LLM 생성. provider='로컬 서버'면 OpenAI 호환 엔드포인트에 붙는다."""
    from langchain_openai import ChatOpenAI
    if provider == "로컬 서버":
        if model.startswith("openai:"):
            model = model[len("openai:"):]
        local_kwargs = dict(
            model=model, base_url=base_url, api_key=(api_key or "EMPTY"),
            max_retries=5,
            model_kwargs={"extra_body": {
                "chat_template_kwargs": {"enable_thinking": False},
            }},
        )
        if _supports_custom_temperature(model):
            local_kwargs["temperature"] = 0.1
        llm = ChatOpenAI(**local_kwargs)
        return llm.with_structured_output(Receipt, method="json_schema")
    # OpenAI 경로: 키가 반드시 필요. 없으면 명확히 안내(사내 모델을 쓰려는데 여기로 빠진 경우 방지).
    if not api_key:
        raise ValueError(
            "OpenAI 모델에는 API Key가 필요합니다. 사내 모델을 쓰려면 사이드바에서 "
            "‘사내 기본 모델’을 선택하거나, 서버 .env의 RECEIPT_PROVIDER='로컬 서버'를 확인하세요.")
    kwargs = {"model": model, "api_key": api_key}
    if _supports_custom_temperature(model):
        kwargs["temperature"] = 0.0
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
        return {"purpose": [], "payment": [], "limits": {}, "meal": [],
                "note_examples": {}, "template_name": None, "is_default": True}
    try:
        opts = get_dropdown_options(tpl)
        limits = get_support_limits(tpl)
        note_examples = get_note_examples(tpl)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"양식을 읽지 못했습니다: {e}")
    return {
        "purpose": opts.get("purpose", []),
        "payment": opts.get("payment", []),
        "limits": limits,
        # 참여자 필수(식대) 목적 — 프론트에서 성명 자동 채움에 사용
        "meal": sorted(MEAL_PURPOSES),
        # 목적별 비고작성예시 — 프론트에서 비고칸 placeholder(얕은 글씨)로 사용
        "note_examples": note_examples,
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
    try:
        llm = _build_receipt_llm(
            provider or ENV_PROVIDER, model or ENV_MODEL,
            api_key or ENV_API_KEY, base_url or ENV_BASE_URL,
        )
        results = _parse_receipts(payload, llm)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()   # 서버 터미널에 전체 원인 출력
        raise HTTPException(500, f"분석 실패: {type(e).__name__}: {e}")
    # 모든 이미지가 같은 이유로 실패하면(예: 모델·토큰·주소 오류) 명확히 알려준다.
    errs = [r["error"] for r in results if r["error"]]
    if results and len(errs) == len(results):
        raise HTTPException(502, f"AI 모델 분석 실패: {errs[0]}")
    # 청구금액 기본값 = 영수금액
    rows = [{
        "date": r["date"], "store": r["store"], "purpose": "",
        "amount": r["amount"], "claim": r["amount"], "payment": "",
        "time": r["time"], "region": r["region"], "note": "",
        "filename": r["filename"], "error": r["error"],
    } for r in results]
    return {"rows": rows}


def _prepare_expense(data, tpl):
    """payload 를 파싱해 (rows, basic, records, norm, limits) 로 정규화한다.
    records: 작성시트 기입용(키 claim_amount).  norm: 검증·비용청구서용(키 claim).
    청구금액은 복지비면 배분, 아니면 목적 한도로 캡한 값을 양쪽에 동일 적용한다.
    """
    rows = [r for r in data.get("rows", []) if (r.get("store") or r.get("date"))]
    basic = data.get("basic", {})
    welfare_budget = _to_int(data.get("welfare_budget")) or 0
    try:
        limits = get_support_limits(tpl)
    except Exception:  # noqa: BLE001
        limits = {}
    alloc = _welfare_alloc(rows, welfare_budget) if welfare_budget else None
    records, norm = [], []
    for i, r in enumerate(rows):
        claim = (alloc[i] if welfare_budget
                 else _capped_claim(r.get("purpose"), r.get("claim"),
                                    r.get("amount"), limits))
        participants = (r.get("participants") or "").strip()
        records.append({
            "date": r.get("date"), "store": r.get("store"),
            "purpose": r.get("purpose"), "amount": _to_int(r.get("amount")),
            "payment": r.get("payment"), "time": r.get("time"),
            "claim_amount": claim, "region": r.get("region"),
            "participants": participants, "note": r.get("note"),
        })
        norm.append({
            "date": r.get("date"), "store": r.get("store"),
            "purpose": r.get("purpose"), "amount": _to_int(r.get("amount")),
            "payment": r.get("payment"), "claim": claim,
            "participants": participants, "time": r.get("time"),
            "region": r.get("region"), "note": r.get("note"),
        })
    return rows, basic, records, norm, limits


@app.post("/api/expense/review")
async def expense_review(
    payload: str = Form(...),
    template: UploadFile = File(None),
):
    """'매크로검토' 대체 — 검증 결과와 정렬·매핑된 미리보기를 반환(파일 생성 없음).
    반환: {ok, issues:[{row,field,code,message}], preview:[...], count, over_limit, year}
    """
    data = json.loads(payload)
    tpl = _template_bytes(await template.read() if template else None,
                          DEFAULT_EXPENSE_TPL)
    if tpl is None:
        raise HTTPException(400, "양식을 찾지 못했습니다.")
    rows, basic, records, norm, limits = _prepare_expense(data, tpl)
    if not rows:
        raise HTTPException(400, "검토할 데이터가 없습니다.")

    year = datetime.now().year
    issues = validate_claims(norm, limits, year=year)
    ordered = sort_for_claim(norm)
    preview = [{
        "date": (_to_date(r["date"]).isoformat() if _to_date(r["date"])
                 else (r["date"] or "")),
        "store": r["store"] or "", "purpose": r["purpose"] or "",
        "participants": r["participants"] or "", "note": r["note"] or "",
        "payment": r["payment"] or "", "amount": r["amount"], "claim": r["claim"],
        "time": r["time"] or "", "region": r["region"] or "",
    } for r in ordered]
    return {"ok": (not issues) and len(norm) <= 25, "issues": issues,
            "preview": preview, "count": len(norm),
            "over_limit": len(norm) > 25, "year": year}


@app.post("/api/expense/generate")
async def expense_generate(
    payload: str = Form(...),
    template: UploadFile = File(None),
):
    """검토를 통과한 데이터로 비용청구서를 생성해 다운로드로 반환한다. 형식 선택:
      fmt="xlsx"(기본) — 매크로 '비용청구서생성'처럼 완성된 비용청구서(+교통비상세) 독립 .xlsx
      fmt="xlsm"       — 작성시트를 채운 원본 양식 .xlsm(매크로 포함, Excel에서 직접 매크로 실행)
    payload(JSON): {rows:[...], basic:{dept,name,title,card}, welfare_budget:int, fmt:str}
    """
    data = json.loads(payload)
    fmt = (data.get("fmt") or "xlsx").lower()
    tpl = _template_bytes(await template.read() if template else None,
                          DEFAULT_EXPENSE_TPL)
    if tpl is None:
        raise HTTPException(400, "양식을 찾지 못했습니다.")
    rows, basic, records, norm, limits = _prepare_expense(data, tpl)
    if not rows:
        raise HTTPException(400, "채울 데이터가 없습니다.")

    # 서버측 재검증(방어) — 통과해야만 생성
    issues = validate_claims(norm, limits, year=datetime.now().year)
    if issues:
        raise HTTPException(400, detail={"message": "검토를 통과하지 못했습니다.",
                                         "issues": issues})
    basic_info = {"dept": basic.get("dept"), "name": basic.get("name"),
                  "title": basic.get("title"), "card": basic.get("card")}
    # 파일명: 소속법인명_{청구항목}청구서_이름_작성일 (예: 넥스노우_비용청구서_남소희_20260701)
    stamp = datetime.now().strftime("%Y%m%d")
    who = (basic.get("name") or "").strip() or "작성완료"
    ttl = str(basic.get("title") or "비용").strip()
    company = str(basic.get("company") or "").strip()
    base = "_".join(x for x in [company, f"{ttl}청구서", who, stamp] if x)
    try:
        if fmt == "xlsm":
            # 작성시트를 채운 원본 양식(.xlsm) — 사용자가 Excel에서 매크로를 직접 실행하는 버전
            buf, _start, n = fill_workbook(tpl, records, append=False,
                                           basic_info=basic_info)
            return _download(buf.getvalue(), f"{base}.xlsm", XLSM_MIME, count=n)
        buf, n = build_claim_xlsx(tpl, sort_for_claim(norm), basic_info=basic_info)
        fname = f"{base}.xlsx"
        # '제출용' 다운로드 = 실제 제출 행위로 간주해 경영지원팀 관리자 페이지에 기록한다.
        total_claim = sum(int(r["claim"] or 0) for r in norm)
        submission_store.add_submission(
            dept=basic.get("dept"), name=basic.get("name"), title=basic.get("title"),
            count=n, total_claim=total_claim, filename=fname,
        )
        return _download(buf.getvalue(), fname, XLSX_MIME, count=n)
    except ValueError as e:  # 25건 초과 등
        raise HTTPException(400, str(e))
    except Exception as e:  # noqa: BLE001
        raise HTTPException(500, f"생성 실패: {e}")


# ======================================================================
# 초과근무 API
# ======================================================================
@app.post("/api/overtime/parse")
async def overtime_parse(attendance: UploadFile = File(...)):
    """근태현황(.xlsx)을 읽어 초과근무 대상일 목록을 반환한다."""
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
        "exclude": "", "exclude_reason": "",      # 제외할 시간 / 제외 사유
    } for r in records]
    return {"name": name, "year": year, "month": month, "rows": rows}


@app.post("/api/overtime/generate")
async def overtime_generate(
    attendance: UploadFile = File(...),
    payload: str = Form(...),
    template: UploadFile = File(None),
):
    """초과근무신청서(.xlsx)를 채워 다운로드로 반환한다.
    payload(JSON): {extras:{day:{payoff,hours,note}}, dept_position:str}"""
    data = json.loads(payload)
    extras_in = data.get("extras", {})
    dept_position = data.get("dept_position") or None
    company = str(data.get("company") or "").strip()
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
    stamp = datetime.now().strftime("%Y%m%d")
    who = (parse_attendance(att)[0] or "").strip() or "작성완료"
    # 파일명: 소속법인명_초과근무신청서_이름_작성일
    fname = "_".join(x for x in [company, "초과근무신청서", who, stamp] if x) + ".xlsx"
    return _download(buf.getvalue(), fname, XLSX_MIME, count=n)


# ======================================================================
# 문의/이슈 API
# ======================================================================
@app.post("/api/feedback")
async def submit_feedback(
    title: str = Form(...),
    content: str = Form(...),
    screenshot: UploadFile = File(None),
):
    title, content = title.strip(), content.strip()
    if not title or not content:
        raise HTTPException(400, "제목과 내용을 입력해 주세요.")
    shot_bytes = await screenshot.read() if screenshot else None
    shot_mime = screenshot.content_type if screenshot else None
    fid = feedback_store.add_feedback(title, content, shot_bytes, shot_mime)
    return {"ok": True, "id": fid}


@app.get("/api/feedback")
def list_feedback_api(_: bool = Depends(require_admin)):
    return {"items": feedback_store.list_feedback()}


@app.get("/api/feedback/{fid}/screenshot")
def feedback_screenshot(fid: int, _: bool = Depends(require_admin)):
    data, mime = feedback_store.get_screenshot(fid)
    if data is None:
        raise HTTPException(404, "첨부된 스크린샷이 없습니다.")
    return StreamingResponse(io.BytesIO(data), media_type=mime or "image/png")


@app.post("/api/feedback/{fid}/status")
async def update_feedback_status(fid: int, status: str = Form(...),
                                 _: bool = Depends(require_admin)):
    if status not in ("open", "done"):
        raise HTTPException(400, "status는 open/done 이어야 합니다.")
    feedback_store.set_status(fid, status)
    return {"ok": True}


@app.delete("/api/feedback/{fid}")
def delete_feedback_api(fid: int, _: bool = Depends(require_admin)):
    feedback_store.delete_feedback(fid)
    return {"ok": True}


# ======================================================================
# 청구서 제출내역 API (경영지원팀 전용)
# ======================================================================
@app.get("/api/submissions")
def list_submissions_api(_: bool = Depends(require_admin)):
    return {"items": submission_store.list_submissions()}


@app.post("/api/submissions/{sid}/status")
async def update_submission_status(sid: int, status: str = Form(...),
                                   note: str = Form(""),
                                   _: bool = Depends(require_admin)):
    if status not in submission_store.STATUSES:
        raise HTTPException(400, f"status는 {submission_store.STATUSES} 중 하나여야 합니다.")
    submission_store.set_status(sid, status, note.strip())
    return {"ok": True}


@app.delete("/api/submissions/{sid}")
def delete_submission_api(sid: int, _: bool = Depends(require_admin)):
    submission_store.delete_submission(sid)
    return {"ok": True}


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


@app.get("/admin_feedback.html")
def admin_feedback_page(_: bool = Depends(require_admin)):
    return FileResponse(os.path.join(STATIC_DIR, "admin_feedback.html"))


@app.get("/admin_submissions.html")
def admin_submissions_page(_: bool = Depends(require_admin)):
    return FileResponse(os.path.join(STATIC_DIR, "admin_submissions.html"))


app.mount("/", StaticFiles(directory=STATIC_DIR), name="static")
