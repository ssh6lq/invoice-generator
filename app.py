"""
app.py — 청구서 자동 작성 (Streamlit)

두 가지 모드를 선택해서 사용한다.
  1) 🧾 비용청구  : 영수증 이미지를 GPT 비전으로 분석해 비용청구 양식(.xlsm) 작성
  2) 🌙 연장근무(야근) 청구 : 월간 근태현황(.xlsx)을 읽어 연장근무(수당)신청서 양식 작성

실행:  streamlit run app.py
"""

import os
import html
from datetime import datetime

import pandas as pd
import streamlit as st

from langchain_core.messages import HumanMessage, SystemMessage

from receipt_parser import Receipt, SYSTEM_PROMPT, _image_to_data_url
from excel_filler import fill_workbook, get_dropdown_options, get_support_limits, _to_date
from overtime_filler import parse_attendance, fill_overtime

st.set_page_config(page_title="청구서 자동 작성", page_icon="🧾", layout="wide")

# 번들 기본 양식 (사용자가 따로 업로드하지 않으면 이 파일을 사용)
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXPENSE_TPL = os.path.join(APP_DIR, "비용청구양식.xlsm")
DEFAULT_OVERTIME_TPL = os.path.join(APP_DIR, "연장근무(수당)신청서_양식.xlsx")


def _build_receipt_llm(provider, model, api_key, base_url):
    """영수증 파싱용 LLM 생성. provider='로컬 서버'면 OpenAI 호환 엔드포인트(base_url)에
    붙고 API 키가 없어도 된다. 어느 쪽이든 비전(이미지 입력) 모델이어야 한다."""
    from langchain_openai import ChatOpenAI
    if provider == "로컬 서버":
        # init_chat_model 식 접두어('openai:')를 붙여 넣어도 자동 제거 (실제 API엔 모델명만)
        if model.startswith("openai:"):
            model = model[len("openai:"):]
        llm = ChatOpenAI(
            model=model,
            base_url=base_url,
            api_key=(api_key or "EMPTY"),   # 키가 필요 없는 서버용 더미값
            temperature=0.1,
            max_retries=5,
            model_kwargs={"extra_body": {
                # Qwen3 계열: 추론(thinking) 토큰을 끄지 않으면 JSON 출력이 깨질 수 있음
                "chat_template_kwargs": {"enable_thinking": False},
            }},
        )
        # vLLM은 tool-calling 파서가 꺼져 있을 수 있어 guided JSON(json_schema)이 더 안전
        return llm.with_structured_output(Receipt, method="json_schema")
    kwargs = {"model": model, "temperature": 0.0}
    if api_key:
        kwargs["api_key"] = api_key
    llm = ChatOpenAI(**kwargs)
    return llm.with_structured_output(Receipt)


def _parse_receipts(payload, llm, on_progress=None):
    """payload=list[(filename, bytes)] -> list[dict]. 주어진 llm으로 한 장씩 파싱한다.
    메시지는 [이미지 → 텍스트] 순서로 담는다(로컬 서버 요구 형식)."""
    results = []
    for idx, (fname, content) in enumerate(payload):
        rec = {"filename": fname, "date": None, "store": None,
               "amount": None, "time": None, "region": None, "error": None}
        try:
            data_url = _image_to_data_url(content, fname)
            messages = [
                SystemMessage(content=SYSTEM_PROMPT),
                HumanMessage(content=[
                    {"type": "image_url", "image_url": {"url": data_url}},
                    {"type": "text", "text": "이 영수증에서 정보를 추출해줘."},
                ]),
            ]
            r = llm.invoke(messages)
            rec.update(date=r.date, store=r.store, amount=r.amount,
                       time=r.time, region=r.region)
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        results.append(rec)
        if on_progress:
            on_progress(idx + 1, len(payload), rec)
    return results


def _template_bytes_name(uploaded, default_path):
    """업로드 파일이 있으면 그것을, 없으면 번들 기본 양식을 (bytes, name, is_default)로 반환.
    둘 다 없으면 (None, None, False)."""
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name, False
    if os.path.exists(default_path):
        with open(default_path, "rb") as f:
            return f.read(), os.path.basename(default_path), True
    return None, None, False


# =========================================================================
# 테마 / UI 헬퍼
# =========================================================================
THEMES = {
    "expense": {
        "tag": "비용청구",
        "headline": "영수증 사진 한 장으로\n비용청구서를 완성하세요",
        "desc": "영수증 사진을 올리면 날짜·상호명·금액을 인식해 비용청구 양식을 자동으로 채웁니다.",
    },
    "overtime": {
        "tag": "연장근무청구",
        "headline": "근태현황 파일로\n연장근무신청서를 자동 작성하세요",
        "desc": "월간 근태현황을 올리면 '승인 초과 근로시간'이 있는 날을 찾아 신청서를 자동으로 채웁니다.",
    },
}


def inject_css(_t):
    st.markdown(
        """
        <style>
        .stApp { background: #f9f9f8; }
        .block-container { padding-top: 1.8rem; max-width: 1100px; }

        /* 히어로 영역 */
        .ed-hero {
            border-top: 2px solid #1a1a1a;
            padding-top: 16px; padding-bottom: 16px;
            border-bottom: 1px solid #e5e5e3; margin-bottom: 20px;
        }
        .ed-hero .meta { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
        .ed-hero .appname { font-size: 12px; font-weight: 600; color: #1a1a1a; }
        .ed-hero .vdiv { display: inline-block; width: 1px; height: 11px;
            background: #ccc; vertical-align: middle; }
        .ed-hero .sub { font-size: 12px; color: #999; }
        .ed-hero .bottom { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
        .ed-hero .headline { font-size: 22px; font-weight: 600;
            color: #1a1a1a; line-height: 1.35; white-space: pre-line; }
        .ed-hero .tag { font-size: 11px; font-weight: 500; color: #666;
            border: 1px solid #d0d0cc; padding: 4px 12px; border-radius: 20px;
            white-space: nowrap; margin-top: 4px; }

        /* 단계 헤더 */
        .step { display: flex; align-items: baseline; gap: 10px; margin: .3rem 0 .15rem; }
        .step .num { font-size: 11px; font-weight: 500; color: #aaa; min-width: 20px; }
        .step .ttl { font-size: 15px; font-weight: 600; color: #1a1a1a; }

        /* 버튼 */
        .stButton > button, .stDownloadButton > button {
            border-radius: 6px; font-weight: 600; padding: .45rem 1rem;
        }
        .stButton > button[kind="primary"], .stDownloadButton > button {
            background: #1a1a1a; border: none; color: #fff;
        }
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
            background: #333;
        }

        /* 지표 카드 */
        div[data-testid="stMetric"] {
            background: #fff; border: 1px solid #e5e5e3; border-radius: 6px;
            padding: 12px 16px;
        }
        div[data-testid="stMetricValue"] { font-size: 1.3rem; }

        /* 섹션 컨테이너(카드) */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: #fff; border-radius: 6px; border-color: #e5e5e3 !important;
        }

        /* 입력칸 라벨 */
        div[data-testid="stWidgetLabel"] label,
        div[data-testid="stWidgetLabel"] p {
            color: #333 !important; font-weight: 600;
        }

        /* 입력칸 — 연보라(테마 secondaryBackground) 제거: 바깥 래퍼·안쪽 input 모두 흰색 */
        .stTextInput div[data-baseweb="input"],
        .stNumberInput div[data-baseweb="input"],
        .stTextInput div[data-baseweb="base-input"],
        .stNumberInput div[data-baseweb="base-input"],
        .stTextInput input, .stNumberInput input {
            background: #fff !important;
            background-color: #fff !important;
        }
        .stTextInput div[data-baseweb="input"],
        .stNumberInput div[data-baseweb="input"] {
            border: 1px solid #d5d5d2 !important; border-radius: 6px !important;
        }
        .stTextInput div[data-baseweb="input"]:focus-within,
        .stNumberInput div[data-baseweb="input"]:focus-within {
            border-color: #1a1a1a !important;
            box-shadow: 0 0 0 2px rgba(26,26,26,.1) !important;
        }
        .stTextInput input, .stNumberInput input { color: #1a1a1a !important; font-weight: 500; }
        .stTextInput input::placeholder, .stNumberInput input::placeholder { color: #aaa !important; }

        /* 선택박스도 같은 흰색 톤으로 통일 */
        .stSelectbox div[data-baseweb="select"] > div {
            background: #fff !important; border: 1px solid #d5d5d2 !important; border-radius: 6px !important;
        }
        .stSelectbox div[data-baseweb="select"] > div:focus-within {
            border-color: #1a1a1a !important; box-shadow: 0 0 0 2px rgba(26,26,26,.1) !important;
        }
        /* 일괄 툴바: 목적 셀렉트와 결제방식 세그먼트(칩) 높이 통일 (둘 다 34px) */
        .stSelectbox div[data-baseweb="select"] > div {
            min-height: 34px !important; height: 34px !important;
        }
        div[data-testid="stButtonGroup"] button[data-testid^="stBaseButton-segmented_control"] {
            min-height: 34px !important; height: 34px !important;
            padding-top: 0 !important; padding-bottom: 0 !important;
        }
        /* 일괄 툴바 '일괄 채우기' 라벨 — 마크다운 기본 여백 제거 + 세로 중앙 */
        .st-key-bulkbar [data-testid="stMarkdownContainer"],
        .st-key-bulkbar [data-testid="stMarkdownContainer"] p { margin: 0 !important; }

        /* 파일 업로드 드롭존 — 연보라(테마색) 제거 + 정돈된 점선 박스 */
        [data-testid="stFileUploaderDropzone"] {
            background: #fff !important;
            border: 1px dashed #cfcfca !important; border-radius: 8px !important;
            padding: 14px 18px !important;
            transition: border-color .15s ease, background .15s ease;
        }
        [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #1a1a1a !important; background: #fafafa !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] { color: #666 !important; }
        [data-testid="stFileUploaderDropzoneInstructions"] svg,
        [data-testid="stFileUploaderDropzoneInstructions"] span[data-testid="stIconMaterial"] {
            color: #aaa !important; fill: #aaa !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] small { color: #aaa !important; }
        /* Browse files 버튼 — 흰 바탕 + 회색 테두리 (에디토리얼 톤) */
        [data-testid="stFileUploaderDropzone"] button {
            background: #fff !important; color: #1a1a1a !important;
            border: 1px solid #d5d5d2 !important; border-radius: 6px !important;
            font-weight: 600 !important; min-height: 36px !important;
        }
        [data-testid="stFileUploaderDropzone"] button:hover {
            border-color: #1a1a1a !important; background: #f5f5f4 !important; color: #1a1a1a !important;
        }

        /* 영수증 업로더: 네이티브 칩 목록만 숨기고(커스텀 목록만 노출),
           기본 업로드 안내문(아이콘·Browse·용량)은 그대로 유지한다. */
        .st-key-receipt_box [data-testid="stFileChips"] { display: none !important; }
        .st-key-receipt_box [data-testid="stFileUploaderFile"] { display: none !important; }
        /* 파일이 올라가 있어도 기본 안내문이 사라지지 않도록 강제로 표시 */
        .st-key-receipt_box [data-testid="stFileUploaderDropzoneInstructions"] {
            display: flex !important;
        }

        /* 업로드 파일 목록 */
        .file-lines { margin: .3rem 0 1rem; display: flex; flex-direction: column; gap: 4px; }
        .file-line { display: flex; align-items: center; gap: .5rem;
            font-size: .84rem; color: #666; }
        .file-line .num { flex: none; font-size: 11px; font-weight: 500; color: #aaa; min-width: 18px; }
        .file-line .nm { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        /* ✕(파일 제거) 버튼만 투명 처리 — Browse files 버튼은 건드리지 않도록 stButton으로 한정 */
        .st-key-receipt_box [data-testid="stButton"] button,
        .st-key-receipt_box [data-testid="stButton"] button:hover,
        .st-key-receipt_box [data-testid="stButton"] button:active,
        .st-key-receipt_box [data-testid="stButton"] button:focus {
            background: transparent !important; background-color: transparent !important;
            border: none !important; box-shadow: none !important; outline: none !important;
            color: #999 !important; padding: 0 !important;
            min-height: 0 !important; height: auto !important;
            font-size: .9rem !important; line-height: 1 !important;
        }
        .st-key-receipt_box [data-testid="stButton"] button:hover { color: #c0392b !important; }
        .st-key-receipt_box [data-testid="stButton"] {
            display: flex; align-items: center; min-height: 22px; background: transparent !important;
        }

        /* 사이드바 */
        section[data-testid="stSidebar"] { background: #fff; border-right: 1px solid #e5e5e3; }
        .side-brand {
            font-size: 14px; font-weight: 600; color: #1a1a1a;
            display: flex; align-items: center; gap: .4rem; margin-bottom: .2rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(t):
    st.markdown(
        '<div class="ed-hero">'
        '<div class="meta">'
        '<span class="appname">청구서 자동 작성</span>'
        '<span class="vdiv"></span>'
        '<span class="sub">영수증 인식 · 양식 자동완성</span>'
        '</div>'
        '<div class="bottom">'
        f'<div class="headline">{t["headline"]}</div>'
        f'<span class="tag">{t["tag"]}</span>'
        '</div>'
        '</div>',
        unsafe_allow_html=True,
    )


def step(n, title, desc=None):
    num_str = f"{n:02d}" if isinstance(n, int) else str(n)
    st.markdown(
        f'<div class="step"><span class="num">{num_str}</span>'
        f'<span class="ttl">{title}</span></div>',
        unsafe_allow_html=True,
    )
    if desc:
        st.caption(desc)


def _img_key(f):
    """업로드 파일의 중복 판별 키 = 파일명. 같은 파일명은 중복으로 본다."""
    return f.name


def _capped_claim(purpose, claim_amount, receipt_amount, limits):
    """목적별 지원한도를 적용한 청구금액을 반환.
    한도가 없거나 금액이 비어 있으면 원래 값을 그대로 돌려준다."""
    base = claim_amount if pd.notna(claim_amount) and claim_amount != "" else receipt_amount
    if not (pd.notna(base) and base != ""):
        return base
    base = int(base)
    lim = limits.get(purpose) if purpose else None
    return base if lim is None else min(base, int(lim))


@st.cache_data(show_spinner=False)
def _cached_dropdowns(tpl_bytes):
    """양식의 드롭다운(목적/결제방식) 목록을 캐시. 양식이 바뀔 때만 다시 파싱한다."""
    return get_dropdown_options(tpl_bytes)


@st.cache_data(show_spinner=False)
def _cached_limits(tpl_bytes):
    """양식의 목적별 지원한도를 캐시. 양식이 바뀔 때만 다시 파싱한다."""
    return get_support_limits(tpl_bytes)


def _sort_by_date(df):
    """검토 표를 영수일자 오름차순으로 정렬(날짜 없는 행은 맨 뒤, 같은 날짜는 기존 순서 유지)."""
    order = sorted(range(len(df)), key=lambda i: (
        _to_date(df.iloc[i]["영수일자"]) is None,
        _to_date(df.iloc[i]["영수일자"]) or datetime.max.date(),
    ))
    return df.iloc[order].reset_index(drop=True)


def _welfare_fill(df, budget):
    """복지비 모드: 영수일자 빠른 순으로 누적해 한도(budget)까지 청구금액을 채운다.
    한도를 넘는 영수증은 남은 한도만큼만, 그 뒤 영수증은 0을 청구한다.
    (날짜 없는 행은 맨 뒤로 미뤄 배분) 반환값은 df 원래 순서에 맞춘 청구금액 리스트."""
    order = sorted(df.index, key=lambda i: (
        _to_date(df.at[i, "영수일자"]) is None,
        _to_date(df.at[i, "영수일자"]) or datetime.max.date(),
    ))
    remaining = int(budget) if budget else 0
    give_by_idx = {}
    for i in order:
        amt = df.at[i, "영수금액"]
        a = int(amt) if pd.notna(amt) and amt != "" else 0
        give = min(a, remaining)
        remaining -= give
        give_by_idx[i] = give
    return [give_by_idx[i] for i in df.index]


def _results_to_df(results):
    """parse_many 결과 리스트 -> 검토용 DataFrame.
    양식 작성시트 열 순서에 맞춰 사용자 입력 열(목적·청구금액·결제방식·지역·참여자·비고)을 포함한다.
    청구금액은 기본값으로 영수금액을 채워두고 필요 시 수정한다."""
    df = pd.DataFrame(results)
    return pd.DataFrame({
        "영수일자": df["date"],
        "거래처명": df["store"],
        "목적": "",
        "영수금액": df["amount"],
        "청구금액": df["amount"],   # 기본 = 영수금액
        "결제방식": "",
        "영수시간": df["time"],
        "지역": df["region"],       # 영수증 주소에서 자동 추출(구)
        "비고": "",                 # 비워둠 (필요 시 직접 입력)
        "파일명": df["filename"],
        "오류": df["error"],
    })


# =========================================================================
# 사이드바 — 작업 선택
# =========================================================================
with st.sidebar:
    st.markdown('<div class="side-brand">청구서 자동 작성</div>', unsafe_allow_html=True)
    st.caption("영수증·근태현황을 올리면 양식을 자동으로 채워드려요.")
    st.divider()
    mode = st.radio(
        "작업 선택",
        ["🧾 비용청구", "🌙 연장근무(야근) 청구"],
        index=0,
    )
    st.divider()

theme = THEMES["expense"] if mode.startswith("🧾") else THEMES["overtime"]
inject_css(theme)


# =========================================================================
# 1) 비용청구 모드
# =========================================================================
def render_expense():
    hero(theme)

    with st.sidebar:
        st.markdown("##### ⚙️ 설정")
        provider = st.radio(
            "모델 제공자", ["로컬 서버", "OpenAI"], index=0, horizontal=True,
            help="로컬 서버(OpenAI 호환, 예: vLLM)는 API 키 없이 쓸 수 있어요. "
                 "어느 쪽이든 비전(이미지 입력) 모델이어야 영수증 인식이 됩니다.",
        )
        if provider == "로컬 서버":
            base_url = st.text_input(
                "서버 주소 (base_url)", value="http://192.168.1.51:8001/v1",
                help="OpenAI 호환 엔드포인트. vLLM은 보통 끝에 /v1 을 붙입니다.",
            )
            model = st.text_input(
                "모델명", value="/models/Qwen3.6-35B-A3B-FP8",
                placeholder="예: /models/Qwen3.6-35B-A3B-FP8",
                help="서버 구동 시 지정한 모델명과 정확히 같아야 합니다. "
                     "'openai:' 접두어는 붙이지 마세요(자동 제거됨).",
            )
            api_key = st.text_input(
                "API Key (선택)", type="password", value="",
                help="키 인증이 필요한 서버만 입력. 없으면 비워두세요.",
            )
        else:
            base_url = None
            api_key = st.text_input(
                "OpenAI API Key",
                type="password",
                value=os.getenv("OPENAI_API_KEY", ""),
                help="sk-... 형식. 입력값은 이 세션에서만 사용됩니다.",
            )
            model = st.selectbox(
                "모델",
                ["gpt-4o", "gpt-4o-mini", "gpt-4.1", "gpt-4.1-mini"],
                index=0,
                help="비전(이미지) 입력을 지원하는 모델이어야 합니다.",
            )
        st.divider()
        st.markdown("##### 📄 양식")
        xlsm_file = st.file_uploader(
            "비용청구 양식(.xlsm)", type=["xlsm", "xlsx"],
            help="비워두면 기본 제공 양식(비용청구양식.xlsm)을 자동으로 사용합니다.",
        )
        tpl_bytes, tpl_name, tpl_is_default = _template_bytes_name(
            xlsm_file, DEFAULT_EXPENSE_TPL)
        if tpl_is_default:
            st.caption(f"✅ 기본 양식 사용 중: **{tpl_name}**")
        elif tpl_bytes is None:
            st.caption("⚠️ 기본 양식을 찾지 못했습니다. 양식을 업로드하세요.")
    # 양식에서 목적/결제방식 드롭다운 목록 + 목적별 지원한도를 추출해 둔다.
    purpose_opts, payment_opts, support_limits = [], [], {}
    if tpl_bytes is not None:
        try:
            opts = _cached_dropdowns(tpl_bytes)
            purpose_opts = opts.get("purpose", [])
            payment_opts = opts.get("payment", [])
            support_limits = _cached_limits(tpl_bytes)
        except Exception as e:  # noqa: BLE001
            st.sidebar.warning(f"드롭다운 목록을 읽지 못했습니다: {e}")

    # ---- 1. 기초정보 입력 ------------------------------------------------
    with st.container(border=True, key="basic_box"):
        step(1, "기초정보 입력",
             "비용청구서에 들어갈 기본 정보예요. 먼저 입력하세요. "
             "성명을 넣으면 본인확인이 '서명완료'로 처리되어 엑셀에선 매크로검토만 누르면 됩니다.")
        fv = st.session_state.get("form_ver", 0)  # 전체 초기화 시 +1 → 입력칸이 새 위젯으로 비워짐

        # 청구항목명 — 한 줄 가로 토글 (텍스트칸과 높이 충돌 없이 위에 단독 배치)
        bi_title = st.radio("청구항목명", ["비용", "복지비"], key=f"bi_title_{fv}",
                            horizontal=True,
                            help="선택값 뒤에 '청구서'가 붙어 제목이 됩니다(예: '비용청구서').")
        welfare_budget = 0
        if bi_title == "복지비":
            wcol, _ = st.columns([1, 2])
            man = wcol.number_input(
                "복지비 한도(만원)", min_value=0, step=10, value=None,
                key=f"welfare_budget_{fv}", format="%d", placeholder="예: 50",
                help="사원별 복지비 예산을 만원 단위로 입력하세요(예: 50 → 50만원). "
                     "영수일자 빠른 순으로 누적해 이 금액까지 청구금액이 자동으로 채워집니다. "
                     "(목적별 한도 보정은 적용 안 함)")
            welfare_budget = int(man) * 10000 if man else 0

        # 작성자 정보 — 소속부서명·성명 2칸 정렬
        ci2, ci3 = st.columns(2, gap="medium")
        bi_dept = ci2.text_input("소속부서명", key=f"bi_dept_{fv}")
        bi_name = ci3.text_input("성명", key=f"bi_name_{fv}")

    # ---- 2. 영수증 업로드 ------------------------------------------------
    with st.container(border=True, key="receipt_box"):
        step(2, "영수증 이미지 업로드",
             "여러 장을 한 번에 올릴 수 있어요. 분석한 뒤 사진을 더 추가해도 됩니다.")
        images = st.file_uploader(
            "영수증 사진",
            type=["jpg", "jpeg", "png", "webp", "bmp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            key=f"uploader_{st.session_state.get('uploader_ver', 0)}",
        )
        # 네이티브 칩 목록은 CSS로 숨기고(아래) 커스텀 목록만 노출한다.
        # 사용자가 ✕로 제거한 파일은 ignored 집합으로 걸러낸다(숨긴 업로더엔 남아 있어도 앱에선 제외).
        ignored = st.session_state.setdefault("ignored_uploads", set())
        if images:
            images = [im for im in images if im.name not in ignored]
        # 같은 파일명이 같은 묶음에 두 번 이상 들어온 경우(중복) 감지 후, 첫 장만 남기고 제거
        names = [im.name for im in (images or [])]
        batch_dups = sorted({n for n in names if names.count(n) > 1})
        already_dups = sorted({im.name for im in (images or [])
                               if im.name in st.session_state.get("analyzed_keys", set())})
        if images:
            seen_names, uniq = set(), []
            for im in images:
                if im.name in seen_names:
                    continue
                seen_names.add(im.name)
                uniq.append(im)
            images = uniq   # 이후 미리보기·분석은 중복 제거된 것만 사용
        if batch_dups:
            st.warning("같은 파일명이 중복으로 올라와 자동으로 첫 번째만 남기고 제외했어요: "
                       + ", ".join(batch_dups))
        if already_dups:
            st.warning("이미 분석한 파일명이라 다시 분석하지 않고 건너뜁니다: "
                       + ", ".join(already_dups))
        if images:
            st.caption(f"🖼️ {len(images)}장 업로드됨 (중복 제외)")
            for i, im in enumerate(images, 1):
                lc, rc = st.columns([0.93, 0.07])
                lc.markdown(
                    f'<div class="file-line"><span class="num">{i}</span>'
                    f'<span class="nm">{html.escape(im.name)}</span></div>',
                    unsafe_allow_html=True)
                if rc.button("✕", key=f"rm_{im.name}", help="이 파일을 목록에서 제거"):
                    ignored.add(im.name)
                    st.rerun()
            # 사진은 켤 때만 표시 — 매 편집마다 다시 그리지 않아 표 입력이 매끄러움.
            if st.toggle("사진 미리보기", value=False, key="show_thumbs",
                         help="켜면 업로드한 영수증 사진을 표시합니다. 표 편집이 느리면 꺼두세요."):
                # 칸 수를 항상 5로 고정 — 1~2장이어도 왼쪽부터 좁은 칸에 나란히 붙도록
                cols = st.columns(5, gap="small")
                for i, img in enumerate(images):
                    with cols[i % len(cols)]:
                        st.image(img.getvalue(), caption=img.name, width=180)

    # ---- 3. 분석 실행 (새로 추가된 사진만) -------------------------------
    with st.container(border=True):
        step(3, "영수증 분석",
             "새로 추가한 영수증만 분석해서 기존 목록에 덧붙입니다. (이미 분석한 사진은 건너뜀)")

        analyzed = st.session_state.setdefault("analyzed_keys", set())
        # 같은 파일명은 자동으로 첫 번째 한 장만 분석 (이미 분석한 이름·묶음 내 중복 모두 제외)
        pending, seen = [], set()
        for im in (images or []):
            if im.name in analyzed or im.name in seen:
                continue
            seen.add(im.name)
            pending.append(im)

        b1, b2 = st.columns([3, 1])
        with b1:
            run = st.button(
                f"🔍 새 영수증 분석 ({len(pending)}장)",
                type="primary", disabled=not pending, width='stretch',
            )
        with b2:
            if st.button("🗑️ 목록 초기화", disabled="df" not in st.session_state,
                         width='stretch',
                         help="누적된 분석 결과를 모두 지우고 처음부터 다시 시작합니다."):
                st.session_state.pop("df", None)
                st.session_state["analyzed_keys"] = set()
                st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1
                for k in ("gen_buf", "gen_name", "gen_msg", "gen_preview", "edited_snapshot"):
                    st.session_state.pop(k, None)
                st.rerun()

        if images and not pending and run is False:
            st.info("새로 추가된 영수증이 없습니다. (올라온 사진은 모두 분석됨)")
        elif not images:
            st.info("먼저 위에서 영수증 사진을 올려주세요.")

        if run:
            err = None
            if provider == "OpenAI" and not api_key:
                err = "사이드바에 OpenAI API Key를 입력하세요."
            elif provider == "로컬 서버" and not base_url:
                err = "사이드바에 로컬 서버 주소(base_url)를 입력하세요."
            elif provider == "로컬 서버" and not model:
                err = "사이드바에 로컬 서버의 모델명을 입력하세요."
            if err:
                st.error(err)
            else:
                payload = [(im.name, im.getvalue()) for im in pending]
                progress = st.progress(0.0, text="분석 준비 중...")

                def _cb(done, total, rec):
                    label = rec.get("store") or rec["filename"]
                    progress.progress(done / total, text=f"분석 중 {done}/{total} — {label}")

                spin = "로컬 모델로" if provider == "로컬 서버" else "GPT로"
                with st.spinner(f"{spin} 새 영수증을 분석하는 중..."):
                    llm = _build_receipt_llm(provider, model, api_key, base_url)
                    results = _parse_receipts(payload, llm, on_progress=_cb)
                progress.empty()

                new_df = _results_to_df(results)
                # 지금까지의 편집 내용(편집표 스냅샷)을 보존한 채 새 행을 덧붙이고 날짜순 정렬
                base = st.session_state.get("edited_snapshot")
                if base is None:
                    base = st.session_state.get("df")
                combined = (pd.concat([base, new_df], ignore_index=True)
                            if base is not None else new_df)
                st.session_state["df"] = _sort_by_date(combined)
                st.session_state.pop("edited_snapshot", None)
                # 새 행이 추가됐으니 편집 표를 새 데이터 기준으로 다시 맞춤
                st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1
                for im in pending:
                    analyzed.add(_img_key(im))

                errs = [r for r in results if r["error"]]
                if errs:
                    st.warning(f"{len(errs)}건 분석 실패. 표의 '오류' 열을 확인하세요.")
                st.success(f"새 영수증 {len(results)}건 분석 완료 "
                           f"(누적 {len(st.session_state['df'])}건). 아래 표에서 검토·수정하세요.")

    # ---- 4. 검토/수정 표 -------------------------------------------------
    if "df" in st.session_state:
        with st.container(border=True):
            welfare = (bi_title == "복지비")
            if welfare:
                desc4 = ("복지비 청구는 한도까지 영수증 순서대로 청구금액이 자동으로 채워집니다. "
                         "셀을 더블클릭해 수정하세요.")
            else:
                desc4 = ("비용 청구는 목적별 한도 초과분이 생성 시 차감됩니다. "
                         "셀을 더블클릭해 수정하세요.")
            step(4, "검토 및 수정", desc4)
            if not purpose_opts:
                st.info("목적·결제방식 선택지를 채우려면 사이드바에서 비용청구 양식(.xlsm)을 업로드하세요.")

            # 목적·결제방식을 모든 행에 한 번에 채우기 — 고른 값(여러 열)을 한 번에 적용
            def _apply_to_all(updates):
                """updates: {열이름: 값} — 모든 행의 해당 열을 값으로 채운다."""
                base = st.session_state.get("edited_snapshot")
                if base is None or len(base) != len(st.session_state["df"]):
                    base = st.session_state["df"]
                df2 = base.copy()
                for col, value in updates.items():
                    df2[col] = value
                st.session_state["df"] = df2.reset_index(drop=True)
                st.session_state.pop("edited_snapshot", None)
                st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1

            # 표 바로 위 인라인 툴바 — 선택/클릭 즉시 모든 행에 적용(적용 버튼 없음)
            # [일괄] 목적[▾] 결제방식[세그먼트]  · · · (여백)
            def _bulk_apply_purpose():
                v = st.session_state.get("bulk_purpose")
                if v:
                    _apply_to_all({"목적": v})

            def _bulk_apply_pay():
                v = st.session_state.get("bulk_pay")
                if v:
                    _apply_to_all({"결제방식": v})

            if purpose_opts or payment_opts:
                with st.container(key="bulkbar"):
                    tb = st.columns([0.6, 1.2, 3.0, 1.2], vertical_alignment="center")
                    tb[0].markdown(
                        "<div style='display:flex;align-items:center;height:34px;"
                        "font-size:12px;font-weight:600;color:#555;'>일괄 채우기</div>",
                        unsafe_allow_html=True)
                    if purpose_opts:
                        tb[1].selectbox(
                            "목적 일괄", purpose_opts, key="bulk_purpose",
                            index=None, placeholder="목적 선택", label_visibility="collapsed",
                            on_change=_bulk_apply_purpose,
                            help="고르면 모든 행의 목적에 바로 채워져요. 이후 표에서 다른 행만 고치면 됩니다.")
                    if payment_opts:
                        tb[2].segmented_control(
                            "결제방식 일괄", payment_opts, key="bulk_pay",
                            selection_mode="single", label_visibility="collapsed",
                            on_change=_bulk_apply_pay,
                            help="클릭하면 모든 행의 결제방식에 바로 채워져요.")

            # 복지비 모드에선 청구금액을 표에서 숨긴다(자동 배분이라 수정 불가).
            # 표 아래에 별도 미리보기로 보여줘, 편집 표 입력을 건드리지 않으므로 선택이 풀리지 않는다.
            claim_cfg = None if welfare else st.column_config.NumberColumn(
                "청구금액", format="%d",
                help="실제 청구할 금액. 기본은 영수금액이며 부분 청구 시 수정하세요.")

            # 원본 df는 그대로 두고(수정값은 위젯이 자체 보관) 새 분석 추가 때만 합친다.
            # → 매 rerun마다 원본을 되써넣지 않으므로 목적·결제방식 선택이 풀리지 않는다.
            edited = st.data_editor(
                st.session_state["df"],
                width='stretch',
                num_rows="dynamic",
                key=f"review_editor_{st.session_state.get('editor_ver', 0)}",
                column_config={
                    "목적": st.column_config.SelectboxColumn(
                        "목적", options=purpose_opts, required=False,
                        help="양식의 목적 목록에서 선택"),
                    "결제방식": st.column_config.SelectboxColumn(
                        "결제방식", options=payment_opts, required=False,
                        help="양식의 결제방식 목록에서 선택"),
                    "영수금액": st.column_config.NumberColumn("영수금액", format="%d"),
                    "청구금액": claim_cfg,
                    "지역": st.column_config.TextColumn(
                        "지역", help="영수증 주소에서 자동 추출(구). 필요 시 수정하세요."),
                    "비고": st.column_config.TextColumn("비고", help="기본은 비워둠"),
                    "오류": st.column_config.TextColumn(disabled=True),
                },
            )

            # 현재 편집 결과를 스냅샷으로만 보관(새 분석 추가·엑셀 생성 시 사용).
            st.session_state["edited_snapshot"] = edited

            if welfare:
                alloc = _welfare_fill(edited, welfare_budget)
                total_receipt = sum(int(a) for a in edited["영수금액"]
                                    if pd.notna(a) and a != "")
                total_claim = sum(alloc)
                note = "  — 한도 초과분은 청구에서 제외됨" if total_receipt > welfare_budget else ""
                st.caption(f"💼 복지비 한도 {int(welfare_budget):,}원 · "
                           f"영수 합계 {total_receipt:,}원 · 청구(자동) {total_claim:,}원{note}")
                preview = edited[["영수일자", "거래처명", "영수금액"]].copy()
                preview["복지비 청구금액"] = alloc
                st.caption("아래는 영수일자 순으로 한도까지 자동 배분된 청구금액 미리보기예요.")
                st.dataframe(preview, width='stretch', hide_index=True)
            else:
                # 목적별 한도 보정은 '생성' 시 적용 — 초과 건수만 안내
                over = 0
                for idx in edited.index:
                    cur = edited.at[idx, "청구금액"]
                    capped_val = _capped_claim(edited.at[idx, "목적"], cur,
                                               edited.at[idx, "영수금액"], support_limits)
                    if pd.notna(cur) and cur != "" and capped_val != int(cur):
                        over += 1
                if over:
                    st.caption(f"⚠️ 한도 초과 {over}건은 '비용청구서 생성' 시 "
                               "목적별 한도로 자동 보정됩니다.")

        # ---- 5. 엑셀 생성 ------------------------------------------------
        with st.container(border=True):
            step(5, "비용청구서 생성")

            def _do_generate(append):
                """편집된 표로 비용청구서를 채워 결과를 세션에 저장한다.
                append=False 면 첫 줄부터 새로, True 면 기존 내용 뒤에 이어서 작성."""
                alloc = _welfare_fill(edited, welfare_budget) if welfare else None
                records = []
                for i, (_, row) in enumerate(edited.iterrows()):
                    if not row.get("거래처명") and not row.get("영수일자"):
                        continue
                    claim = (alloc[i] if welfare
                             else _capped_claim(row.get("목적"), row.get("청구금액"),
                                                row.get("영수금액"), support_limits))
                    records.append({
                        "date": row.get("영수일자"),
                        "store": row.get("거래처명"),
                        "purpose": row.get("목적"),
                        "amount": row.get("영수금액"),
                        "payment": row.get("결제방식"),
                        "time": row.get("영수시간"),
                        "claim_amount": claim,
                        "region": row.get("지역"),
                        "participants": bi_name,   # 참여자 = 기초정보 성명
                        "note": row.get("비고"),
                    })
                if not records:
                    st.error("채울 데이터가 없습니다.")
                    return
                # 영수일자 오름차순 정렬 (날짜 없는 행은 맨 뒤로)
                records.sort(key=lambda r: (
                    _to_date(r.get("date")) is None,
                    _to_date(r.get("date")) or datetime.max.date(),
                ))
                buf, start, n = fill_workbook(
                    tpl_bytes, records, append=append,
                    basic_info={"dept": bi_dept, "name": bi_name, "title": bi_title},
                )
                stamp = datetime.now().strftime("%Y%m%d_%H%M")
                base = os.path.splitext(tpl_name)[0]
                st.session_state["gen_buf"] = buf.getvalue()
                st.session_state["gen_name"] = f"{base}_작성완료_{stamp}.xlsm"
                st.session_state["gen_msg"] = f"작성시트 {start}행부터 {n}건을 채웠습니다."
                # 다운로드 전 미리보기용: 실제 작성시트에 채워진 내용을 표로 보관
                st.session_state["gen_preview"] = pd.DataFrame([{
                    "영수일자": r["date"], "거래처명": r["store"], "목적": r["purpose"],
                    "영수금액": r["amount"], "청구금액": r["claim_amount"],
                    "결제방식": r["payment"], "영수시간": r["time"],
                    "지역": r["region"], "비고": r["note"],
                } for r in records])
                st.rerun()

            def _reset_all():
                """소속·성명·복지비·영수증·검토표·생성결과를 모두 비워 처음부터 시작한다.
                (위젯 값 초기화는 콜백에서 처리 — 위젯 생성 전에 실행되어야 안전)"""
                for k in ("prev_title", "show_thumbs", "df", "edited_snapshot",
                          "gen_buf", "gen_name", "gen_msg", "gen_preview",
                          "ignored_uploads"):
                    st.session_state.pop(k, None)
                st.session_state["analyzed_keys"] = set()
                # key를 바꿔 입력칸·업로더·편집표를 새 위젯으로 갈아끼워 확실히 비운다
                st.session_state["form_ver"] = st.session_state.get("form_ver", 0) + 1
                st.session_state["editor_ver"] = st.session_state.get("editor_ver", 0) + 1
                st.session_state["uploader_ver"] = st.session_state.get("uploader_ver", 0) + 1

            if tpl_bytes is None:
                st.info("사이드바에서 비용청구 양식(.xlsm)을 먼저 업로드하세요. "
                        "(기본 양식이 폴더에 있으면 자동으로 사용됩니다)")
            else:
                if st.button("📥 비용청구서 생성", type="primary",
                             width='stretch',
                             help="작성시트 첫 줄부터 채워 비용청구서를 만듭니다."):
                    _do_generate(append=False)

                # 한 번 생성해 다운로드가 뜬 뒤에만, '이어서 추가' 옵션을 작게 노출
                if st.session_state.get("gen_buf"):
                    st.success(st.session_state["gen_msg"])
                    if st.session_state.get("gen_preview") is not None:
                        st.caption("📋 생성된 청구서에 채워진 내용 미리보기 (다운로드 전 확인용)")
                        st.dataframe(st.session_state["gen_preview"],
                                     width='stretch', hide_index=True)
                    st.download_button(
                        "⬇️ 완성된 비용청구서 다운로드",
                        data=st.session_state["gen_buf"],
                        file_name=st.session_state["gen_name"],
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        width='stretch',
                    )
                    st.divider()
                    st.button("🆕 새로 작성 (전체 초기화)", on_click=_reset_all,
                              help="소속·성명·복지비·영수증·검토표를 모두 비우고 "
                                   "처음부터 새 청구서를 작성합니다.")


# =========================================================================
# 2) 연장근무(야근) 청구 모드
# =========================================================================
def render_overtime():
    hero(theme)

    with st.sidebar:
        st.markdown("##### 📄 양식")
        tpl_file = st.file_uploader(
            "연장근무(수당)신청서 양식(.xlsx)", type=["xlsx"],
            help="비워두면 기본 제공 양식(연장근무(수당)신청서_양식.xlsx)을 자동으로 사용합니다.",
        )
        tpl_bytes, tpl_name, tpl_is_default = _template_bytes_name(
            tpl_file, DEFAULT_OVERTIME_TPL)
        if tpl_is_default:
            st.caption(f"✅ 기본 양식 사용 중: **{tpl_name}**")
        elif tpl_bytes is None:
            st.caption("⚠️ 기본 양식을 찾지 못했습니다. 양식을 업로드하세요.")
        st.divider()
        st.caption("규칙: 근무시작 = 출근시간 + 9시간(정규8h+점심1h), "
                   "근무종료 = 퇴근시간. 근무시간·신청시간은 양식 수식이 자동 계산.")

    # ---- 1. 근태 미리보기 ------------------------------------------------
    with st.container(border=True):
        step(1, "근태현황 업로드 및 미리보기",
             "월간 근태현황(.xlsx)을 올리면 연장근무 대상일을 자동으로 찾아 보여줍니다.")
        att_file = st.file_uploader(
            "월간 근태현황(.xlsx)", type=["xlsx"],
            help="이름·조회기간·일자별 출퇴근/승인초과가 담긴 근태현황 파일.",
            label_visibility="collapsed",
        )
        if not att_file:
            st.info("위에서 월간 근태현황(.xlsx)을 업로드하세요.")
            return

        try:
            name, year, month, records = parse_attendance(att_file.getvalue())
        except Exception as e:  # noqa: BLE001
            st.error(f"근태현황을 읽지 못했습니다: {e}")
            return

        c1, c2, c3 = st.columns(3)
        c1.metric("👤 성명", name or "-")
        c2.metric("🗓️ 조회기간", f"{year or '-'}-{month:02d}" if month else "-")
        c3.metric("🌙 연장근무 대상일", f"{len(records)}일")

        st.caption("아래 부서명·직위는 신청서 상단의 '부서명 / 직위' 칸(D7)에 채워집니다. "
                   "성명은 근태현황에서 자동으로 가져옵니다.")
        d1, d2 = st.columns(2)
        ot_dept = d1.text_input("부서명", key="ot_dept", placeholder="예: 인공지능 개발팀")
        ot_pos = d2.text_input("직위", key="ot_pos", placeholder="예: 연구원")

        if not records:
            st.warning("'승인 초과 근로시간'이 0보다 큰 날이 없습니다. 채울 데이터가 없습니다.")
            return

    def _fmt(sec):
        return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"

    # ---- 2. 대체휴무 입력 ------------------------------------------------
    with st.container(border=True):
        step(2, "대체휴무 입력",
             "기본은 X(대체휴무 미지급) → 근무시간 전체가 신청시간에 기록됩니다. "
             "대체휴무를 받은 날만 O로 바꾸고 대체휴무시간을 넣으면 그만큼 신청시간에서 차감됩니다. "
             "대체휴무시간은 HH:MM(예: 01:30) 또는 시간 숫자(예: 1.5).")

        # 대체휴무지급 일괄 적용: 모든 행을 한 번에 X(기본) / O 로 채운다.
        bulk = st.radio(
            "대체휴무지급 일괄 적용",
            ["전체 X (기본·전체 신청)", "전체 O", "개별 선택"],
            horizontal=True,
            help="'전체 X'면 모든 날의 근무시간 전체가 신청됩니다. '개별 선택'이면 표에서 행마다 O/X를 고릅니다.",
            key="ot_bulk",
        )
        default_payoff = {"전체 O": "O"}.get(bulk, "X")  # 기본 X

        table = pd.DataFrame([
            {
                "일자": f"{month:02d}-{r['day']:02d}" if month else str(r["day"]),
                "출근": _fmt(r["clock_in"]),
                "퇴근": _fmt(r["clock_out"]),
                "근무시작(출근+9h)": _fmt(r["clock_in"] + 9 * 3600),
                "근무종료": _fmt(r["clock_out"]),
                "승인초과": _fmt(r["approved_ot"]),
                "대체휴무지급": default_payoff,
                "대체휴무시간": "",
                "비고": "",
            }
            for r in records
        ]).astype(str)
        edited = st.data_editor(
            table,
            width='stretch',
            hide_index=True,
            disabled=["일자", "출근", "퇴근", "근무시작(출근+9h)", "근무종료", "승인초과"],
            column_config={
                "대체휴무지급": st.column_config.SelectboxColumn(
                    "대체휴무지급", options=["X", "O"], required=True,
                    help="X=대체휴무 미지급(전체 신청), O=대체휴무 지급(대체휴무시간 차감)"),
                "대체휴무시간": st.column_config.TextColumn(
                    "대체휴무시간", help="O일 때만 사용. HH:MM 또는 시간 숫자(예: 1.5)"),
                "비고": st.column_config.TextColumn("비고"),
            },
            key="ot_editor",
        )

    # ---- 3. 생성 ---------------------------------------------------------
    with st.container(border=True):
        step(3, "연장근무신청서 생성")
        if tpl_bytes is None:
            st.info("사이드바에서 연장근무(수당)신청서 양식(.xlsx)을 업로드하세요. "
                    "(기본 양식이 폴더에 있으면 자동으로 사용됩니다)")
            return

        if st.button("📥 신청서 생성", type="primary", width='stretch'):
            # 표에서 고른 대체휴무지급/대체휴무시간/비고를 일자별로 모은다.
            extras = {}
            for rec, (_, row) in zip(records, edited.iterrows()):
                extras[rec["day"]] = {
                    "payoff": row.get("대체휴무지급", ""),
                    "hours": row.get("대체휴무시간", ""),
                    "note": row.get("비고", ""),
                }
            parts = [p for p in (ot_dept.strip(), ot_pos.strip()) if p]
            dept_position = " / ".join(parts) if parts else None
            try:
                buf, n = fill_overtime(tpl_bytes, att_file.getvalue(),
                                       extras=extras, dept_position=dept_position)
            except Exception as e:  # noqa: BLE001
                st.error(f"생성에 실패했습니다: {e}")
                return
            stamp = datetime.now().strftime("%Y%m%d_%H%M")
            base = os.path.splitext(tpl_name)[0]
            out_name = f"{base}_작성완료_{stamp}.xlsx"
            st.success(f"{n}건을 채웠습니다.")
            st.download_button(
                "⬇️ 완성된 연장근무신청서 다운로드",
                data=buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width='stretch',
            )


# ---- 라우팅 --------------------------------------------------------------
if mode.startswith("🧾"):
    render_expense()
else:
    render_overtime()
