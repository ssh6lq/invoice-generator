"""
app.py — 청구서 자동 작성 (Streamlit)

두 가지 모드를 선택해서 사용한다.
  1) 🧾 비용청구  : 영수증 이미지를 GPT 비전으로 분석해 비용청구 양식(.xlsm) 작성
  2) 🌙 연장근무(야근) 청구 : 월간 근태현황(.xlsx)을 읽어 연장근무(수당)신청서 양식 작성

실행:  streamlit run app.py
"""

import os
from datetime import datetime

import pandas as pd
import streamlit as st

from receipt_parser import parse_many
from excel_filler import fill_workbook, get_dropdown_options, get_support_limits, _to_date
from overtime_filler import parse_attendance, fill_overtime

st.set_page_config(page_title="청구서 자동 작성", page_icon="🧾", layout="wide")

# 번들 기본 양식 (사용자가 따로 업로드하지 않으면 이 파일을 사용)
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_EXPENSE_TPL = os.path.join(APP_DIR, "비용청구양식.xlsm")
DEFAULT_OVERTIME_TPL = os.path.join(APP_DIR, "연장근무(수당)신청서_양식.xlsx")


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
        "accent": "#6c5ce7", "accent2": "#8e7bff", "bg": "#f6f5fc",
        "icon": "🧾", "title": "영수증 → 비용청구서 자동 작성",
        "desc": "영수증 사진을 올리면 날짜·상호명·금액을 인식해 비용청구 양식을 자동으로 채웁니다.",
    },
    "overtime": {
        "accent": "#3b6fd4", "accent2": "#5c93f0", "bg": "#f2f6fd",
        "icon": "🌙", "title": "근태현황 → 연장근무신청서 자동 작성",
        "desc": "월간 근태현황을 올리면 '승인 초과 근로시간'이 있는 날을 찾아 신청서를 자동으로 채웁니다.",
    },
}


def inject_css(t):
    st.markdown(
        f"""
        <style>
        .stApp {{ background: {t['bg']}; }}
        .block-container {{ padding-top: 2.2rem; max-width: 1100px; }}

        /* 히어로 배너 */
        .hero {{
            background: linear-gradient(120deg, {t['accent']}, {t['accent2']});
            color: #fff; padding: 24px 28px; border-radius: 18px;
            margin-bottom: 18px;
            box-shadow: 0 10px 28px {t['accent']}33;
        }}
        .hero h1 {{ margin: 0; font-size: 1.55rem; font-weight: 800; letter-spacing: -.3px; }}
        .hero p {{ margin: .45rem 0 0; opacity: .92; font-size: .96rem; }}

        /* 단계 배지 헤더 */
        .step {{ display: flex; align-items: center; gap: .6rem; margin: .4rem 0 .2rem; }}
        .step .num {{
            background: {t['accent']}; color: #fff; width: 30px; height: 30px;
            border-radius: 50%; display: flex; align-items: center; justify-content: center;
            font-weight: 700; font-size: .95rem; flex: none;
            box-shadow: 0 3px 8px {t['accent']}55;
        }}
        .step .ttl {{ font-size: 1.18rem; font-weight: 800; color: #2d3436; }}

        /* 버튼 */
        .stButton > button, .stDownloadButton > button {{
            border-radius: 11px; font-weight: 700; padding: .5rem 1rem;
        }}
        .stButton > button[kind="primary"], .stDownloadButton > button {{
            background: {t['accent']}; border: none;
        }}
        .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {{
            background: {t['accent2']};
        }}

        /* 지표 카드 */
        div[data-testid="stMetric"] {{
            background: #fff; border: 1px solid #eceaf6; border-radius: 14px;
            padding: 12px 16px; box-shadow: 0 2px 10px rgba(0,0,0,.04);
        }}
        div[data-testid="stMetricValue"] {{ font-size: 1.45rem; }}

        /* 테두리 컨테이너(카드) */
        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background: #fff; border-radius: 14px;
        }}

        /* 사이드바 */
        section[data-testid="stSidebar"] {{ background: #fff; border-right: 1px solid #ececf4; }}
        .side-brand {{
            font-size: 1.15rem; font-weight: 800; color: {t['accent']};
            display: flex; align-items: center; gap: .4rem; margin-bottom: .2rem;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero(t):
    st.markdown(
        f'<div class="hero"><h1>{t["icon"]} {t["title"]}</h1>'
        f'<p>{t["desc"]}</p></div>',
        unsafe_allow_html=True,
    )


def step(n, title, desc=None):
    st.markdown(
        f'<div class="step"><span class="num">{n}</span>'
        f'<span class="ttl">{title}</span></div>',
        unsafe_allow_html=True,
    )
    if desc:
        st.caption(desc)


def _img_key(f):
    """업로드 파일의 중복 판별 키 (이름 + 크기)."""
    return f"{f.name}::{f.size}"


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
    st.markdown('<div class="side-brand">🧾 청구서 자동 작성</div>', unsafe_allow_html=True)
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
        append_mode = st.radio(
            "작성 방식",
            ["기존 데이터 뒤에 추가", "15행부터 새로 작성"],
            index=0,
        )
        st.divider()
        st.markdown("##### 💳 법인카드")
        bi_card = st.text_input("법인카드번호 뒤 4자리", max_chars=4, key="bi_card",
                                help="개인할당 법인카드 결제 시. 없으면 비워두세요.")

    # 양식에서 목적/결제방식 드롭다운 목록 + 목적별 지원한도를 추출해 둔다.
    purpose_opts, payment_opts, support_limits = [], [], {}
    if tpl_bytes is not None:
        try:
            opts = get_dropdown_options(tpl_bytes)
            purpose_opts = opts.get("purpose", [])
            payment_opts = opts.get("payment", [])
            support_limits = get_support_limits(tpl_bytes)
        except Exception as e:  # noqa: BLE001
            st.sidebar.warning(f"드롭다운 목록을 읽지 못했습니다: {e}")

    # ---- 1. 기초정보 입력 ------------------------------------------------
    with st.container(border=True):
        step(1, "기초정보 입력",
             "비용청구서에 들어갈 기본 정보예요. 먼저 입력하세요. "
             "성명을 넣으면 본인확인이 '서명완료'로 처리되어 엑셀에선 매크로검토만 누르면 됩니다.")
        ci1, ci2, ci3 = st.columns(3)
        bi_dept = ci1.text_input("소속부서명", key="bi_dept")
        bi_name = ci2.text_input("성명", key="bi_name")
        bi_title = ci3.selectbox("청구항목명", ["비용", "복지비"], key="bi_title",
                                 help="선택값 뒤에 '청구서'가 붙어 제목이 됩니다(예: '비용청구서').")

    # ---- 2. 영수증 업로드 ------------------------------------------------
    with st.container(border=True):
        step(2, "영수증 이미지 업로드",
             "여러 장을 한 번에 올릴 수 있어요. 분석한 뒤 사진을 더 추가해도 됩니다.")
        images = st.file_uploader(
            "영수증 사진",
            type=["jpg", "jpeg", "png", "webp", "bmp"],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )
        if images:
            st.caption(f"🖼️ {len(images)}장 업로드됨")
            cols = st.columns(min(len(images), 5))
            for i, img in enumerate(images):
                with cols[i % len(cols)]:
                    st.image(img.getvalue(), caption=img.name, use_container_width=True)

    # ---- 3. 분석 실행 (새로 추가된 사진만) -------------------------------
    with st.container(border=True):
        step(3, "영수증 분석",
             "새로 추가한 영수증만 분석해서 기존 목록에 덧붙입니다. (이미 분석한 사진은 건너뜀)")

        analyzed = st.session_state.setdefault("analyzed_keys", set())
        pending = [im for im in (images or []) if _img_key(im) not in analyzed]

        b1, b2 = st.columns([3, 1])
        with b1:
            run = st.button(
                f"🔍 새 영수증 분석 ({len(pending)}장)",
                type="primary", disabled=not pending, use_container_width=True,
            )
        with b2:
            if st.button("🗑️ 목록 초기화", disabled="df" not in st.session_state,
                         use_container_width=True,
                         help="누적된 분석 결과를 모두 지우고 처음부터 다시 시작합니다."):
                st.session_state.pop("df", None)
                st.session_state["analyzed_keys"] = set()
                st.rerun()

        if images and not pending and run is False:
            st.info("새로 추가된 영수증이 없습니다. (올라온 사진은 모두 분석됨)")
        elif not images:
            st.info("먼저 위에서 영수증 사진을 올려주세요.")

        if run:
            if not api_key:
                st.error("사이드바에 OpenAI API Key를 입력하세요.")
            else:
                payload = [(im.name, im.getvalue()) for im in pending]
                progress = st.progress(0.0, text="분석 준비 중...")

                def _cb(done, total, rec):
                    label = rec.get("store") or rec["filename"]
                    progress.progress(done / total, text=f"분석 중 {done}/{total} — {label}")

                with st.spinner("GPT로 새 영수증을 분석하는 중..."):
                    results = parse_many(payload, model=model, api_key=api_key, on_progress=_cb)
                progress.empty()

                new_df = _results_to_df(results)
                if "df" in st.session_state:
                    st.session_state["df"] = pd.concat(
                        [st.session_state["df"], new_df], ignore_index=True)
                else:
                    st.session_state["df"] = new_df
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
            step(4, "검토 및 수정",
                 "셀을 더블클릭해 수정하세요. 날짜 YYYY-MM-DD, 시간 HH:MM. "
                 "목적을 고르면 청구금액이 그 목적의 지원한도로 자동 보정됩니다(한도 초과분 차감).")
            if not purpose_opts:
                st.info("목적·결제방식 선택지를 채우려면 사이드바에서 비용청구 양식(.xlsm)을 업로드하세요.")
            edited = st.data_editor(
                st.session_state["df"],
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "목적": st.column_config.SelectboxColumn(
                        "목적", options=purpose_opts, required=False,
                        help="양식의 목적 목록에서 선택"),
                    "결제방식": st.column_config.SelectboxColumn(
                        "결제방식", options=payment_opts, required=False,
                        help="양식의 결제방식 목록에서 선택"),
                    "영수금액": st.column_config.NumberColumn("영수금액", format="%d"),
                    "청구금액": st.column_config.NumberColumn(
                        "청구금액", format="%d",
                        help="실제 청구할 금액. 기본은 영수금액이며 부분 청구 시 수정하세요."),
                    "지역": st.column_config.TextColumn(
                        "지역", help="영수증 주소에서 자동 추출(구). 필요 시 수정하세요."),
                    "비고": st.column_config.TextColumn("비고", help="기본은 비워둠"),
                    "오류": st.column_config.TextColumn(disabled=True),
                },
            )

            # 목적별 지원한도로 청구금액 상한 적용 (한도 초과분은 한도로 보정)
            capped = False
            for idx in edited.index:
                purpose = edited.at[idx, "목적"]
                lim = support_limits.get(purpose) if purpose else None
                if lim is None:
                    continue
                cur = edited.at[idx, "청구금액"]
                base = cur if pd.notna(cur) and cur != "" else edited.at[idx, "영수금액"]
                if pd.notna(base) and base != "":
                    new_val = min(int(base), int(lim))
                    if cur != new_val:
                        edited.at[idx, "청구금액"] = new_val
                        capped = True

            # 편집 내용을 누적 결과에 반영(추가 분석 시 보존)
            st.session_state["df"] = edited
            if capped:
                st.rerun()  # 보정된 청구금액을 표에 즉시 반영

        # ---- 5. 엑셀 생성 ------------------------------------------------
        with st.container(border=True):
            step(5, "비용청구서 생성")
            if tpl_bytes is None:
                st.info("사이드바에서 비용청구 양식(.xlsm)을 먼저 업로드하세요. "
                        "(기본 양식이 폴더에 있으면 자동으로 사용됩니다)")
            elif st.button("📥 엑셀 생성", type="primary", use_container_width=True):
                records = []
                for _, row in edited.iterrows():
                    if not row.get("거래처명") and not row.get("영수일자"):
                        continue
                    records.append({
                        "date": row.get("영수일자"),
                        "store": row.get("거래처명"),
                        "purpose": row.get("목적"),
                        "amount": row.get("영수금액"),
                        "payment": row.get("결제방식"),
                        "time": row.get("영수시간"),
                        "claim_amount": row.get("청구금액"),
                        "region": row.get("지역"),
                        "participants": bi_name,   # 참여자 = 기초정보 성명
                        "note": row.get("비고"),
                    })
                if not records:
                    st.error("채울 데이터가 없습니다.")
                else:
                    # 영수일자 오름차순 정렬 (날짜 없는 행은 맨 뒤로)
                    records.sort(key=lambda r: (
                        _to_date(r.get("date")) is None,
                        _to_date(r.get("date")) or datetime.max.date(),
                    ))
                    buf, start, n = fill_workbook(
                        tpl_bytes,
                        records,
                        append=(append_mode == "기존 데이터 뒤에 추가"),
                        basic_info={
                            "dept": bi_dept,
                            "name": bi_name,
                            "card": bi_card,
                            "title": bi_title,
                        },
                    )
                    stamp = datetime.now().strftime("%Y%m%d_%H%M")
                    base = os.path.splitext(tpl_name)[0]
                    out_name = f"{base}_작성완료_{stamp}.xlsm"
                    st.success(f"작성시트 {start}행부터 {n}건을 채웠습니다.")
                    st.download_button(
                        "⬇️ 완성된 비용청구서 다운로드",
                        data=buf,
                        file_name=out_name,
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        use_container_width=True,
                    )


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
            use_container_width=True,
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

        if st.button("📥 신청서 생성", type="primary", use_container_width=True):
            # 표에서 고른 대체휴무지급/대체휴무시간/비고를 일자별로 모은다.
            extras = {}
            for rec, (_, row) in zip(records, edited.iterrows()):
                extras[rec["day"]] = {
                    "payoff": row.get("대체휴무지급", ""),
                    "hours": row.get("대체휴무시간", ""),
                    "note": row.get("비고", ""),
                }
            try:
                buf, n = fill_overtime(tpl_bytes, att_file.getvalue(), extras=extras)
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
                use_container_width=True,
            )


# ---- 라우팅 --------------------------------------------------------------
if mode.startswith("🧾"):
    render_expense()
else:
    render_overtime()
