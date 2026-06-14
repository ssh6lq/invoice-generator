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
from excel_filler import fill_workbook, get_dropdown_options
from overtime_filler import parse_attendance, fill_overtime

st.set_page_config(page_title="청구서 자동 작성", page_icon="🧾", layout="wide")

# ---- 모드 선택 -----------------------------------------------------------
with st.sidebar:
    st.header("📋 작업 선택")
    mode = st.radio(
        "무엇을 생성할까요?",
        ["🧾 비용청구", "🌙 연장근무(야근) 청구"],
        index=0,
    )
    st.divider()


# =========================================================================
# 1) 비용청구 모드
# =========================================================================
def render_expense():
    st.title("🧾 영수증 → 비용청구서 자동 작성")
    st.caption("영수증 사진을 올리면 날짜·상호명·금액을 인식해 비용청구 양식을 채워줍니다.")

    with st.sidebar:
        st.header("⚙️ 설정")
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
        xlsm_file = st.file_uploader(
            "비용청구 양식(.xlsm)", type=["xlsm", "xlsx"],
            help="작성시트가 있는 비용청구 양식 파일을 올리세요.",
        )
        append_mode = st.radio(
            "작성 방식",
            ["기존 데이터 뒤에 추가", "15행부터 새로 작성"],
            index=0,
        )

    # 양식이 올라오면 목적/결제방식 드롭다운 목록을 추출해 둔다.
    purpose_opts, payment_opts = [], []
    if xlsm_file is not None:
        try:
            opts = get_dropdown_options(xlsm_file.getvalue())
            purpose_opts = opts.get("purpose", [])
            payment_opts = opts.get("payment", [])
        except Exception as e:  # noqa: BLE001
            st.sidebar.warning(f"드롭다운 목록을 읽지 못했습니다: {e}")

    # ---- 영수증 업로드 ---------------------------------------------------
    st.subheader("1. 영수증 이미지 업로드")
    images = st.file_uploader(
        "영수증 사진 (여러 장 가능)",
        type=["jpg", "jpeg", "png", "webp", "bmp"],
        accept_multiple_files=True,
    )

    if images:
        cols = st.columns(min(len(images), 5))
        for i, img in enumerate(images):
            with cols[i % len(cols)]:
                st.image(img.getvalue(), caption=img.name, use_container_width=True)

    # ---- 분석 실행 -------------------------------------------------------
    st.subheader("2. 영수증 분석")
    if st.button("🔍 영수증 분석", type="primary", disabled=not images):
        if not api_key:
            st.error("사이드바에 OpenAI API Key를 입력하세요.")
        else:
            payload = [(img.name, img.getvalue()) for img in images]
            progress = st.progress(0.0, text="분석 준비 중...")

            def _cb(done, total, rec):
                label = rec.get("store") or rec["filename"]
                progress.progress(done / total, text=f"분석 중 {done}/{total} — {label}")

            with st.spinner("GPT로 영수증을 분석하는 중..."):
                results = parse_many(payload, model=model, api_key=api_key, on_progress=_cb)
            progress.empty()

            df = pd.DataFrame(results)[["date", "store", "amount", "time", "filename", "error"]]
            df.columns = ["영수일자", "거래처명", "영수금액", "영수시간", "파일명", "오류"]
            # 사용자가 직접 고르는 항목 (양식 순서: 거래처명 다음 목적, 영수금액 다음 결제방식)
            df.insert(2, "목적", "")
            df.insert(4, "결제방식", "")
            st.session_state["df"] = df

            errs = [r for r in results if r["error"]]
            if errs:
                st.warning(f"{len(errs)}건 분석 실패. 표의 '오류' 열을 확인하세요.")
            st.success(f"{len(results)}건 분석 완료. 아래 표에서 검토·수정하세요.")

    # ---- 검토/수정 표 ----------------------------------------------------
    if "df" in st.session_state:
        st.subheader("3. 검토 및 수정")
        st.caption("셀을 더블클릭하면 수정할 수 있습니다. 날짜는 YYYY-MM-DD, 시간은 HH:MM 형식. "
                   "목적·결제방식은 칸을 클릭해 목록에서 선택하세요.")
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
                "영수금액": st.column_config.NumberColumn(format="%d"),
                "오류": st.column_config.TextColumn(disabled=True),
            },
            key="editor",
        )

        # ---- 엑셀 생성 ---------------------------------------------------
        st.subheader("4. 비용청구서 생성")
        if not xlsm_file:
            st.info("사이드바에서 비용청구 양식(.xlsm)을 먼저 업로드하세요.")
        elif st.button("📥 엑셀 생성", type="primary"):
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
                })
            if not records:
                st.error("채울 데이터가 없습니다.")
            else:
                buf, start, n = fill_workbook(
                    xlsm_file.getvalue(),
                    records,
                    append=(append_mode == "기존 데이터 뒤에 추가"),
                )
                stamp = datetime.now().strftime("%Y%m%d_%H%M")
                base = os.path.splitext(xlsm_file.name)[0]
                out_name = f"{base}_작성완료_{stamp}.xlsm"
                st.success(f"작성시트 {start}행부터 {n}건을 채웠습니다.")
                st.download_button(
                    "⬇️ 완성된 비용청구서 다운로드",
                    data=buf,
                    file_name=out_name,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                )


# =========================================================================
# 2) 연장근무(야근) 청구 모드
# =========================================================================
def render_overtime():
    st.title("🌙 근태현황 → 연장근무(야근)신청서 자동 작성")
    st.caption("월간 근태현황을 올리면 '승인 초과 근로시간'이 있는 날을 찾아 연장근무신청서 양식을 채워줍니다.")

    with st.sidebar:
        st.header("⚙️ 파일")
        att_file = st.file_uploader(
            "월간 근태현황(.xlsx)", type=["xlsx"],
            help="이름·조회기간·일자별 출퇴근/승인초과가 담긴 근태현황 파일.",
        )
        tpl_file = st.file_uploader(
            "연장근무(수당)신청서 양식(.xlsx)", type=["xlsx"],
            help="'양식' 시트가 있는 빈 신청서 파일.",
        )
        st.divider()
        st.caption("규칙: 근무시작 = 출근시간 + 9시간(정규8h+점심1h), "
                   "근무종료 = 퇴근시간. 근무시간은 양식 수식이 자동 계산.")

    # ---- 근태 미리보기 ---------------------------------------------------
    st.subheader("1. 근태현황 미리보기")
    if not att_file:
        st.info("사이드바에서 월간 근태현황(.xlsx)을 업로드하세요.")
        return

    try:
        name, year, month, records = parse_attendance(att_file.getvalue())
    except Exception as e:  # noqa: BLE001
        st.error(f"근태현황을 읽지 못했습니다: {e}")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("성명", name or "-")
    c2.metric("조회기간", f"{year or '-'}-{month:02d}" if month else "-")
    c3.metric("연장근무 대상일", f"{len(records)}일")

    if not records:
        st.warning("'승인 초과 근로시간'이 0보다 큰 날이 없습니다. 채울 데이터가 없습니다.")
        return

    def _fmt(sec):
        return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"

    prev = pd.DataFrame([
        {
            "일자": f"{month:02d}-{r['day']:02d}" if month else r["day"],
            "출근": _fmt(r["clock_in"]),
            "퇴근": _fmt(r["clock_out"]),
            "근무시작(출근+9h)": _fmt(r["clock_in"] + 9 * 3600),
            "근무종료": _fmt(r["clock_out"]),
            "승인초과": _fmt(r["approved_ot"]),
        }
        for r in records
    ])
    st.dataframe(prev, use_container_width=True, hide_index=True)

    # ---- 생성 -----------------------------------------------------------
    st.subheader("2. 연장근무신청서 생성")
    if not tpl_file:
        st.info("사이드바에서 연장근무(수당)신청서 양식(.xlsx)을 업로드하세요.")
        return

    if st.button("📥 신청서 생성", type="primary"):
        try:
            buf, n = fill_overtime(tpl_file.getvalue(), att_file.getvalue())
        except Exception as e:  # noqa: BLE001
            st.error(f"생성에 실패했습니다: {e}")
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        base = os.path.splitext(tpl_file.name)[0]
        out_name = f"{base}_작성완료_{stamp}.xlsx"
        st.success(f"{n}건을 채웠습니다.")
        st.download_button(
            "⬇️ 완성된 연장근무신청서 다운로드",
            data=buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ---- 라우팅 --------------------------------------------------------------
if mode.startswith("🧾"):
    render_expense()
else:
    render_overtime()
