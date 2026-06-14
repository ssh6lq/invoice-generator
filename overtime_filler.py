"""
overtime_filler.py
월간 근태현황(.xlsx)을 읽어 '연장근무(수당)신청서' 양식을 채운다.

규칙 (사용자 지정)
  - 포함 대상: '승인 초과 근로시간' > 0 인 날만.
  - 근무시작(I) = 출근시간 + 9시간   (표준근무 8h + 점심 1h)
                  예) 08:00 출근 -> 17:00, 09:12 출근 -> 18:12
  - 근무종료(J) = 퇴근시간
  - 근무시간(K) = 양식 수식이 J-I 로 자동 계산 (0.5시간 단위)
  - 실 근무시작(L) = 근무시작(I), 실 근무종료(M) = 근무종료(J)

★ 도형 보존 ★
양식에는 결재칸 등 도형/VML 이 있어, openpyxl 재저장 시 사라진다.
그래서 .xlsx 를 zip 으로 열어 시트 XML 의 대상 셀 값만 직접 교체한다.
근무시간(K)·신청시간(S) 등 수식과 날짜 자동생성(C2 기반)은 그대로 둔다.
"""

import re
import zipfile
from io import BytesIO

import openpyxl

# excel_filler 의 범용 zip/XML 헬퍼 재사용
from excel_filler import _read_bytes, _sheet_path_for, _set_cell

FORM_SHEET = "양식"
STANDARD_WORK_SECONDS = 9 * 3600   # 정규근무 9시간(점심 포함)
DAY_SECONDS = 86400


# ---------------------------------------------------------------- 시간 파싱
def _parse_hms(v):
    """'HH:MM:SS' 또는 time -> 총 초(int). 빈값/0 이면 None 또는 0."""
    if v is None or v == "":
        return None
    if hasattr(v, "hour"):  # datetime.time / datetime
        return v.hour * 3600 + v.minute * 60 + getattr(v, "second", 0)
    s = str(v).strip()
    parts = s.split(":")
    try:
        parts = [int(p) for p in parts]
    except ValueError:
        return None
    while len(parts) < 3:
        parts.append(0)
    h, m, sec = parts[:3]
    return h * 3600 + m * 60 + sec


def _fraction(seconds):
    return seconds / DAY_SECONDS


# ---------------------------------------------------------------- 근태 읽기
def parse_attendance(src_path_or_bytes):
    """
    월간 근태현황을 읽어 (name, year, month, records) 반환.
    records: list[dict] 키 = day(int), clock_in(sec), clock_out(sec),
             approved_ot(sec)  — 승인초과>0 인 날만.
    """
    raw = _read_bytes(src_path_or_bytes)
    wb = openpyxl.load_workbook(BytesIO(raw), data_only=True)
    ws = wb.active

    # 이름: 'B2 = "남소희 월간 근태현황"'
    title = str(ws["B2"].value or "").strip()
    name = title.split()[0] if title else ""

    # 조회기간: 'C5 = "202605"'
    period = str(ws["C5"].value or "").strip()
    year, month = None, None
    if len(period) >= 6 and period[:6].isdigit():
        year, month = int(period[:4]), int(period[4:6])

    # 헤더가 7행, 데이터 8행부터. 열: B(일자) C(출근) F(퇴근) O(승인초과)
    col = {c.value: c.column for c in ws[7] if c.value}
    c_date = col.get("일자", 2)
    c_in = col.get("출근시간", 3)
    c_out = col.get("퇴근시간", 6)
    c_ot = col.get("승인 초과 근로시간", 15)

    records = []
    for r in range(8, ws.max_row + 1):
        dval = ws.cell(r, c_date).value
        if not dval:
            continue
        ot = _parse_hms(ws.cell(r, c_ot).value) or 0
        cin = _parse_hms(ws.cell(r, c_in).value)
        cout = _parse_hms(ws.cell(r, c_out).value)
        if ot <= 0 or cin is None or cout is None:
            continue
        # 일자(day) 추출
        ds = str(dval)
        day = None
        m = re.search(r"-(\d{2})$", ds) or re.search(r"-(\d{1,2})\b", ds)
        if hasattr(dval, "day"):
            day = dval.day
        elif m:
            day = int(m.group(1))
        if day is None:
            continue
        if year is None and hasattr(dval, "year"):
            year, month = dval.year, dval.month
        records.append({"day": day, "clock_in": cin,
                        "clock_out": cout, "approved_ot": ot})
    return name, year, month, records


# ---------------------------------------------------------------- 양식 채우기
def fill_overtime(template_path_or_bytes, attendance_path_or_bytes,
                  name=None, month=None):
    """
    근태현황을 읽어 연장근무신청서 양식을 채워 (BytesIO, count) 반환.
    name/month 를 직접 주면 근태 파일 값보다 우선한다.
    """
    a_name, a_year, a_month, records = parse_attendance(attendance_path_or_bytes)
    name = name or a_name
    month = month or a_month

    raw = _read_bytes(template_path_or_bytes)
    zin = zipfile.ZipFile(BytesIO(raw))
    sheet_path = _sheet_path_for(zin, FORM_SHEET)
    xml = zin.read(sheet_path).decode("utf-8")

    # 기본정보: 월(C2), 성명(D8)
    if month:
        xml = _set_cell(xml, "C2", "num", int(month))
    if name:
        xml = _set_cell(xml, "D8", "str", name)

    # 일자별 행 채우기 (양식: 1일=17행, day -> 16+day)
    for rec in records:
        r = 16 + rec["day"]
        i_sec = rec["clock_in"] + STANDARD_WORK_SECONDS   # 근무시작 = 출근+9h
        j_sec = rec["clock_out"]                          # 근무종료 = 퇴근
        xml = _set_cell(xml, f"I{r}", "num", repr(_fraction(i_sec)))
        xml = _set_cell(xml, f"J{r}", "num", repr(_fraction(j_sec)))
        xml = _set_cell(xml, f"L{r}", "num", repr(_fraction(i_sec)))  # 실 근무시작
        xml = _set_cell(xml, f"M{r}", "num", repr(_fraction(j_sec)))  # 실 근무종료

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = xml.encode("utf-8")
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zout.writestr(zi, data)
    zin.close()
    out.seek(0)
    return out, len(records)


if __name__ == "__main__":
    import sys
    att = sys.argv[1] if len(sys.argv) > 1 else "남소희_월간근태현황_202605.xlsx"
    tpl = sys.argv[2] if len(sys.argv) > 2 else "연장근무(수당)신청서_양식.xlsx"
    name, year, month, recs = parse_attendance(att)
    print(f"이름={name} 연={year} 월={month} 대상일수={len(recs)}")
    for rc in recs:
        print(rc)
    buf, n = fill_overtime(tpl, att)
    with open("test_overtime.xlsx", "wb") as f:
        f.write(buf.read())
    print(f"채움 완료: {n}건 -> test_overtime.xlsx")
