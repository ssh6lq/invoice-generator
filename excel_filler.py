"""
excel_filler.py
업로드된 비용청구 .xlsm 의 '작성시트'에 영수증 데이터를 채운다.

★ 도형 보존 방식 ★
openpyxl 로 재저장하면 버튼 그림·도형(텍스트박스/그룹)·VML 등 drawing 객체가
사라진다. 그래서 이 모듈은 .xlsm 을 zip 으로 열어 해당 시트의 XML 에서
대상 셀 값만 직접 교체한다. 나머지(도형/이미지/VBA/서식/수식)는 원본 그대로
복사되므로 100% 보존된다.

채우는 열: C(영수일자), D(거래처명), E(목적), F(영수금액),
           G(결제방식), J(영수시간)
목적/결제방식은 양식의 드롭다운 목록에서 사용자가 고른 값을 그대로 기재한다.
"""

import re
import zipfile
from copy import copy  # noqa: F401  (호환용 import 유지)
from datetime import datetime, date, time
from io import BytesIO
from xml.sax.saxutils import escape

SHEET_NAME = "작성시트"
FIRST_DATA_ROW = 15        # 실제 데이터 시작 행
EXCEL_EPOCH = date(1899, 12, 30)


# ---------------------------------------------------------------- 변환 함수
def _to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip().replace(".", "-").replace("/", "-")
    s = "-".join(p for p in s.split("-") if p != "")
    for fmt in ("%Y-%m-%d", "%y-%m-%d", "%m-%d"):
        try:
            d = datetime.strptime(s, fmt)
            if fmt == "%m-%d":
                d = d.replace(year=datetime.now().year)
            return d.date()
        except ValueError:
            continue
    return None


def _to_time(v):
    if v is None or v == "":
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    s = str(v).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            continue
    return None


def _to_amount(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = "".join(ch for ch in str(v) if ch.isdigit())
    return int(s) if s else None


def _date_serial(d: date) -> int:
    return (d - EXCEL_EPOCH).days


def _time_fraction(t: time) -> float:
    return (t.hour * 3600 + t.minute * 60 + t.second) / 86400.0


# ---------------------------------------------------------------- zip 헬퍼
def _read_bytes(src):
    if isinstance(src, (bytes, bytearray)):
        return bytes(src)
    if hasattr(src, "read"):
        return src.read()
    with open(src, "rb") as f:
        return f.read()


def _sheet_path_for(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """workbook.xml + rels 를 읽어 시트 이름 -> worksheets/sheetN.xml 매핑."""
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        if m.group(1) == sheet_name:
            rid = m.group(2)
            break
    if rid is None:
        # name 과 r:id 순서가 바뀐 경우 대비
        for m in re.finditer(r'<sheet [^>]*r:id="([^"]+)"[^>]*name="([^"]+)"', wb):
            if m.group(2) == sheet_name:
                rid = m.group(1)
                break
    if rid is None:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")

    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for m in re.finditer(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
        if m.group(1) == rid:
            target = m.group(2)
            return "xl/" + target.lstrip("/").replace("xl/", "", 1) if not target.startswith("xl/") else target
    # 일반적 형태: Target="worksheets/sheetN.xml"
    for m in re.finditer(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
        if m.group(1) == rid:
            return "xl/" + m.group(2)
    raise ValueError(f"시트 경로를 찾을 수 없습니다: rid={rid}")


# ---------------------------------------------------------------- 셀 편집
def _cell_block(xml: str, coord: str):
    """coord(예: 'C25') 셀의 (start, end, style, has_value) 반환. 없으면 None."""
    m = re.search(r'<c r="%s"((?:\s+[^>]*?)?)(/>|>.*?</c>)' % re.escape(coord), xml, re.S)
    if not m:
        return None
    attrs = m.group(1)
    sm = re.search(r's="(\d+)"', attrs)
    style = sm.group(1) if sm else None
    has_value = "<v>" in m.group(2) or "<is>" in m.group(2)
    return m.start(), m.end(), style, has_value


def _build_cell(coord, style, kind, value):
    s_attr = f' s="{style}"' if style is not None else ""
    if kind == "num":
        return f'<c r="{coord}"{s_attr}><v>{value}</v></c>'
    if kind == "str":
        txt = escape(str(value))
        return (f'<c r="{coord}"{s_attr} t="inlineStr">'
                f'<is><t xml:space="preserve">{txt}</t></is></c>')
    raise ValueError(kind)


def _set_cell(xml: str, coord: str, kind: str, value) -> str:
    blk = _cell_block(xml, coord)
    if blk is None:
        raise ValueError(f"셀 {coord} 을(를) 찾을 수 없습니다 (양식 범위 초과).")
    start, end, style, _ = blk
    new = _build_cell(coord, style, kind, value)
    return xml[:start] + new + xml[end:]


def _first_empty_row(xml: str) -> int:
    r = FIRST_DATA_ROW
    while True:
        c = _cell_block(xml, f"C{r}")
        d = _cell_block(xml, f"D{r}")
        c_empty = (c is None) or (not c[3])
        d_empty = (d is None) or (not d[3])
        if c_empty and d_empty:
            return r
        r += 1


# ---------------------------------------------------------------- 드롭다운 목록
def _shared_strings(zf: zipfile.ZipFile):
    """sharedStrings.xml -> index 순서의 문자열 리스트."""
    try:
        xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.finditer(r"<si>(.*?)</si>", xml, re.S):
        body = si.group(1)
        # <si> 안의 모든 <t>...</t> 합치기 (rich text 대응)
        text = "".join(re.findall(r"<t[^>]*>(.*?)</t>", body, re.S))
        text = (text.replace("&amp;", "&").replace("&lt;", "<")
                .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
        out.append(text)
    return out


def _col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _read_range_column(zf, sheet_name, col, row1, row2):
    """특정 시트의 한 열(col) row1~row2 셀 값을 리스트로 반환."""
    path = _sheet_path_for(zf, sheet_name)
    xml = zf.read(path).decode("utf-8")
    shared = _shared_strings(zf)
    out = []
    for r in range(row1, row2 + 1):
        coord = f"{col}{r}"
        m = re.search(r'<c r="%s"((?:\s+[^>]*?)?)(?:/>|>(.*?)</c>)' % re.escape(coord),
                      xml, re.S)
        if not m:
            continue
        attrs, inner = m.group(1), (m.group(2) or "")
        t = re.search(r't="([^"]+)"', attrs)
        ttype = t.group(1) if t else None
        vm = re.search(r"<v>(.*?)</v>", inner, re.S)
        if ttype == "s" and vm:
            idx = int(vm.group(1))
            if 0 <= idx < len(shared):
                out.append(shared[idx])
        elif ttype == "inlineStr":
            im = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
            if im:
                out.append(im.group(1))
        elif vm:
            out.append(vm.group(1))
    return [v for v in (s.strip() for s in out) if v]


def get_dropdown_options(src_path_or_bytes):
    """
    양식에서 목적(E)·결제방식(G) 드롭다운 선택지를 추출한다.
    반환: {"purpose": [...], "payment": [...]}
    파싱 실패 시 빈 리스트.
    """
    raw = _read_bytes(src_path_or_bytes)
    zf = zipfile.ZipFile(BytesIO(raw))
    sheet_path = _sheet_path_for(zf, SHEET_NAME)
    xml = zf.read(sheet_path).decode("utf-8")

    payment, purpose = [], []

    # 결제방식(G): 표준 dataValidation list, formula1 이 따옴표 문자열
    for m in re.finditer(r"<dataValidation\b[^>]*?type=\"list\"[^>]*?>(.*?)</dataValidation>",
                         xml, re.S):
        block = m.group(0)
        sq = re.search(r'sqref="([^"]*)"', block)
        if sq and re.search(r'\bG\d', sq.group(1)):
            f1 = re.search(r"<formula1>\"?(.*?)\"?</formula1>", m.group(1), re.S)
            if f1:
                payment = [s.strip() for s in f1.group(1).split(",") if s.strip()]
            break

    # 목적(E): x14 dataValidation, formula1 이 다른 시트 범위 참조
    mx = re.search(r"<x14:dataValidation\b.*?</x14:dataValidation>", xml, re.S)
    for m in re.finditer(r"<x14:dataValidation\b(.*?)</x14:dataValidation>", xml, re.S):
        block = m.group(1)
        sq = re.search(r"<xm:sqref>([^<]*)</xm:sqref>", block)
        if sq and re.search(r'\bE\d', sq.group(1)):
            f = re.search(r"<xm:f>(.*?)</xm:f>", block, re.S)
            if f:
                ref = f.group(1)  # 예: 비용지원안내!$B$5:$B$42
                rm = re.match(r"(?:'?)([^'!]+)(?:'?)!\$?([A-Z]+)\$?(\d+):\$?[A-Z]+\$?(\d+)", ref)
                if rm:
                    sn, col, r1, r2 = rm.group(1), rm.group(2), int(rm.group(3)), int(rm.group(4))
                    purpose = _read_range_column(zf, sn, col, r1, r2)
            break

    return {"purpose": purpose, "payment": payment}


# ---------------------------------------------------------------- 메인
def fill_workbook(src_path_or_bytes, records, append=True):
    """
    영수증 레코드를 워크북에 채워 BytesIO 로 반환한다.
    원본의 도형/이미지/매크로/서식을 100% 보존한다.

    records: list[dict]  키 = date, store, purpose, amount, payment, time
    append : True 면 기존 데이터 다음 빈 행부터, False 면 15행부터
    반환    : (BytesIO, start_row, count)
    """
    raw = _read_bytes(src_path_or_bytes)
    zin = zipfile.ZipFile(BytesIO(raw))
    sheet_path = _sheet_path_for(zin, SHEET_NAME)
    xml = zin.read(sheet_path).decode("utf-8")

    start = _first_empty_row(xml) if append else FIRST_DATA_ROW

    for i, rec in enumerate(records):
        r = start + i
        d = _to_date(rec.get("date"))
        if d is not None:
            xml = _set_cell(xml, f"C{r}", "num", _date_serial(d))
        store = rec.get("store")
        if store:
            xml = _set_cell(xml, f"D{r}", "str", str(store).strip())
        purpose = rec.get("purpose")
        if purpose:
            xml = _set_cell(xml, f"E{r}", "str", str(purpose).strip())
        payment = rec.get("payment")
        if payment:
            xml = _set_cell(xml, f"G{r}", "str", str(payment).strip())
        amt = _to_amount(rec.get("amount"))
        if amt is not None:
            xml = _set_cell(xml, f"F{r}", "num", amt)
        t = _to_time(rec.get("time"))
        if t is not None:
            xml = _set_cell(xml, f"J{r}", "num", repr(_time_fraction(t)))

    # 새 zip 작성 (대상 시트만 교체, 나머지 원본 그대로 복사)
    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sheet_path:
                data = xml.encode("utf-8")
            # 압축 방식/속성 보존
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zi.internal_attr = item.internal_attr
            zi.create_system = item.create_system
            zout.writestr(zi, data)
    zin.close()
    out.seek(0)
    return out, start, len(records)


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "비용청구_남소희_20260604.xlsm"
    opts = get_dropdown_options(path)
    print("목적 옵션:", len(opts["purpose"]), opts["purpose"][:3], "...")
    print("결제방식 옵션:", opts["payment"])
    demo = [
        {"date": "2026-06-01", "store": "테스트상회", "purpose": "야근식비",
         "amount": "12,500", "payment": "1.개인카드", "time": "19:30"},
        {"date": "2026.06.02", "store": "분식나라 <김밥>", "purpose": "기타식비",
         "amount": 8000, "payment": "2.현금", "time": "20:05"},
    ]
    buf, start, n = fill_workbook(path, demo, append=True)
    with open("test_output.xlsm", "wb") as f:
        f.write(buf.read())
    print(f"채움 완료: {start}행부터 {n}건 -> test_output.xlsm")
