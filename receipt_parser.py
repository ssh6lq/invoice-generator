"""
receipt_parser.py
LangChain + OpenAI(GPT) 비전 모델로 영수증 이미지를 파싱한다.
영수증 1장에서 날짜·상호명·금액·시간을 구조화된 JSON으로 추출한다.

필요 패키지: langchain-openai, langchain-core, pydantic, pillow
환경변수: OPENAI_API_KEY
"""

import base64
import mimetypes
from typing import List, Optional

from pydantic import BaseModel, Field
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, SystemMessage


# ---- 추출 결과 스키마 ----------------------------------------------------
class Receipt(BaseModel):
    """영수증 1건에서 추출한 정보."""
    date: Optional[str] = Field(
        None, description="영수일자. 반드시 YYYY-MM-DD 형식. 없으면 null"
    )
    store: Optional[str] = Field(
        None,
        description="상호명(거래처명). '가맹점 정보'/'판매자 정보' 구분이 없는 일반 영수증의 "
                    "상호. 없으면 null",
    )
    amount: Optional[int] = Field(
        None, description="총 결제(영수) 금액. 숫자만(원 단위 정수). 없으면 null"
    )
    time: Optional[str] = Field(
        None, description="결제 시각. HH:MM(24시간) 형식. 없으면 null"
    )
    region: Optional[str] = Field(
        None,
        description="가맹점 주소의 자치구(구 단위). 예: '송파구', '구로구'. "
                    "구 단위가 없으면 시/군 단위(예: '성남시'). 주소가 없으면 null",
    )
    biz_no: Optional[str] = Field(
        None,
        description="사업자등록번호(XXX-XX-XXXXX, 하이픈 포함 10자리). "
                    "'가맹점 정보'/'판매자 정보' 구분이 없는 일반 영수증의 번호. 없으면 null",
    )
    # 배달앱 등 '가맹점 정보'와 '판매자 정보'가 나뉜 영수증 — 각 섹션 값을 '그대로' 읽는다.
    # (어느 값을 쓸지 결정은 코드가 함: resolve_store_biz)
    merchant_name: Optional[str] = Field(
        None, description="'가맹점 정보' 섹션의 상호. 그 섹션이 없으면 null",
    )
    merchant_biz_no: Optional[str] = Field(
        None, description="'가맹점 정보' 섹션의 사업자등록번호(XXX-XX-XXXXX). 없으면 null",
    )
    seller_name: Optional[str] = Field(
        None, description="'판매자 정보' 섹션의 상호. 그 섹션이 없으면 null",
    )
    seller_biz_no: Optional[str] = Field(
        None, description="'판매자 정보' 섹션의 사업자등록번호(XXX-XX-XXXXX). 없으면 null",
    )


def resolve_store_biz(r):
    """영수증 파싱 결과에서 최종 거래처명·사업자등록번호(+짝이 되는 상호)를 결정한다.
    - 거래처명(store): '가맹점 정보' 상호 우선(2섹션) → 일반 상호(store) → '판매자 정보' 상호.
    - 사업자번호(biz_no): '판매자 정보' 번호 우선(2섹션) → 일반 번호(biz_no) → '가맹점 정보' 번호.
    - biz_name: 그 사업자번호와 같은 섹션의 상호(비고에 '상호/번호'로 함께 넣기 위함).
    (모델은 각 섹션 값을 그대로 읽기만 하고, 규칙 적용은 여기서 확정 → 작은 모델도 안정적.)
    반환: (store, biz_no, biz_name)
    """
    store = r.merchant_name or r.store or r.seller_name
    if r.seller_biz_no:
        biz_no, biz_name = r.seller_biz_no, r.seller_name
    elif r.biz_no:
        biz_no, biz_name = r.biz_no, r.store
    elif r.merchant_biz_no:
        biz_no, biz_name = r.merchant_biz_no, r.merchant_name
    else:
        biz_no, biz_name = None, None
    return store, biz_no, biz_name


SYSTEM_PROMPT = (
    "당신은 한국 영수증 OCR 전문가입니다. 주어진 영수증 이미지에서 "
    "영수일자, 상호명(가맹점명), 총 결제금액, 결제시각, 지역(구)을 정확히 추출하세요.\n"
    "\n"
    "[정확도 원칙 — 반드시 준수]\n"
    "- 이미지에 인쇄된 글자를 한 글자 한 글자 그대로 정확히 읽으세요. "
    "임의로 추측하거나 그럴듯하게 보정·교정하지 마세요.\n"
    "- 특히 상호명은 실제로 보이는 글자 그대로 옮기세요. 비슷한 브랜드명으로 "
    "바꾸거나 맞춤법을 임의 수정하지 마세요 (예: 보이는 대로 '스타벅스코리아', "
    "'(주)'·'㈜' 등 표기도 그대로).\n"
    "- 혼동하기 쉬운 글자를 주의 깊게 구분하세요: 숫자 0과 한글 ㅇ·영문 O, "
    "숫자 1과 영문 l·I, 숫자 5와 S, 숫자 8과 B, 한글 '으/응', '리/니' 등.\n"
    "- 금액 숫자는 자릿수와 쉼표 위치를 정확히 확인하여 한 자리도 누락·추가하지 "
    "마세요. 흐릿하면 가장 신뢰도 높은 판독값을 쓰되, 도저히 읽을 수 없으면 null.\n"
    "- 이미지가 회전·기울어져 있거나 일부가 잘려 있어도 보이는 범위에서 최대한 "
    "정확히 판독하세요.\n"
    "\n"
    "[형식 규칙]\n"
    "- 날짜는 YYYY-MM-DD 형식으로 변환 (연도가 두 자리면 20xx로 해석).\n"
    "- 금액은 '합계', '총액', '받을금액', '승인금액' 등 최종 결제금액을 사용하고 "
    "쉼표·원 표기를 제거한 정수로.\n"
    "- 시각은 HH:MM(24시간)으로.\n"
    "- 지역은 가맹점 주소에서 자치구(구 단위, 예: '송파구', '구로구')만 뽑되, "
    "구가 없으면 시/군 단위로. 주소가 없으면 null.\n"
    "\n"
    "[상호·사업자등록번호 — 두 섹션을 각각 '그대로' 읽기]\n"
    "- 배달앱 등 영수증에 '가맹점 정보'와 '판매자 정보' 섹션이 따로 있으면, 각 섹션의 "
    "값을 있는 그대로 채우세요(어느 것을 최종 사용할지는 시스템이 정하므로 그대로만 읽으면 됨):\n"
    "    · merchant_name = '가맹점 정보' 섹션의 상호\n"
    "    · merchant_biz_no = '가맹점 정보' 섹션의 사업자등록번호\n"
    "    · seller_name = '판매자 정보' 섹션의 상호\n"
    "    · seller_biz_no = '판매자 정보' 섹션의 사업자등록번호\n"
    "  그리고 이 경우 store·biz_no 는 null 로 둡니다.\n"
    "  예) 가맹점 정보: 상호 '(주)우아한형제들', 사업자등록번호 120-87-65763 / "
    "판매자 정보: 상호 '명인카츠', 사업자등록번호 581-08-02838 → "
    "merchant_name='(주)우아한형제들', merchant_biz_no='120-87-65763', "
    "seller_name='명인카츠', seller_biz_no='581-08-02838'.\n"
    "- 섹션 구분이 없는 '일반 영수증'이면, 그 상호를 store 에, 사업자등록번호를 biz_no 에 "
    "채우고 merchant_*/seller_* 는 모두 null.\n"
    "- 사업자등록번호는 'XXX-XX-XXXXX'(하이픈 포함 10자리 숫자) 형식. 주문번호·승인번호·"
    "전화번호 등 다른 번호와 혼동하지 마세요.\n"
    "- 정보가 없으면(또는 판독 불가하면) 해당 항목은 null. 추측해서 채우지 마세요."
)


def _image_to_data_url(image_bytes: bytes, filename: str = "receipt.jpg") -> str:
    mime, _ = mimetypes.guess_type(filename)
    if not mime:
        mime = "image/jpeg"
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def _supports_custom_temperature(model: str) -> bool:
    """gpt-5 계열·o1/o3/o4 추론형 모델은 temperature 기본값(1)만 허용한다.
    이런 모델엔 temperature를 아예 보내면 안 되므로(400 에러) False를 돌려준다."""
    m = (model or "").lower().removeprefix("openai:")
    return not m.startswith(("gpt-5", "o1", "o3", "o4"))


def make_llm(model: str = "gpt-4o", api_key: Optional[str] = None, temperature: float = 0.0):
    """구조화 출력이 바인딩된 LLM 생성."""
    kwargs = {"model": model}
    if _supports_custom_temperature(model):
        kwargs["temperature"] = temperature
    if api_key:
        kwargs["api_key"] = api_key
    llm = ChatOpenAI(**kwargs)
    return llm.with_structured_output(Receipt)


def parse_receipt(image_bytes: bytes, filename: str = "receipt.jpg",
                  llm=None, model: str = "gpt-4o",
                  api_key: Optional[str] = None) -> Receipt:
    """영수증 이미지 1장을 파싱해 Receipt 반환."""
    if llm is None:
        llm = make_llm(model=model, api_key=api_key)

    data_url = _image_to_data_url(image_bytes, filename)
    messages = [
        SystemMessage(content=SYSTEM_PROMPT),
        HumanMessage(content=[
            {"type": "text", "text": "이 영수증 이미지를 자세히 보고, 인쇄된 "
             "글자를 오타 없이 한 글자도 틀리지 않게 그대로 읽어 정보를 추출해줘. "
             "특히 상호명과 금액 숫자를 정확히."},
            {"type": "image_url", "image_url": {"url": data_url}},
        ]),
    ]
    return llm.invoke(messages)


def parse_many(images: List[tuple], model: str = "gpt-4o",
               api_key: Optional[str] = None, on_progress=None) -> List[dict]:
    """
    images: list[(filename, bytes)]
    반환  : list[dict]  키 = filename, date, store, amount, time, error
    """
    llm = make_llm(model=model, api_key=api_key)
    results = []
    for idx, (fname, content) in enumerate(images):
        rec = {"filename": fname, "date": None, "store": None,
               "amount": None, "time": None, "region": None, "biz_no": None,
               "biz_name": None, "error": None}
        try:
            r = parse_receipt(content, fname, llm=llm)
            store, biz_no, biz_name = resolve_store_biz(r)
            rec.update(date=r.date, store=store, amount=r.amount,
                       time=r.time, region=r.region, biz_no=biz_no, biz_name=biz_name)
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        results.append(rec)
        if on_progress:
            on_progress(idx + 1, len(images), rec)
    return results
