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
        None, description="상호명(가맹점/거래처명). 없으면 null"
    )
    amount: Optional[int] = Field(
        None, description="총 결제(영수) 금액. 숫자만(원 단위 정수). 없으면 null"
    )
    time: Optional[str] = Field(
        None, description="결제 시각. HH:MM(24시간) 형식. 없으면 null"
    )


SYSTEM_PROMPT = (
    "당신은 한국 영수증 OCR 전문가입니다. 주어진 영수증 이미지에서 "
    "영수일자, 상호명(가맹점명), 총 결제금액, 결제시각을 정확히 추출하세요.\n"
    "규칙:\n"
    "- 날짜는 YYYY-MM-DD 형식으로 변환 (연도가 두 자리면 20xx로 해석).\n"
    "- 상호명은 사업자명/가맹점명을 우선하고, 지점명이 있으면 함께 포함.\n"
    "- 금액은 '합계', '총액', '받을금액', '승인금액' 등 최종 결제금액을 사용하고 "
    "쉼표·원 표기를 제거한 정수로.\n"
    "- 시각은 HH:MM(24시간)으로. 정보가 없으면 해당 항목은 null."
)


def _image_to_data_url(image_bytes: bytes, filename: str = "receipt.jpg") -> str:
    mime, _ = mimetypes.guess_type(filename)
    if not mime:
        mime = "image/jpeg"
    b64 = base64.b64encode(image_bytes).decode("utf-8")
    return f"data:{mime};base64,{b64}"


def make_llm(model: str = "gpt-4o", api_key: Optional[str] = None, temperature: float = 0.0):
    """구조화 출력이 바인딩된 LLM 생성."""
    kwargs = {"model": model, "temperature": temperature}
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
            {"type": "text", "text": "이 영수증에서 정보를 추출해줘."},
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
               "amount": None, "time": None, "error": None}
        try:
            r = parse_receipt(content, fname, llm=llm)
            rec.update(date=r.date, store=r.store, amount=r.amount, time=r.time)
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        results.append(rec)
        if on_progress:
            on_progress(idx + 1, len(images), rec)
    return results
