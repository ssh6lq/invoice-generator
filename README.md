# 영수증 → 비용청구서 자동 작성

영수증 사진 여러 장을 올리면 GPT 비전(LangChain, OpenAI 호환 API)으로 **영수일자·거래처명·영수금액·영수시간**을 추출해서 비용청구 양식에 자동으로 채워주는 사내 웹앱입니다. 초과근무(수당) 신청서 자동 작성, 문의/이슈 접수, 청구서 제출내역 관리 기능도 함께 제공합니다.

현재 서비스는 **FastAPI(`server.py`) 기반**이며, 사내망에서 여러 사용자가 브라우저로 동시에 접속해 쓰는 무상태(stateless) 서버입니다. (`app.py`는 초기 Streamlit 프로토타입으로 더 이상 운영에 쓰지 않습니다.)

## 구성

| 파일 | 역할 |
|------|------|
| `server.py` | FastAPI 서버 — API 라우팅, 정적 페이지 서빙, 관리자 인증 |
| `receipt_parser.py` | LangChain + GPT 비전으로 영수증 1장 파싱 |
| `excel_filler.py` | 비용청구 양식 채우기·검토(정렬/매핑/한도 검증)·매크로검토 대체 로직 |
| `overtime_filler.py` | 근태현황 파싱 및 초과근무신청서 채우기 |
| `feedback_store.py` | 문의/이슈 SQLite 저장소 (`feedback.db`) |
| `submission_store.py` | 청구서 제출내역 SQLite 저장소 (`submissions.db`) — 제출 시 매크로검토 최종 표도 함께 보관 |
| `static/` | 프론트엔드 (index.html·app.js = 사용자 화면, admin_*.html/js = 관리자 화면) |
| `비용청구양식.xlsm` / `초과근무(수당)신청서_양식.xlsx` | 기본 양식 파일 |
| `requirements-server.txt` | `server.py` 실행에 필요한 최소 의존성 (Docker 이미지용) |
| `requirements.txt` | 구버전 Streamlit(`app.py`)용 의존성 |

## 설치

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows (macOS/Linux: source .venv/bin/activate)
pip install -r requirements-server.txt
```

## 환경변수 (.env)

프로젝트 루트에 `.env` 파일을 만들어 설정합니다.

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `ADMIN_PASSWORD` | (없음, 필수) | 관리자 페이지 접근용 HTTP Basic 비밀번호. 미설정 시 관리자 API가 500 에러를 반환합니다. |
| `ADMIN_USER` | `admin` | 관리자 페이지 로그인 아이디 |
| `RECEIPT_PROVIDER` | `로컬 서버` | 영수증 파싱 LLM 공급자. `로컬 서버`(사내 OpenAI 호환 엔드포인트) 또는 그 외(OpenAI 직접 호출) |
| `RECEIPT_BASE_URL` | `http://192.168.1.51:8000` | 로컬 서버 모드일 때 OpenAI 호환 엔드포인트 주소 |
| `RECEIPT_MODEL` | `vllm-qwen` | 사용할 모델명 |
| `RECEIPT_API_KEY` / `RECEIPT_BEARER` / `OPENAI_API_KEY` | (없음) | 인증 토큰. 이 순서로 먼저 설정된 값을 사용 |

## 서버 실행

```bash
.venv\Scripts\activate           # 활성화 안 돼 있으면
uvicorn server:app --host 0.0.0.0 --port 8000
```

실행하면 `http://localhost:8000` 에서 앱이 열립니다. 사내망의 다른 PC에서는 `http://<이 PC IP>:8000` 으로 접속합니다.

## 서버 종료

- **포그라운드 실행(터미널이 앱에 묶인 경우)**: 해당 터미널에서 `Ctrl + C`
- **백그라운드로 실행한 경우**: 포트로 프로세스를 찾아 종료

  ```bash
  # Windows (PowerShell)
  Get-NetTCPConnection -LocalPort 8000 | Select-Object -Expand OwningProcess | Stop-Process

  # macOS / Linux
  lsof -ti:8000 | xargs kill
  ```

작업이 끝나면 `deactivate` 로 가상환경을 빠져나올 수 있습니다.

## 화면 구성

### 사용자 화면 (`/`)

1. **비용청구서**: 영수증 사진 업로드 → 🔍 영수증 분석 → 표에서 검토·수정 → 매크로 검토(정렬·매핑·한도 검증) → 제출용 청구서(.xlsx) 또는 원본 양식(.xlsm) 다운로드
2. **초과근무(수당) 신청서**: 근태현황(.xlsx) 업로드 → 초과근무 대상일 확인·수정 → 신청서(.xlsx) 다운로드

'제출용 청구서 다운로드'를 누르면 그 시점의 검토 완료 표가 자동으로 관리자 페이지의 제출내역에 기록됩니다.

### 관리자 화면 (HTTP Basic 인증 필요)

| 경로 | 내용 |
|------|------|
| `/admin_submissions.html` | 청구서 제출내역 — 부서/이름/건수/총액 확인, 접수 상태(접수·검토중·승인·반려·지급완료) 변경, 제출 시점의 매크로검토 최종 표(영수일자·거래처명·목적·금액·청구금액 등) 펼쳐보기 |
| `/admin_feedback.html` | 문의/이슈 접수 내역 확인·상태 변경 |

## 채워지는 열 (비용청구서 작성시트)

| 열 | 항목 | 자동 입력 |
|----|------|-----------|
| C | 영수일자 | ✅ 영수증 인식 |
| D | 거래처명 | ✅ 영수증 인식 |
| E | 목적 | ✅ 표에서 드롭다운 선택 |
| F | 영수금액 | ✅ 영수증 인식 |
| G | 결제방식 | ✅ 표에서 드롭다운 선택 |
| J | 영수시간 | ✅ 영수증 인식 (인식 시) |
| B | 순번 | 양식 수식 그대로 |

목적·결제방식은 영수증만으로 판단이 어려워 자동 인식 대신, 업로드한 양식에 들어있는 **드롭다운 목록 그대로** 표에서 선택하게 했습니다. 목록은 양식에서 자동 추출됩니다(목적은 `비용지원안내` 시트, 결제방식은 4종 고정 목록).

## 참고

- 비전 입력을 지원하는 모델만 사용 가능합니다.
- 분석 결과는 항상 표에서 검토·수정한 뒤 매크로 검토(한도·형식 검증)를 통과해야 다운로드할 수 있습니다.
- 제출용 다운로드(.xlsx)는 완성된 독립 비용청구서를, 양식 다운로드(.xlsm)는 매크로가 포함된 원본 양식(작성시트만 채움)을 내려줍니다.

## 도형/버튼이 사라지지 않게 한 방법

엑셀 라이브러리(openpyxl)로 파일을 다시 저장하면 버튼 그림·도형(텍스트박스/그룹)·VML 같은 그리기 객체가 사라집니다. 그래서 `excel_filler.py`는 파일을 다시 저장하지 않고, `.xlsm`을 zip으로 열어 **작성시트 XML의 대상 셀(C/D/F/J) 값만 직접 교체**합니다. 나머지(도형·이미지·VBA·서식·수식)는 원본 그대로 복사되므로 100% 보존됩니다. 검증 결과 원본 70개 내부 파일이 모두 유지되고 `작성시트`만 변경됩니다.
