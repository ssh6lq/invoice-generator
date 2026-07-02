# 영수증 → 비용청구서 앱 (FastAPI) 컨테이너 이미지
FROM python:3.11-slim

# 파이썬 로그가 버퍼링 없이 바로 나오도록
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# 1) 의존성 먼저 설치 (소스보다 먼저 복사해야 캐시가 살아남음)
#    서버 전용 최소 의존성만 설치 → 이미지 슬림 (pandas/pillow/streamlit 제외)
COPY requirements-server.txt .
RUN pip install --upgrade pip && pip install -r requirements-server.txt

# 2) 앱 소스 + 번들 양식 + 정적 파일 복사
COPY . .

# 3) 앱 포트
EXPOSE 8000

# 4) 실행 (컨테이너 안에서는 venv 없이 시스템 파이썬으로 실행)
#    --log-config: 로그에 날짜/시간(타임스탬프)을 찍기 위한 설정
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000", \
     "--log-config", "log_config.json"]
