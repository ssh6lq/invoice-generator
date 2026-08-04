"""log_tz.py — 로그 타임스탬프를 항상 한국시간(KST)으로 찍기 위한 포매터.

컨테이너는 Dockerfile/compose 의 TZ=Asia/Seoul 로 이미 KST 지만, 로컬에서
./start.sh 나 uvicorn 을 직접 띄우면 그 PC의 시간대를 따라간다.
로그 시각만큼은 실행 환경과 무관하게 KST 로 고정해, 사내 문의 대응 시
"몇 시에 생긴 오류인지"가 어긋나지 않게 한다.

log_config.json 의 formatters 에서 "()" 로 이 클래스들을 지정해 쓴다.
"""

from datetime import datetime, timezone, timedelta

from uvicorn.logging import AccessFormatter, DefaultFormatter

KST = timezone(timedelta(hours=9), "KST")


class _KSTMixin:
    """logging 이 기본으로 쓰는 time.localtime 대신 KST 로 변환한다."""

    def formatTime(self, record, datefmt=None):  # noqa: N802 (logging 규약)
        dt = datetime.fromtimestamp(record.created, KST)
        return dt.strftime(datefmt) if datefmt else dt.isoformat()


class KSTDefaultFormatter(_KSTMixin, DefaultFormatter):
    """일반 로그(uvicorn.error 포함)용."""


class KSTAccessFormatter(_KSTMixin, AccessFormatter):
    """접속 로그(uvicorn.access)용."""
