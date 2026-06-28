#!/usr/bin/env bash
# 청구서 앱 종료 (포트 8000)
# 사용법:  ./stop.sh
PORT=8000

PIDS=$(lsof -ti:"$PORT" 2>/dev/null)
if [ -z "$PIDS" ]; then
  echo "실행 중인 앱이 없어요 (포트 $PORT 비어 있음)."
  exit 0
fi

echo "포트 $PORT 종료 중... (PID: $PIDS)"
kill $PIDS 2>/dev/null
sleep 1

# 아직 살아 있으면 강제 종료
PIDS=$(lsof -ti:"$PORT" 2>/dev/null)
if [ -n "$PIDS" ]; then
  kill -9 $PIDS 2>/dev/null
fi
echo "✅ 종료 완료."
