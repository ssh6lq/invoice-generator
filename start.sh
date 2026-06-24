#!/usr/bin/env bash
# 청구서 앱 실행 (같은 네트워크 공유 모드)
# 사용법:  ./start.sh
cd "$(dirname "$0")"

PORT=8501

# 이미 떠 있으면 중복 실행 방지
if lsof -ti:"$PORT" >/dev/null 2>&1; then
  echo "⚠️  포트 $PORT 에 이미 앱이 실행 중이에요. 끄려면 ./stop.sh"
  exit 1
fi

# 같은 네트워크에서 접속할 LAN IP 찾기
IFACE=$(route get default 2>/dev/null | awk '/interface:/{print $2}')
IP=$(ipconfig getifaddr "$IFACE" 2>/dev/null)
[ -z "$IP" ] && IP=$(ipconfig getifaddr en0 2>/dev/null)

echo "▶ 앱을 실행합니다."
echo "   같은 와이파이/사무실 네트워크에서 접속 주소:"
echo "     http://${IP:-<내IP>}:$PORT"
echo "   종료하려면: 이 창에서 Ctrl+C  또는  다른 창에서 ./stop.sh"
echo

exec .venv/bin/streamlit run app.py \
  --server.address 0.0.0.0 \
  --server.port "$PORT" \
  --browser.gatherUsageStats false
