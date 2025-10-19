#!/bin/bash
# LibreOffice UNO 서버 시작 스크립트

echo "=== LibreOffice UNO 서버 시작 ==="

# 1. 기존 LibreOffice 프로세스 종료
echo "1. 기존 LibreOffice 프로세스 확인 및 종료..."
pkill -f "soffice" 2>/dev/null || true
sleep 2

# 2. LibreOffice 헤드리스 서버 시작
echo "2. LibreOffice 헤드리스 서버 시작..."
libreoffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;StarOffice.ServiceManager" &
LIBREOFFICE_PID=$!

echo "LibreOffice 서버 PID: $LIBREOFFICE_PID"
echo "서버 시작 대기 중..."
sleep 5

# 3. 서버 상태 확인
echo "3. 서버 상태 확인..."
if ps -p $LIBREOFFICE_PID > /dev/null; then
    echo "✓ LibreOffice 서버 실행 중"
    echo "포트 2002에서 UNO 연결 대기 중"
else
    echo "✗ LibreOffice 서버 시작 실패"
    exit 1
fi

echo ""
echo "=== 사용법 ==="
echo "다른 터미널에서 다음 명령 실행:"
echo "python scripts/uno/read_sheet_server.py"
echo ""
echo "종료하려면 Ctrl+C 누르세요"

# 4. 서버 유지
wait $LIBREOFFICE_PID