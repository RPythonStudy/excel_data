#!/bin/bash
# LibreOffice UNO 설치 확인 스크립트

echo "=== LibreOffice UNO 설치 확인 ==="

# 1. 설치된 패키지 확인
echo "1. 설치된 패키지 확인:"
dpkg -l | grep -E "(libreoffice|python3-uno)" | awk '{print $2 " " $3}'

echo ""

# 2. LibreOffice 버전 확인
echo "2. LibreOffice 버전:"
libreoffice --version 2>/dev/null || echo "LibreOffice 미설치"

echo ""

# 3. Python UNO 모듈 확인
echo "3. Python UNO 모듈 확인:"
python3 -c "
import sys
sys.path.append('/usr/lib/python3/dist-packages')
try:
    import uno
    print('✓ UNO 모듈 import 성공')
except ImportError as e:
    print(f'✗ UNO 모듈 import 실패: {e}')
"

echo ""

# 4. UNO 관련 파일 위치 확인
echo "4. UNO 관련 파일 위치:"
echo "Python UNO 모듈:"
find /usr/lib/python* -name "uno.py" 2>/dev/null | head -3

echo "LibreOffice 프로그램:"
ls -la /usr/lib/libreoffice/program/ 2>/dev/null | head -5

echo ""
echo "=== 설치 완료 확인 ==="