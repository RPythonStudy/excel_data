#!/bin/bash
# LibreOffice UNO 환경 설정 스크립트

echo "=== LibreOffice UNO 환경 설정 ==="

# 1. LibreOffice 설치 확인
echo "1. LibreOffice 설치 확인..."
if command -v libreoffice &> /dev/null; then
    echo "✓ LibreOffice 설치됨: $(libreoffice --version)"
else
    echo "✗ LibreOffice 미설치"
    echo "설치 명령: sudo apt install libreoffice libreoffice-dev python3-uno"
    exit 1
fi

# 2. Python UNO 모듈 경로 확인
echo "2. Python UNO 모듈 경로 확인..."
PYTHON_UNO_PATH="/usr/lib/python3/dist-packages"
if [ -d "$PYTHON_UNO_PATH/uno.py" ] || [ -d "$PYTHON_UNO_PATH/uno" ]; then
    echo "✓ Python UNO 모듈 발견: $PYTHON_UNO_PATH"
else
    echo "✗ Python UNO 모듈 미발견"
    echo "대체 경로들 확인 중..."
    find /usr -name "uno.py" 2>/dev/null | head -5
fi

# 3. LibreOffice 프로그램 경로 확인
echo "3. LibreOffice 프로그램 경로 확인..."
LO_PROGRAM_PATH="/usr/lib/libreoffice/program"
if [ -d "$LO_PROGRAM_PATH" ]; then
    echo "✓ LibreOffice 프로그램 경로: $LO_PROGRAM_PATH"
else
    echo "✗ LibreOffice 프로그램 경로 미발견"
fi

# 4. 환경 변수 설정 가이드
echo "4. 환경 변수 설정 가이드:"
echo "다음 명령을 ~/.bashrc에 추가하세요:"
echo ""
echo "export PYTHONPATH=\$PYTHONPATH:/usr/lib/python3/dist-packages"
echo "export PYTHONPATH=\$PYTHONPATH:/usr/lib/libreoffice/program"
echo "export URE_BOOTSTRAP=\"vnd.sun.star.pathname:/usr/lib/libreoffice/program/fundamentalrc\""
echo ""

# 5. 테스트 스크립트 생성
echo "5. 테스트 스크립트 생성..."
cat > test_uno.py << 'EOF'
#!/usr/bin/env python3
import sys
import os

# UNO 경로 추가
sys.path.append('/usr/lib/python3/dist-packages')
sys.path.append('/usr/lib/libreoffice/program')

try:
    import uno
    print("✓ UNO 모듈 import 성공")
    
    # UNO 컨텍스트 테스트
    ctx = uno.getComponentContext()
    print("✓ UNO 컨텍스트 생성 성공")
    
    smgr = ctx.ServiceManager
    print("✓ ServiceManager 생성 성공")
    
    print("🎉 LibreOffice UNO 설정 완료!")
    
except ImportError as e:
    print(f"✗ UNO import 실패: {e}")
    print("LibreOffice 및 python3-uno 패키지를 설치하세요.")
except Exception as e:
    print(f"✗ UNO 초기화 실패: {e}")
EOF

chmod +x test_uno.py
echo "테스트 스크립트 생성됨: test_uno.py"
echo "실행: python3 test_uno.py"