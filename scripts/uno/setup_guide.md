# LibreOffice UNO 설정 가이드

## 1. 현재 잘못된 uno 패키지 제거
pip uninstall uno -y

## 2. LibreOffice 설치 확인
sudo apt update
sudo apt install libreoffice libreoffice-dev

## 3. LibreOffice Python 환경 설정
# LibreOffice에 내장된 Python 사용하거나
# 시스템 Python에서 LibreOffice UNO 모듈 경로 추가

## 4. 환경 변수 설정
export PYTHONPATH=$PYTHONPATH:/usr/lib/libreoffice/program
export URE_BOOTSTRAP="vnd.sun.star.pathname:/usr/lib/libreoffice/program/fundamentalrc"

## 5. 대안: 별도 스크립트로 LibreOffice 실행
# libreoffice --headless --calc --convert-to csv filename.xlsx

## 권장사항: pandas + openpyxl 사용
# 더 간단하고 안정적인 방법입니다.