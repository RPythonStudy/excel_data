import sys
import glob
import os
from pathlib import Path

# LibreOffice UNO 경로 추가
sys.path.append('/usr/lib/python3/dist-packages')
sys.path.append('/usr/lib/libreoffice/program')

# 프로젝트 경로 추가
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(script_dir, "..", "..")
sys.path.append(os.path.abspath(project_root))

try:
    import uno
    from com.sun.star.beans import PropertyValue
    from src.common.logger import log_info, log_error
    UNO_AVAILABLE = True
except ImportError as e:
    print(f"UNO import 실패: {e}")
    print("LibreOffice 설치 필요: sudo apt install libreoffice python3-uno")
    UNO_AVAILABLE = False


def get_excel_file():
    """데이터 폴더에서 Excel 파일 자동 검색"""
    # 현재 스크립트 위치에서 프로젝트 루트로 이동
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.join(script_dir, "..", "..")
    data_dir = os.path.join(project_root, "data", "deid", "raw")
    
    # 절대 경로로 변환
    data_dir = os.path.abspath(data_dir)
    
    excel_files = glob.glob(f"{data_dir}/*.xlsx") + glob.glob(f"{data_dir}/*.xlsm")
    
    if excel_files:
        return excel_files[0]  # 첫 번째 파일 반환
    else:
        raise FileNotFoundError(f"Excel 파일을 찾을 수 없습니다. 검색 경로: {data_dir}")

def list_sheets(xlsx_path):
    """LibreOffice UNO를 사용하여 Excel 파일의 시트 목록 가져오기"""
    if not UNO_AVAILABLE:
        raise ImportError("LibreOffice UNO가 설치되지 않았습니다.")
    
    try:
        # LibreOffice 실행 컨텍스트 획득
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        # 파일 경로를 UNO URL로 변환
        file_url = uno.systemPathToFileUrl(xlsx_path)
        log_info(f"파일 URL: {file_url}")

        # 엑셀 문서 열기 (헤드리스 모드)
        props = [PropertyValue("Hidden", 0, True, 0)]
        doc = desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))

        # 시트 이름 목록 추출
        sheet_names = [doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.getCount())]
        log_info(f"발견된 시트: {sheet_names}")
        print("시트 목록:", sheet_names)

        # 문서 닫기
        doc.close(True)
        return sheet_names
        
    except Exception as e:
        log_error(f"UNO Excel 처리 오류: {e}")
        raise

if __name__ == "__main__":
    try:
        # 방법 1: 자동으로 Excel 파일 찾기
        excel_file = get_excel_file()
        print(f"발견된 Excel 파일: {excel_file}")
        
        # 방법 2: 직접 경로 지정 (예시)
        xexcel_file = "/home/ben/projects/excel_data/data/deid/raw/KQIPS eCRF (수신-전산팀) 20250508 수정_수술전후검사결과제공_20250529.xls"
        
        # 시트 목록 출력
        sheets = list_sheets(excel_file)
        print(f"총 {len(sheets)}개의 시트가 있습니다.")
        
    except FileNotFoundError as e:
        print(f"오류: {e}")
    except Exception as e:
        print(f"예상치 못한 오류: {e}")
