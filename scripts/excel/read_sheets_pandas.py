import pandas as pd
import glob
import os
import sys
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.join(script_dir, "..", "..")
sys.path.append(os.path.abspath(project_root))

from src.common.logger import log_info, log_error


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


def list_sheets_pandas(xlsx_path):
    """pandas를 사용하여 Excel 파일의 시트 목록 가져오기"""
    try:
        # Excel 파일의 모든 시트 이름 가져오기
        xl_file = pd.ExcelFile(xlsx_path)
        sheet_names = xl_file.sheet_names
        
        log_info(f"시트 목록: {sheet_names}")
        print("시트 목록:", sheet_names)
        
        # 각 시트의 기본 정보 출력
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(xlsx_path, sheet_name=sheet_name, nrows=0)  # 헤더만 읽기
                print(f"  - {sheet_name}: {len(df.columns)}개 컬럼")
            except Exception as e:
                print(f"  - {sheet_name}: 읽기 오류 ({str(e)[:50]}...)")
        
        return sheet_names
        
    except Exception as e:
        log_error(f"Excel 파일 읽기 오류: {e}")
        raise


def read_sheet_data(xlsx_path, sheet_name, nrows=5):
    """특정 시트의 데이터 미리보기"""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name, nrows=nrows)
        print(f"\n=== {sheet_name} 시트 미리보기 (최대 {nrows}행) ===")
        print(df)
        print(f"전체 모양: {df.shape}")
        return df
        
    except Exception as e:
        log_error(f"시트 '{sheet_name}' 읽기 오류: {e}")
        return None


if __name__ == "__main__":
    try:
        # 방법 1: 자동으로 Excel 파일 찾기
        excel_file = get_excel_file()
        print(f"발견된 Excel 파일: {excel_file}")
        
        # 시트 목록 출력
        sheets = list_sheets_pandas(excel_file)
        print(f"\n총 {len(sheets)}개의 시트가 있습니다.")
        
        # 첫 번째 시트 미리보기
        if sheets:
            first_sheet = sheets[0]
            read_sheet_data(excel_file, first_sheet, nrows=3)
        
    except FileNotFoundError as e:
        print(f"오류: {e}")
    except Exception as e:
        print(f"예상치 못한 오류: {e}")