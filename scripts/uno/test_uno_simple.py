#!/usr/bin/env python3
# 시스템 Python으로 실행하는 간단한 UNO 테스트

import sys
import os

# UNO 경로 추가
sys.path.append('/usr/lib/python3/dist-packages')
sys.path.append('/usr/lib/libreoffice/program')

def test_uno_basic():
    """기본 UNO 기능 테스트"""
    try:
        import uno
        print("✓ UNO 모듈 import 성공")
        
        # 컨텍스트 생성
        ctx = uno.getComponentContext()
        print("✓ UNO 컨텍스트 생성 성공")
        
        # ServiceManager 생성
        smgr = ctx.ServiceManager
        print("✓ ServiceManager 생성 성공")
        
        return True
        
    except Exception as e:
        print(f"✗ UNO 테스트 실패: {e}")
        return False

def simple_excel_test():
    """간단한 Excel 파일 테스트"""
    try:
        import uno
        from com.sun.star.beans import PropertyValue
        
        # 테스트할 Excel 파일 경로
        excel_path = "/home/ben/projects/excel_data/data/deid/raw/KQIPS eCRF (수신-전산팀) 20250508 수정_수술전후검사결과제공_20250529.xlsm"
        
        if not os.path.exists(excel_path):
            print(f"✗ Excel 파일이 존재하지 않습니다: {excel_path}")
            return False
        
        # LibreOffice 시작
        ctx = uno.getComponentContext()
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        
        # 파일 URL 변환
        file_url = uno.systemPathToFileUrl(excel_path)
        print(f"파일 URL: {file_url}")
        
        # 문서 열기
        props = [PropertyValue("Hidden", 0, True, 0)]
        doc = desktop.loadComponentFromURL(file_url, "_blank", 0, tuple(props))
        
        # 시트 정보 가져오기
        sheet_count = doc.Sheets.getCount()
        print(f"총 시트 수: {sheet_count}")
        
        # 시트 이름 출력
        for i in range(sheet_count):
            sheet_name = doc.Sheets.getByIndex(i).Name
            print(f"시트 {i+1}: {sheet_name}")
        
        # 문서 닫기
        doc.close(True)
        print("✓ Excel 파일 처리 성공")
        return True
        
    except Exception as e:
        print(f"✗ Excel 처리 실패: {e}")
        return False

if __name__ == "__main__":
    print("=== LibreOffice UNO 테스트 ===")
    
    print("\n1. 기본 UNO 테스트:")
    if test_uno_basic():
        print("\n2. Excel 파일 테스트:")
        simple_excel_test()
    
    print("\n=== 테스트 완료 ===")