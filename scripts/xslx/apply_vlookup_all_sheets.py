import uno

def set_vlookup(sheet, row):
    """지정 시트의 지정 행에 VLOOKUP 수식 입력"""
    cell = sheet.getCellByPosition(1, row)  # B열
    cell.FormulaLocal = f"=VLOOKUP(A{row + 1}; $mapping.$A:$B; 2; 0)"

def apply_vlookup_current_sheet():
    """현재 시트만 처리"""
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet
    _apply_vlookup_to_sheet(sheet)

def apply_vlookup_all_sheets():
    """모든 시트 처리 (mapping 시트 제외)"""
    doc = XSCRIPTCONTEXT.getDocument()
    for sheet in doc.Sheets:
        if sheet.Name.lower() == "mapping":
            continue
        _apply_vlookup_to_sheet(sheet)

def _apply_vlookup_to_sheet(sheet):
    """공통 처리 로직"""
    cursor = sheet.createCursorByRange(sheet.getCellByPosition(1, 7))  # B8 기준
    cursor.gotoEndOfUsedArea(True)
    last_row = cursor.RangeAddress.EndRow

    for row in range(7, last_row + 1):
        set_vlookup(sheet, row)
