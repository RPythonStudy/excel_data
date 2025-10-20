import uno

def apply_vlookup_current_cell():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet
    addr = doc.CurrentSelection.RangeAddress

    row = addr.StartRow + 1
    cell = sheet.getCellByPosition(addr.StartColumn, addr.StartRow)

    # 한글 UI 기준: 세미콜론(;) 구분자 사용
    cell.FormulaLocal = f"=VLOOKUP(A{row}; $mapping.$A:$B; 2; 0)"

