def set_vlookup(sheet, row):
    cell = sheet.getCellByPosition(1, row)
    cell.FormulaLocal = f"=VLOOKUP(A{row + 1}; $mapping.$A:$B; 2; 0)"

def apply_vlookup_current_sheet():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet
    cursor = sheet.createCursorByRange(sheet.getCellByPosition(1, 7))
    cursor.gotoEndOfUsedArea(True)
    last_row = cursor.RangeAddress.EndRow

    for row in range(7, last_row + 1):
        set_vlookup(sheet, row)
