import xlwings as xw

EXCEL_FILE = 'AssignmentforDataScientist.xlsx'

try:
    excel_app = xw.App(visible=False)
    wb = excel_app.books.open(EXCEL_FILE)
    for sheet in wb.sheets:
        sheet.api.Copy()
        wb_new = xw.books.active
        wb_new.save(f'{sheet.name}.xlsx')
        wb_new.close()

finally:
    excel_app.quite()