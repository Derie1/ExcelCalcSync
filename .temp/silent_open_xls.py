import xlwings as xw

# option 1
"""
app = xw.App(visible=False)
book = xw.Book('PATH_TO_YOUR_XLSX_FILE')
sheet = book.sheets('sheetname')
df = sheet.range('A1:D10')
book.close()
app.quit()
"""

#option 2
"""
import xlwings as xw

excel_app = xw.App(visible=False)
excel_book = excel_app.books.open('PATH_TO_YOUR_XLSX_FILE')
excel_book.save()
excel_book.close()
excel_app.quit()
"""