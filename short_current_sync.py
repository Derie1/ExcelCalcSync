# Синхронизирует участки для расчетов токов КЗ из выбранного файла в другие файлы по выбору пользователя

import xlwings as xw
from PyQt5 import QtCore, QtWidgets


# =============== def's put here
def data_to_transfer(sheet_source):
    transferred_data = {
        'B6': sheet_source.range('B6').value,
        'E6': sheet_source.range('E6').value,
        'F6': sheet_source.range('F6').value,
        'E7': sheet_source.range('E7').value,
        'F7': sheet_source.range('F7').value,
        'K6': sheet_source.range('K6').value,
        'L6': sheet_source.range('L6').value,
        'K7': sheet_source.range('K7').value,
        'L7': sheet_source.range('L7').value,
        'Q6': sheet_source.range('Q6').value,
        'R6': sheet_source.range('R6').value,
        'W6': sheet_source.range('W6').formula,
        'X6': sheet_source.range('X6').formula
    }
    return transferred_data


# =============== main code
app = QtWidgets.QApplication([])

source_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                    filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл

wb_source = xw.Book(source_file)
sheet_source = wb_source.sheets['Расчет КЗ']
t_data = data_to_transfer(sheet_source)

target_file_list = QtWidgets.QFileDialog.getOpenFileNames(
    caption="Выберите файлы Excel в которых надо внести изменения... ", filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем целевые файлы
for file in target_file_list:
    with xw.App(visible=False) as xl_app:
        wb = xw.Book(file)
        current_sheet = wb.sheets['Расчет КЗ']
        current_sheet.range('B6').value = t_data['B6']
        current_sheet.range('E6').value = t_data['E6']
        current_sheet.range('F6').value = t_data['F6']
        current_sheet.range('E7').value = t_data['E7']
        current_sheet.range('F7').value = t_data['F7']
        current_sheet.range('K6').value = t_data['K6']
        current_sheet.range('L6').value = t_data['L6']
        current_sheet.range('K7').value = t_data['K7']
        current_sheet.range('L7').value = t_data['L7']
        current_sheet.range('Q6').value = t_data['Q6']
        current_sheet.range('R6').value = t_data['R6']
        current_sheet.range('W6').value = t_data['W6']
        current_sheet.range('X6').value = t_data['X6']
        wb.save()
        wb.close()
