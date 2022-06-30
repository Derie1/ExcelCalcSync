from PyQt5 import QtCore, QtWidgets
import openpyxl
import os.path as path

app = QtWidgets.QApplication([])
target_file_list = QtWidgets.QFileDialog.getOpenFileNames(
    caption="Выберите файлы Excel в которых надо внести изменения... ", filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем целевые файлы

busbar_feeded = []


def is_busbar_feeded(db_xlsx):
    wbook = openpyxl.load_workbook(db_xlsx)
    wsheet = wbook['Расчет']
    cb_type = str(wsheet['W16'].value)
    return cb_type.startswith('QFшп')


for file in target_file_list:
    if is_busbar_feeded(file):
        busbar_feeded.append(path.split(file)[1])


print(busbar_feeded)
