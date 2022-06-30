from encodings.utf_8 import encode
from PyQt5 import QtCore, QtWidgets
import openpyxl
import os.path as path
import io
import csv


def get_name_and_short_current(db_xlsx):
    wbook = openpyxl.load_workbook(db_xlsx, data_only=True)
    wsheet = wbook['Расчет']
    db_name = wsheet['D15'].value
    short_current = wsheet['N16'].value
    return [db_name, short_current]


app = QtWidgets.QApplication([])
target_file_list = QtWidgets.QFileDialog.getOpenFileNames(
    caption="Выберите файлы Excel в которых надо внести изменения... ",
    filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем целевые файлы

current_path = path.dirname(target_file_list[0])

short_current_list = []
for file in target_file_list:
    panel_sc = get_name_and_short_current(file)
    short_current_list.append(panel_sc)


sc_log_file = f"{current_path}/short_currents.csv"

with open(sc_log_file, 'w', encoding='utf8', newline='') as csv_file:
    csv_writer = csv.writer(csv_file)
    csv_writer.writerows(short_current_list)
    # for _ in short_current_list:
    #     csv_writer.writerow([_[0], _[1]])


print("Job's DONE!!")
