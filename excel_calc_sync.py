import os
import pandas as pd
import xlwings as xw
from PyQt5 import QtCore, QtWidgets


def data_to_transfer(source_range, row_index): #row_index = panels_with_excel[panel][0]
    transferred_data = {
        'CB_def' :  (source_range[row_index, 38].value + source_range[row_index, 0].value),
        'CB_type' : source_range[row_index, 39].value,
        'CB_In' : source_range[row_index, 42].value,
        'CB_ext_string' : source_range[row_index, 44].value,
        'cable_mark' : source_range[row_index, 32].value,
        'cable_section' : source_range[row_index, 34].value,
        'cable_lenght' : source_range[row_index, 35].value
        }
    return transferred_data

def get_DB_range(file): #file = panels_with_excel[panel][1]
    wb_db = xw.Book(file)
    db_sheet = wb_db.sheets['Расчет']
    db_range = db_sheet.range('transferred_data')
    return db_range

app = QtWidgets.QApplication([])

mdb_file = QtWidgets.QFileDialog.getOpenFileName()[0] # выбираем файл расчетки ГРЩ
mdb_file_dir = os.path.dirname(mdb_file)
# print(mdb_file)

wb_mdb = xw.Book(mdb_file)
calc_sheet = wb_mdb.sheets['Расчет']
calc_range = calc_sheet.range('Calculation') # берём диапазон расчета из файла расчетки ГРЩ
# print(calc_range)

# создаем список отходящих линий ГРЩ
mdb_load_name = [] # Список отходящих линий в ГРЩ
mdb_load_name = calc_range[0:, 3].value


# создаем список расчетных файлов распред щитов в папке с файлом ГРЩ.
db_file_list = [] # Список всех файлов DB_* в папке
for file in os.listdir(mdb_file_dir):
    if file.startswith("DB_"):
        db_file_list.append(file)


# выделяем из имён файлов название щитов
db_list = []
db_list_with_excel = []
panels_with_excel = {}
for file_name in db_file_list:
    db_list.append(file_name.split('_')[1]) #полный список щитов по количеству файлов в папке
    panel = file_name.split('_')[1]
    if panel in mdb_load_name:
        db_list_with_excel.append(panel)
        panels_with_excel[panel] = [mdb_load_name.index(panel), mdb_file_dir + '/' + file_name]

# Освновной код синхронизации
t_data = {}
for panel in db_list_with_excel:
    t_data.clear()
    t_data = data_to_transfer(calc_range, panels_with_excel[panel][0])
    panel_range = get_DB_range(panels_with_excel[panel][1])
    panel_range[0, 0].value = t_data['CB_def']
    panel_range[0, 1].value = t_data['CB_type']
    panel_range[0, 2].value = t_data['CB_In']
    panel_range[0, 3].value = t_data['CB_ext_string']
    panel_range[0, 4].value = t_data['cable_mark']
    panel_range[0, 6].value = t_data['cable_section']
    panel_range[0, 7].value = t_data['cable_lenght']
