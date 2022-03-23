import os
import pandas as pd
import xlwings as xw
from PyQt5 import QtCore, QtWidgets


def data_to_transfer(source_range, row_index):
    transferred_data = {
        'CB_def' :  (source_range[row_index, 38].value + source_range[row_index, 0].value),
        'CB_type' : source_range[row_index, 39].value,
        'CB_In' : source_range[row_index, 42].value,
        'CB_ext_sting' : source_range[row_index, 44].value,
        'cable_mark' : source_range[row_index, 32].value,
        'cable_section' : source_range[row_index, 34].value,
        'cable_lenght' : source_range[row_index, 35].value}
    return transferred_data

"""
def get_DB_range(file_dir):
    db_range_list = []
    for file in os.listdir(file_dir):
        if file.startswith("DB_"):
            #db_files.append(file)
            wb_db = xw.Book(file_dir + '\\' + file)
            db_sheet = wb_db.sheets['Расчет']
            db_range = db_sheet.range('transferred_data')
            db_range_list.append(db_range)
    return db_range_list
"""

def get_DB_range(file):
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
print("\n==============================Список отходящих линий в ГРЩ (mdb_load_name)========================\n", mdb_load_name)

# создаем список расчетных файлов распред щитов в папке с файлом ГРЩ.
db_file_list = [] # Список всех файлов DB_* в папке
for file in os.listdir(mdb_file_dir):
    if file.startswith("DB_"):
        db_file_list.append(file)
print("\n==============================Список файлов DB с MDBв папке (db_file_list)========================\n", db_file_list)

# выделяем из имён файлов название щитов
db_list = []
panels_with_excel = {}
for file_name in db_file_list:
    db_list.append(file_name.split('_')[1])
    panel = file_name.split('_')[1]
    try:
        panels_with_excel[panel] = [mdb_load_name.index(panel), mdb_file_dir + '/' + file_name]
    except:
        continue

print("\n==============================Список щитов из списка файлов (db_list)==============================\n", db_list)
#panel = db_file_list[0].split('_')
#print(panel[1])

# ищем названия щитов в списке отходящих линий ГРЩ и записываем их в словарь с индексом строки
#panels_with_excel = {}
#for panel in db_list:
#    try:
#        panels_with_excel[panel] = [mdb_load_name.index(panel), mdb_file_dir + '\\' + file'
#    except:
#        continue

"""
# ищем названия щитов в списке отходящих линий ГРЩ и записываем их в список со списком щитов и индексов строк
panels_with_excel = []
for panel in db_list:
    try:
        panels_with_excel.append([panel, mdb_load_name.index(panel)])
    except:
        continue
"""

print("\n=====================Список вхождения щитов в список нагрузок (panels_with_excel)==================\n",panels_with_excel)


input('--press Enter to close--')