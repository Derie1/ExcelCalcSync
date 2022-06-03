import os
import xlwings as xw
from PyQt5 import QtCore, QtWidgets


# row_index = panels_with_excel[panel][0]
def data_to_transfer(source_range, row_index):
    transferred_data = {
        'CB_def':  (source_range[row_index, 38].value + source_range[row_index, 0].value),
        'CB_type': source_range[row_index, 39].value,
        'CB_In': source_range[row_index, 42].value,
        'CB_ext_string': source_range[row_index, 44].value,
        'cable_mark': source_range[row_index, 32].value,
        'cable_section': source_range[row_index, 34].value,
        'cable_lenght': source_range[row_index, 35].value
    }
    return transferred_data


# * start main script
app = QtWidgets.QApplication([])

mdb_file = QtWidgets.QFileDialog.getOpenFileName(
    caption="Выберите фай Excel расчета ГРЩ / ВРУ ... ", filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем файл расчетки ГРЩ
mdb_file_dir = os.path.dirname(mdb_file)
# print(f"Директория сканирования файлов DB: {mdb_file}") #? Для проверки

wb_mdb = xw.Book(mdb_file)
calc_sheet = wb_mdb.sheets['Расчет']
# берём диапазон расчета из файла расчетки ГРЩ
calc_range = calc_sheet.range('Calculation')
# print(calc_range)

# создаем список отходящих линий ГРЩ
mdb_load_name = []  # Список отходящих линий в ГРЩ
mdb_load_name = calc_range[0:, 3].value
# print(f"Список отходящих линий в MDB: {mdb_load_name}\nВсего элементов в списке: {len(mdb_load_name)}") #? Для проверки

# создаем список расчетных файлов распред щитов в папке с файлом ГРЩ.
db_file_list = []  # Список всех файлов DB_* в папке
for file in os.listdir(mdb_file_dir):
    if file.startswith("DB_"):
        db_file_list.append(file)
# print(f"Список всех файлов DB в директории: {db_file_list}\nВсего элементов в списке: {len(db_file_list)}") #? Для проверки

# выделяем из имён файлов название щитов
db_list = []
db_list_with_excel = []
panels_with_excel = {}
for file_name in db_file_list:
    # полный список щитов по количеству файлов в папке
    file_name_without_extension = file_name.rsplit('.', maxsplit=1)[0]
    db_list.append(file_name_without_extension.split('_')[1])
    panel = file_name_without_extension.split('_')[1]
    if panel in mdb_load_name:
        db_list_with_excel.append(panel)
        panels_with_excel[panel] = [mdb_load_name.index(
            panel), mdb_file_dir + '/' + file_name]
# print(f"Список щитов в директории MDB: {db_list}") #? Для проверки
# print(f"Список щитов в директории MDB совпадающих с отходящими линиями в MDB: {panels_with_excel}") #? Для проверки

# Освновной код синхронизации
t_data = {}
logger = []
for panel in db_list_with_excel:
    print(f"Синхронизируем щит --== {panel} ==--...")
    logger.append(f"Синхронизируем щит --== {panel} ==--...")
    t_data.clear()
    t_data = data_to_transfer(calc_range, panels_with_excel[panel][0])
    # panel_range = get_DB_range(panels_with_excel[panel][1])
    with xw.App(visible=False) as xl_app:
        wb_db = xw.Book(panels_with_excel[panel][1])
        db_sheet = wb_db.sheets['Расчет']
        panel_range = db_sheet.range('W16:AF16')
        panel_range[0, 0].value = t_data['CB_def']
        panel_range[0, 1].value = t_data['CB_type']
        panel_range[0, 2].value = t_data['CB_In']
        panel_range[0, 3].value = t_data['CB_ext_string']
        panel_range[0, 4].value = t_data['cable_mark']
        panel_range[0, 6].value = t_data['cable_section']
        panel_range[0, 7].value = t_data['cable_lenght']
        wb_db.save()
        wb_db.close()

print(f"Всего щитов синхронизировано: {len(db_list_with_excel)}")

logfile = f"{mdb_file_dir}/db2mdb_loads_log.txt"
with open(logfile, "w") as LOG:
    for elem in logger:
        LOG.write(elem + "\n")
