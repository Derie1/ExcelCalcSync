from cmath import log
import os
from PyQt5 import QtWidgets
import openpyxl


def get_range_dimensions(_xlsx):
    start_row = 0
    row = 1
    col = 0
    while True:
        if mdb_sheet[row][col].value == 1:
            start_row = row
            break
        row += 1
    while True:
        if mdb_sheet[row][col].value == 'EndOfRows':
            last_row = row
            break
        row += 1
    while True:
        if mdb_sheet[row][col].value == 'EndOfColumns':
            last_col = col
            break
        col += 1
    return start_row, last_row - 1, last_col - 1


def update_mdb_load(row, panel, path):
    # db_xlsx = f"{path}/DB_{panel}.xlsx"
    # db_wb = openpyxl.load_workbook(db_xlsx)
    # db_sheet = db_wb['Расчет']
    mdb_sheet[row][5].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$F$2"
    mdb_sheet[row][9].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$G$2"
    mdb_sheet[row][10].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$H$2"
    mdb_sheet[row][11].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$J$16"
    mdb_sheet[row][12].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$J$16"
    mdb_sheet[row][13].value = f"='{path}/[DB_{panel}.xlsx]Расчет'!$L$16"


app = QtWidgets.QApplication([])
mdb_xlsx = QtWidgets.QFileDialog.getOpenFileName(
    caption="Выберите исходный файл Excel... ", filter="XLS (*.xls);XLSX (*.xlsx)")[0]

mdb_path = os.path.dirname(mdb_xlsx)
db_file_list = []  # ? Список всех файлов DB_* в папке формата 'DB_ЩО-1.xlsx'
for file in os.listdir(mdb_path):
    if file.startswith("DB_"):
        db_file_list.append(file)

db_list = []        # ? Список имен щитов в формате 'ЩО-1'
for db in db_file_list:
    _db = (db.rsplit('.', maxsplit=1)[0]).split('_')[1]
    db_list.append(_db)

mdb_wb = openpyxl.load_workbook(mdb_xlsx)
mdb_sheet = mdb_wb['Расчет']

mdb_START_ROW, mdb_LAST_ROW, mdb_LAST_COL = get_range_dimensions(
    mdb_xlsx)  # определяем границы расчетной таблицы

sync_counter = 0
db_list_not_synced = db_list.copy()
logger = []

# * Запуск основного скрипта синхронизации.
for row in range(mdb_START_ROW, mdb_LAST_ROW + 1):
    panel_name = mdb_sheet[row][3].value
    if panel_name in db_list:
        update_mdb_load(row, panel_name, mdb_path)
        db_list_not_synced.remove(panel_name)
        sync_counter += 1
        logger.append(
            f"Линии с именем << {mdb_sheet[row][3].value} >> синхронизирована")
        print(
            f"Линии с именем << {mdb_sheet[row][3].value} >> синхронизирована")
    row += 1

mdb_wb.save(mdb_xlsx)  # сохраняем файл эксель

print(f"\nВсего щитов синхронизировано: {sync_counter}")
print(f"Всего щитов в директории: {len(db_list)}")
print(f"Щиты не были синхронизированы: {db_list_not_synced}")

# собираем лог
logger.append("")
logger.append(f"Всего щитов синхронизировано: {sync_counter}")
logger.append(f"Всего щитов в директории: {len(db_list)}")
logger.append(f"Щиты не были синхронизированы: {db_list_not_synced}")

logfile = './ZAM_CALC/logfile.txt'
with open(logfile, "w") as LOG:
    for elem in logger:
        LOG.write(elem + "\n")
