import pandas
import xlwings as xw
from PyQt5 import QtWidgets


app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл


data = pandas.read_excel(excel_file)

load_name_column = list(data['Load Name'])
# print(load_name_column)
# print("\n\n")

load_name_column = list(data['Load Name'])
load_name_column_uniq = []
for index in range(len(load_name_column)):
    try:
        fragmented_load_name = load_name_column[index].split('; ')
        # print(fragmented_load_name)
        load_name_uniq = []
        for load_name in fragmented_load_name:
            if load_name not in load_name_uniq:
                load_name_uniq.append(load_name)
        load_name_column_uniq.append("; ".join(load_name_uniq))
    except:
        load_name_column_uniq.append('')
# print(load_name_column_uniq, len(load_name_column_uniq))
# print("\n\n")

data['Load Name'] = load_name_column_uniq


with pandas.ExcelWriter(excel_file, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    data.to_excel(writer, sheet_name='Sheet1')
