import pandas as pd
from PyQt5 import QtWidgets

app = QtWidgets.QApplication([])
excel_file = QtWidgets.QFileDialog.getOpenFileName(caption="Выберите исходный файл Excel... ",
                                                   filter="XLS (*.xls);XLSX (*.xlsx)")[0]  # выбираем исхоный файл


data = pd.read_excel(excel_file, sheet_name='AutoCAD')
print(data)

data_calc = pd.read_excel(excel_file, sheet_name='Расчет')
print(data_calc)
