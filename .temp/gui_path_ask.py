from PyQt5 import QtCore, QtWidgets

app = QtWidgets.QApplication([])

mdb_file = QtWidgets.QFileDialog.getOpenFileName()[0] # выбираем файл расчетки ГРЩ
print(mdb_file)