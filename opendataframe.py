# @Time : 2020/5/13 19:11 

# @Author : like

# @File : opendataframe.py 

# @Description: xx

import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QTableView
from PyQt5.QtCore import QAbstractTableModel, Qt
fname=r'C:\Users\Administrator\Desktop\code\python\pyqt5\resource\5002 政府性基金预算本级支出.xls'
df=pd.read_excel(fname,sheet_name=0,header=3)

class pandasModel(QAbstractTableModel):

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

if __name__ == '__main__':
    app = QApplication(sys.argv)
    model = pandasModel(df)
    view = QTableView()
    view.setModel(model)
    view.resize(800, 600)
    view.show()
    sys.exit(app.exec_())

