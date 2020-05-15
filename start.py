import os
import sys

import pandas as pd
from PyQt5.QtCore import QObject, pyqtSignal
from PyQt5.QtGui import QTextCursor
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QDesktopWidget

import excel
from ExcelSolve import ExcelSolve
from opendataframe import pandasModel


# 后台程序，只要把前端集成过来就行
# 后台程序，只要把前端集成过来就行

class Stream(QObject):
    """Redirects console output to text widget."""
    try:
        newText = pyqtSignal(str)
    except:
        print('Stream error')
    def write(self, text):
        self.newText.emit(str(text))

class MainCode(QMainWindow, excel.Ui_MainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        excel.Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.open_btn.clicked.connect(self.on_open)
        #self.start_btn.clicked.connect(partial(self.start, self.keyword.text(), self.column_name.text(),os.path.split(self.file_path.text())[0]))
        self.start_btn.clicked.connect(lambda :self.genMastClicked(self.keyword.text(), self.column_name.text(),os.path.split(self.file_path.text())[0]))
        #self.start_btn.clicked.connect(partial(self.start, 'aaa', 'vvv', './'))
        self.stop_btn.clicked.connect(self.get_arg)

        sys.stdout = Stream(newText=self.onUpdateText)

    def onUpdateText(self, text):
        """Write console output to text widget.修改显示位置"""
        cursor = self.run_info.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text)
        self.run_info.setTextCursor(cursor)
        self.run_info.ensureCursorVisible()


    def on_open(self):
        try:
            FullFileName, _ = QFileDialog.getOpenFileName(self, '选择文件', r'../resource', 'Excel (*.xls*)')
            self.file_path.setText(FullFileName)
            print(FullFileName)
            df = pd.read_excel(FullFileName, sheet_name=0, header=3)
            model = pandasModel(df)
            self.tableView.setModel(model)
        except:
            print('open error')

    def get_arg(self):
        keyword = self.keyword.text()
        column_name = self.column_name.text()
        file_path = os.path.split(self.file_path.text())[0]
        print(keyword)
        print(column_name)
        print(file_path)

    def start(self, keyword, columnname, filepath):
        print(keyword)
        print(columnname)
        print(filepath)
        try:
            ex = ExcelSolve(keyword, columnname, filepath)
            ex.main()
        except:
            print('run error')

    # 关闭事件，重写点击x时的操作

    def closeEvent(self, event):

        reply = QMessageBox.question(self, 'Message',
                                     "Are you sure to quit?", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

        # 控制窗口显示在屏幕中心的方法

    def center(self):

        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点，QtGui,QDesktopWidget类提供了用户的桌面信息,包括屏幕大小
        cp = QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def genMastClicked(self,keyword, columnname, filepath):
        """Runs the main function."""
        print('Running...')
        self.start(keyword, columnname, filepath)
        print('Done.')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    md = MainCode()
    md.center()
    md.show()
    sys.exit(app.exec_())
