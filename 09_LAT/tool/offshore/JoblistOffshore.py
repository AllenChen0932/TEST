#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/8/19 8:28
# @Author  : CE
# @File    : Plot4Dlc12.py

import sys, os

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QDesktopWidget,
                             QLabel,
                             QLineEdit,
                             QPushButton,
                             QComboBox,
                             QFileDialog,
                             QGridLayout,
                             QWidget,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

from tool.offshore.project2joblist import Offshore_Joblist as p2j
from tool.offshore.Create_Joblist import Create_Joblist as write_run

class MyQLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().text():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event): # 放下文件后的动作
        file_path = event.mimeData().text()
        if '///' in file_path:
            self.setText(file_path[8:].replace('/','\\'))   # 删除local path多余开头
        else:
            self.setText(file_path[5:].replace('/','\\'))   # 删除network path多余开头

class JoblistWindow(QMainWindow):

    def __init__(self, parent=None):

        super(JoblistWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 140)
        self.setWindowTitle('Offshore Joblist')
        self.setWindowIcon(QIcon('./Icon/Text.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Run Path")
        self.lbl1.setFont(QFont("SimSun", 8))

        self.lbl2 = QLabel("Joblist Path")
        self.lbl2.setFont(QFont("SimSun", 8))

        self.lbl3 = QLabel("Joblist Name")
        self.lbl3.setFont(QFont("SimSun", 8))

        # self.lbl4 = QLabel('Joblist Type')
        # self.lbl4.setFont(QFont("SimSun", 8))

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(QFont("SimSun", 8))
        self.lin1.setPlaceholderText("Choose run path...")
        
        self.lin2 = MyQLineEdit()
        self.lin2.setFont(QFont("SimSun", 8))
        self.lin2.setPlaceholderText("Choose joblist path...")

        self.lin3 = QLineEdit()
        self.lin3.setFont(QFont("SimSun", 8))
        # self.lin3.setDisabled(True)
        self.lin3.setPlaceholderText("Define joblist name")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("SimSun", 8))
        self.btn1.clicked.connect(self.choose_run_path)

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.choose_joblist_path)

        self.btn3 = QPushButton("Generate Joblist")
        self.btn3.setFont(QFont("SimSun", 8))
        self.btn3.clicked.connect(self.generate_joblist)

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 0, 0, 1, 1)
        self.grid.addWidget(self.lin1, 0, 1, 1, 5)
        self.grid.addWidget(self.btn1, 0, 6, 1, 1)
        self.grid.addWidget(self.lbl2, 1, 0, 1, 1)
        self.grid.addWidget(self.lin2, 1, 1, 1, 5)
        self.grid.addWidget(self.btn2, 1, 6, 1, 1)
        self.grid.addWidget(self.lbl3, 2, 0, 1, 1)
        self.grid.addWidget(self.lin3, 2, 1, 1, 5)
        self.grid.addWidget(self.btn3, 3, 0, 1, 7)

        self.mywidget.setLayout(self.grid)
        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):

        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    # @pysnooper.snoop()     此处加这句之后程序直接退出
    def choose_run_path(self):

        run_path = QFileDialog.getExistingDirectory(self, "Choose path dialog", r".")

        if run_path:

            self.lin1.setText(run_path.replace('/', '\\'))
            self.lin1.home(True)

    def choose_joblist_path(self):

        joblist_path = QFileDialog.getExistingDirectory(self, "Choose project dialog", r".")

        if joblist_path:

            self.lin2.setText(joblist_path.replace('/', '\\'))
            self.lin2.home(True)

    def generate_joblist(self):

        run_path = self.lin1.text()
        jbl_path = self.lin2.text()
        jbl_name = self.lin3.text()


        if os.path.isdir(run_path) and os.path.isdir(jbl_path) and jbl_name:

            try:
                res = p2j(run_path).run()
                if res:
                    write_run(run_path, jbl_path, jbl_name)
            except Exception as e:
                QMessageBox.about(self, 'Warnning', '%s' %e)
            finally:
                joblist = os.path.join(jbl_path, jbl_name+'.joblist')
                if os.path.exists(joblist):
                    QMessageBox.about(self, 'Warnning', 'Joblist is done!')
                
        elif not os.path.isdir(run_path):
            QMessageBox.about(self, 'Warnning', 'Please choose a run path first!')
        
        elif not os.path.isdir(jbl_path):
            QMessageBox.about(self, 'Warnning', 'Please choose a joblist path first!')
            
        elif not jbl_name:
            QMessageBox.about(self, 'Warnning', 'Please define a joblist name first!')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = JoblistWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())