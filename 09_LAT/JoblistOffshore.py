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
                             QFileDialog,
                             QTableWidget,
                             QTableView,
                             QHeaderView,
                             QTableWidgetItem,
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

        self.resize(600, 350)
        self.setWindowTitle('Offshore Joblist')
        self.setWindowIcon(QIcon('./Icon/Text.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Run Path")
        self.lbl1.setFont(QFont("Calibri", 9))

        self.lbl2 = QLabel("Joblist Path")
        self.lbl2.setFont(QFont("Calibri", 9))

        self.lbl3 = QLabel("Joblist Name")
        self.lbl3.setFont(QFont("Calibri", 9))

        self.lbl4 = QLabel('DLC list')
        self.lbl4.setFont(QFont("Calibri", 9))

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(QFont("Calibri", 9))
        self.lin1.setPlaceholderText("Choose run path...")
        
        self.lin2 = MyQLineEdit()
        self.lin2.setFont(QFont("Calibri", 9))
        self.lin2.setPlaceholderText("Choose joblist path...")

        self.lin3 = QLineEdit()
        self.lin3.setFont(QFont("Calibri", 9))
        # self.lin3.setDisabled(True)
        self.lin3.setPlaceholderText("Define joblist name")

        # self.lin4 = QLineEdit()
        # self.lin4.setFont(QFont("Calibri", 9))
        # self.lin3.setDisabled(True)
        # self.lin4.setPlaceholderText("Design load case")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("Calibri", 9))
        self.btn1.clicked.connect(self.choose_run_path)

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("Calibri", 9))
        self.btn2.clicked.connect(self.choose_joblist_path)

        self.btn3 = QPushButton("Generate Joblist")
        self.btn3.setFont(QFont("Calibri", 9))
        self.btn3.clicked.connect(self.generate_joblist)

        self.btn4 = QPushButton("Get DLC")
        self.btn4.setFont(QFont("Calibri", 9))
        self.btn4.clicked.connect(self.get_dlc)

        self.tbl1 = QTableWidget(0,2)
        self.tbl1.setFont(QFont("Calibri", 9))
        self.tbl1.setHorizontalHeaderLabels(['Ultimate Load Case', 'Fatigue Load Case'])

        self.tbl1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.tbl1.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.tbl1.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # self.tbl1.verticalHeader().setVisible(False)
        self.tbl1.setEditTriggers(QTableView.NoEditTriggers)

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
        self.grid.addWidget(self.lbl4, 3, 0, 1, 1)
        self.grid.addWidget(self.tbl1, 3, 1,10, 5)
        self.grid.addWidget(self.btn4, 3, 6, 1, 1)
        self.grid.addWidget(self.btn3,13, 0, 1, 7)

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

    def get_dlc(self):

        run_path = self.lin1.text()
        fat_list = ['DLC12', 'DLC24', 'DLC31', 'DLC41', 'DLC64', 'DLC72']

        fatigue_list  = []
        ultimate_list = []

        if os.path.isdir(run_path):

            try:

                if not os.listdir(run_path):
                    raise Exception('Empty run path!')

                for file in os.listdir(run_path):

                    file_path = os.path.join(run_path, file)

                    if file.upper()[:5] in fat_list and os.path.isdir(file_path):
                        fatigue_list.append(file)

                    elif file.upper()[:5] not in fat_list and os.path.isdir(file_path):
                        ultimate_list.append(file)

                for i, dlc in enumerate(ultimate_list):
                    if i > self.tbl1.rowCount()-1:
                        self.tbl1.insertRow(self.tbl1.rowCount())
                    newitem = QTableWidgetItem(dlc)
                    self.tbl1.setItem(i, 0, newitem)

                for i, dlc in enumerate(fatigue_list):
                    if i > self.tbl1.rowCount()-1:
                        self.tbl1.insertRow(i)
                    newitem = QTableWidgetItem(dlc)
                    self.tbl1.setItem(i, 1, newitem)
            except Exception as e:
                QMessageBox.about(self, 'Warnning', 'Error occurs!\n%s' %e)

        else:
            QMessageBox.about(self, 'Window', 'Please choose a run path first!')

    def generate_joblist(self):

        run_path = self.lin1.text()
        jbl_path = self.lin2.text()
        jbl_name = self.lin3.text()

        selected_item = self.tbl1.selectedItems()
        dlc_list = [item.text() for item in selected_item if item.text()]

        if os.path.isdir(run_path) and os.path.isdir(jbl_path) and jbl_name and dlc_list:

            try:
                res = p2j(run_path, dlc_list).run()
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

        elif not dlc_list:
            QMessageBox.about(self, 'Warnning', 'Please choose load case first!')
        
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