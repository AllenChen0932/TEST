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
                             QTableWidget,
                             QTableWidgetItem,
                             QHeaderView,
                             QTableView,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

from tool.offshore.Create_Joblist import Create_Joblist as write_run
from tool.post_assistant.Write_Joblist import Create_Joblist as write_post

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

        self.resize(650, 550)
        self.setWindowTitle('IN to Joblist')
        self.setWindowIcon(QIcon('./Icon/Text.ico'))

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Run/Post Path")
        self.lbl1.setFont(self.cont_font)

        self.lbl2 = QLabel("Joblist Path")
        self.lbl2.setFont(self.cont_font)

        self.lbl3 = QLabel("Joblist Name")
        self.lbl3.setFont(self.cont_font)

        self.lbl4 = QLabel('Joblist Type')
        self.lbl4.setFont(self.cont_font)

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(self.cont_font)
        self.lin1.setPlaceholderText("Choose run or post path")
        
        self.lin2 = MyQLineEdit()
        self.lin2.setFont(self.cont_font)
        self.lin2.setPlaceholderText("Choose joblist path")

        self.lin3 = QLineEdit()
        self.lin3.setFont(self.cont_font)
        # self.lin3.setDisabled(True)
        self.lin3.setPlaceholderText("Define joblist name")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        self.btn1.clicked.connect(self.choose_run_path)

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        self.btn2.clicked.connect(self.choose_joblist_path)

        self.btn3 = QPushButton("Generate Joblist")
        self.btn3.setFont(self.cont_font)
        self.btn3.clicked.connect(self.generate_joblist)

        self.lbl5 = QLabel('DLC list')
        self.lbl5.setFont(self.cont_font)
        self.tbl1 = QTableWidget(0, 2)
        self.tbl1.setDisabled(True)
        self.tbl1.setFont(self.cont_font)
        self.tbl1.setHorizontalHeaderLabels(['Ultimate Load Case', 'Fatigue Load Case'])

        self.tbl1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.tbl1.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.tbl1.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # self.tbl1.verticalHeader().setVisible(False)
        self.tbl1.setEditTriggers(QTableView.NoEditTriggers)

        self.btn4 = QPushButton('Get DLC')
        self.btn4.setDisabled(True)
        self.btn4.setFont(self.cont_font)
        self.btn4.clicked.connect(self.get_dlc)

        self.cbx1 = QComboBox()
        self.cbx1.setFont(self.cont_font)
        # self.cbx1.setDisabled(True)
        self.cbx1.currentTextChanged.connect(self.type_action)
        self.cbx1.addItem("Please select")
        self.cbx1.addItem("Run")
        self.cbx1.addItem("Post")

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 3, 0, 1, 1)
        self.grid.addWidget(self.lin1, 3, 1, 1, 4)
        self.grid.addWidget(self.btn1, 3, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 2, 0, 1, 1)
        self.grid.addWidget(self.lin2, 2, 1, 1, 4)
        self.grid.addWidget(self.btn2, 2, 5, 1, 1)
        self.grid.addWidget(self.lbl3, 1, 0, 1, 1)
        self.grid.addWidget(self.lin3, 1, 1, 1, 4)
        self.grid.addWidget(self.lbl4, 0, 0, 1, 1)
        self.grid.addWidget(self.cbx1, 0, 1, 1, 4)
        self.grid.addWidget(self.lbl5, 4, 0, 1, 1)
        self.grid.addWidget(self.tbl1, 4, 1, 10, 4)
        self.grid.addWidget(self.btn4, 4, 5, 1, 1)
        self.grid.addWidget(self.btn3, 14, 0, 1, 6)

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

    def type_action(self):

        if self.cbx1.currentText() == 'Run':
            self.tbl1.setDisabled(False)
            self.btn4.setDisabled(False)
        else:
            self.tbl1.setDisabled(True)
            self.btn4.setDisabled(True)

    def get_dlc(self):

        run_path = self.lin1.text()
        fat_list = ['DLC12', 'DLC24', 'DLC31', 'DLC41', 'DLC64', 'DLC72']

        fatigue_list  = []
        ultimate_list = []

        self.tbl1.clear()

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
                # print(ultimate_list)
                if not (fatigue_list or ultimate_list):
                    raise Exception('No DLC under path "%s' %run_path)

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

        option = self.cbx1.currentText()

        if os.path.isdir(run_path) and os.path.isdir(jbl_path) and jbl_name and (option in ['Run', 'Post']):

            try:
                if option == 'Run':

                    lcname_path  = {}

                    selected_item = self.tbl1.selectedItems()

                    if not selected_item:
                        raise Exception('Please choose a DLC first!')
                    dlc_list = [item.text() for item in selected_item if item.text()]
                    dlc_list.sort()
                    print(dlc_list)

                    for dlc in dlc_list:
                        dlc_path = os.path.join(run_path, dlc)
                        lc_list  = [lc for lc in os.listdir(dlc_path) if os.path.join(dlc_path,lc)]

                        for lc in lc_list:
                            lc_path   = os.path.join(dlc_path, lc)
                            file_list = os.listdir(lc_path)

                            pj_file = [file for file in file_list if '$PJ' in file.upper()]
                            in_file = [file for file in file_list if 'dtbladed.in' in file.lower() and '.in_' not in file.lower()]
                            if len(pj_file) > 1:
                                raise Exception('More than 2 $PJ files exist:\n%s:%s' %(lc_path,','.join(pj_file)))
                            if len(in_file) > 1:
                                raise Exception('More than 2 .IN files exist:\n%s:%s' %(lc_path, ','.join(in_file)))
                            if len(pj_file) == 0:
                                raise Exception('No $PJ file in load case:\n%s' %lc_path)
                            if len(in_file) == 0:
                                raise Exception('No .IN file in load case:\n%s' %lc_path)

                            lc_name = os.sep.join([dlc, lc, os.path.splitext(pj_file[0])[0]])
                            lcname_path[lc_name] = lc_path
                    print(lcname_path.keys())
                    write_run(lcname_path, lcname_path, jbl_path, jbl_name)

                elif option == 'Post':
                    write_post(run_path, jbl_name, jbl_path)

                joblist = os.path.join(jbl_path, jbl_name + '.joblist')
                if os.path.isfile(joblist):
                    QMessageBox.about(self, 'Window', 'Joblist is done')
                else:
                    QMessageBox.about(self, 'Window', 'Joblist is not created correctly!')

            except Exception as e:
                QMessageBox.about(self, 'Warnning', '%s'%e)

        elif not os.path.isdir(run_path):
            QMessageBox.about(self, 'Warnning', 'Please choose a run path or post path first!')
        
        elif not os.path.isdir(jbl_path):
            QMessageBox.about(self, 'Warnning', 'Please choose a batch path first!')
            
        elif not jbl_name:
            QMessageBox.about(self, 'Warnning', 'Please define a joblist name first!')

        elif option == 'Please select':
            QMessageBox.about(self, 'Warnning', 'Please choose a type first!')

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