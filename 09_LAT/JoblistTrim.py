#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/8/19 8:28
# @Author  : CE
# @File    : Plot4Dlc12.py

import sys, os

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openpyxl import Workbook, load_workbook

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

class Trim_Joblist(object):
    
    def __init__(self, joblist_path, joblist_name, options, dlc_list=None):
        
        self.joblist  = joblist_path
        self.job_name = joblist_name
        self.options  = options
        self.dlc_list = dlc_list

        self.job_num    = 0

        self.rank_index = 0     # index for <Rank>
        self.resd_index = 0     # index for <ResultDir>
        self.jobr_ender = 0     # index for </JobRuns>
        self.jobr_index = []    # index for <JobRun> and </JobRun>

        self.get_numbers()
        self.write_joblist()

    def get_numbers(self):

        with open(self.joblist, 'r') as f:

            lines = f.readlines()
            for line in lines:
                if '<JobRun>' in line:
                    self.job_num += 1
            print(self.job_num)

            for index, line in enumerate(lines):

                if '<JobRun>' in line:
                    self.jobr_index.append(index)
                    continue

                if '<Rank>' in line:
                    self.rank_index = index
                    continue

                if '<ResultDir>' in line:
                    self.resd_index = index
                    continue

                if '</JobRun>' in line:
                    self.jobr_index.append(index)
                    break

            for index, line in enumerate(lines):

                if '</JobRuns>' in line:
                    self.jobr_ender = index
                    break

    def write_joblist(self):

        new_joblist = os.path.split(self.joblist)[0] + os.sep + self.job_name + '.joblist'
        # print(new_joblist)

        new_list = []
        case_num = 0

        # row numbers for each jobrun
        jobrun_row = self.jobr_index[1]-self.jobr_index[0]+1

        with open(self.joblist, 'r') as f1, open(new_joblist, 'w+') as f2:

            lines = f1.readlines()

            new_list += lines[:self.jobr_index[0]]

            for i in range(self.job_num):

               # resultdir line index
               line = lines[jobrun_row*i+self.resd_index]
               # print(line)

               if self.options == 'Fatigue':

                    if ('DLC12' in line
                        or 'DLC24' in line
                        or 'DLC31' in line
                        or 'DLC41' in line
                        or 'DLC64' in line
                        or 'DLC72' in line):

                        case_num += 1
                        # rank line
                        lines[jobrun_row*i+self.rank_index] = '      <Rank>%s</Rank>\n' %case_num

                        new_list += lines[self.jobr_index[0]+jobrun_row*i:self.jobr_index[1]+jobrun_row*i+1]

               elif self.options == 'Ultimate':

                   if not ('DLC12' in line
                           or 'DLC24' in line
                           or 'DLC31' in line
                           or 'DLC41' in line
                           or 'DLC64' in line
                           or 'DLC72' in line):

                       case_num += 1
                       # rank line
                       lines[jobrun_row*i+self.rank_index] = '      <Rank>%s</Rank>\n' % case_num

                       new_list += lines[self.jobr_index[0]+jobrun_row*i:self.jobr_index[1]+jobrun_row*i+1]

               elif self.options == 'User':

                    for lc in self.dlc_list:

                        if lc in line.upper():

                            case_num += 1

                            # rank line
                            lines[jobrun_row*i+self.rank_index] = '      <Rank>%s</Rank>\n' %case_num

                            new_list += lines[self.jobr_index[0]+jobrun_row*i:self.jobr_index[1]+jobrun_row*i+1]
            print(case_num)
            new_list += lines[self.jobr_ender:]

            # write joblist
            content = ''
            for line in new_list:
                content += line
            f2.write(content)

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

class JobTrimWindow(QMainWindow):

    def __init__(self, parent=None):

        super(JobTrimWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 160)
        self.setWindowTitle('Trim Joblist')
        self.setWindowIcon(QIcon('./Icon/Text.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Original Joblist")
        self.lbl1.setFont(QFont("Calibri", 9))

        self.lbl2 = QLabel("New Joblist Name")
        self.lbl2.setFont(QFont("Calibri", 9))

        self.lbl3 = QLabel("Joblist Type")
        self.lbl3.setFont(QFont("Calibri", 9))

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(QFont("Calibri", 9))
        self.lin1.setPlaceholderText("Choose original joblist")
        
        self.lin2 = QLineEdit()
        self.lin2.setFont(QFont("Calibri", 9))
        self.lin2.setPlaceholderText("Define new joblist name")

        self.lin3 = QLineEdit()
        self.lin3.setFont(QFont("Calibri", 9))
        self.lin3.setDisabled(True)
        self.lin3.setPlaceholderText("DLC list, separated by ','")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("Calibri", 9))
        self.btn1.clicked.connect(self.load_joblist)

        self.btn2 = QPushButton("Run")
        self.btn2.setFont(QFont("Calibri", 9))
        self.btn2.clicked.connect(self.run)

        self.cbx1 = QComboBox()
        self.cbx1.setFont(QFont("Calibri", 9))
        # self.cbx1.setDisabled(True)
        self.cbx1.currentTextChanged.connect(self.type_action)
        self.cbx1.addItem("Please select")
        self.cbx1.addItem("Fatigue")
        self.cbx1.addItem("Ultimate")
        self.cbx1.addItem("User define")

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 0, 0, 1, 1)
        self.grid.addWidget(self.lin1, 0, 1, 1, 4)
        self.grid.addWidget(self.btn1, 0, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 1, 0, 1, 1)
        self.grid.addWidget(self.lin2, 1, 1, 1, 4)
        self.grid.addWidget(self.lbl3, 2, 0, 1, 1)
        self.grid.addWidget(self.cbx1, 2, 1, 1, 4)
        self.grid.addWidget(self.lin3, 3, 1, 1, 4)
        self.grid.addWidget(self.btn2, 4, 0, 1, 6)

        self.mywidget.setLayout(self.grid)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):

        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    # @pysnooper.snoop()     此处加这句之后程序直接退出
    def load_joblist(self):

        joblist, type = QFileDialog.getOpenFileName(self,
                                                         "Choose joblist dialog",
                                                         r".",
                                                         "(*.joblist)")

        if joblist:

            self.lin1.setText(joblist.replace('/', '\\'))
            self.lin1.home(True)

    def type_action(self):

        if self.cbx1.currentText() == 'User define':
            self.lin3.setDisabled(False)
        else:
            self.lin3.setDisabled(True)

    def run(self):

        self.joblist = self.lin1.text()
        self.jobname = self.lin2.text()

        if self.joblist.endswith('joblist') and self.joblist and self.jobname and (self.cbx1.currentText() != 'Please select'):

            try:
                joblist = os.path.join(os.path.split(self.joblist)[0], self.jobname+'.joblist')
                if os.path.isfile(joblist):
                    os.remove(joblist)

                if self.cbx1.currentText() == 'Fatigue':
                    Trim_Joblist(self.joblist, self.jobname, 'Fatigue')

                elif self.cbx1.currentText() == 'Ultimate':
                    Trim_Joblist(self.joblist, self.jobname, 'Ultimate')

                elif self.cbx1.currentText() == 'User define':
                    self.dlc_list = [dlc.strip().upper() for dlc in self.lin3.text().split(',')]
                    Trim_Joblist(self.joblist, self.jobname, 'User', self.dlc_list)

                if os.path.isfile(joblist):
                    QMessageBox.about(self, 'Window', 'Joblist is done!')

            except Exception as e:
                QMessageBox.about(self, 'Warnning', 'Error occurs in trimming joblist\n%s' %e)

        elif not self.joblist:
            QMessageBox.about(self, 'Message', 'Please choose a joblist first!')

        elif not self.joblist.endswith('.joblist'):
            QMessageBox.about(self, 'Message', 'Please choose a right joblist!')

        elif not self.jobname:
            QMessageBox.about(self, 'Message', 'Please input a joblist name first!')

        elif self.cbx1.currentText() == 'Please select':
            QMessageBox.about(self, 'Message', 'Please choose a type first!')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = JobTrimWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())