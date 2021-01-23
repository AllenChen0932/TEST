#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/8/19 8:28
# @Author  : CE
# @File    : dynamic_powercurve.py

import os
import sys

import shutil
from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QDesktopWidget,
                             QLabel,
                             QLineEdit,
                             QPushButton,
                             QFileDialog,
                             QGridLayout,
                             QMessageBox,
                             QWidget)
from PyQt5.QtGui import QIcon, QFont

from ..offshore.Create_Joblist import Create_Joblist

class Check_Joblist(object):
    '''
    check joblist to see whether load cases finished successfully, including no run and error case
    check result will be outputted to check_joblist_list.txt under the specified result path
    '''

    def __init__(self, joblist, res_path):
        '''
        :param run_path: run path which contains all load cases results
        :param res_path: the path to output the check result
        '''

        self.joblist  = joblist
        self.res_path = res_path

        self.lc_in_path  = {}
        self.lc_res_path = {}

        self.check_joblist()

    def write_txt(self, file_path, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''
        with open(file_path, 'w+') as f:
            f.write(content)

    def get_file_num(self, path):
        pass

    def check_joblist(self):
        # list to record load cases
        dlc_list = []

        respath_list = []  # result path specified for each case
        jobpath_list = []  # input path specified for each case

        # read result path and job path from joblists
        with open(self.joblist, 'r') as f:

            for line in f.readlines():

                if '<ResultDir>' in line:
                    result_dir = line.split('>')[1].split('<')[0]
                    result_dir = result_dir.replace('\\', '/')
                    respath_list.append(result_dir)
                    continue

                if '<InputFileDir>' in line:
                    job_dir = line.split('>')[1].split('<')[0]
                    job_dir = job_dir.replace('\\', '/')
                    jobpath_list.append(job_dir)
                    continue

        # get file number for each case
        # file_num = 0
        # for path in respath_list[:10]:
        #     if len(os.listdir(path))>file_num:
        #         file_num = len(os.listdir(path))

        # check result
        for i, dlc_path in enumerate(respath_list):

            dlc_name = dlc_path.split('/')[-1]

            pj_path = dlc_path + '/' + dlc_name + '.$PJ'
            me_path = dlc_path + '/' + dlc_name + '.$ME'
            te_path = dlc_path + '/' + dlc_name + '.$TE'

            # check the result path
            if os.path.exists(dlc_path) and os.path.isfile(pj_path):

                # check TE file
                if os.path.isfile(te_path) and os.path.isfile(me_path):
                    # me and te are not empty
                    if os.path.getsize(te_path) and os.path.getsize(me_path):

                        with open(te_path, 'r') as f:

                            te = f.read()
                            # if 'terminated' in file, then record it as 'Error case'
                            if 'terminated' in te:
                                dlc_list.append('Error case: ' + pj_path)

                                self.lc_in_path[dlc_name]  = jobpath_list[i]
                                self.lc_res_path[dlc_name] = dlc_path

                            elif 'in progress' in te:
                                dlc_list.append('Ongoing case: ' + pj_path)

                            elif 'completed' in te:
                                continue
                    # me or te is empty
                    else:
                        dlc_list.append('No result: ' + pj_path)

                        self.lc_in_path[dlc_name]  = jobpath_list[i]
                        self.lc_res_path[dlc_name] = dlc_path

                # if TE file does not exist, then record it as 'Not finished'
                else:
                    dlc_list.append('Not finished: ' + pj_path)

                    self.lc_in_path[dlc_name]  = jobpath_list[i]
                    self.lc_res_path[dlc_name] = dlc_path

            # if the result path is empty, then record it as 'No run'
            else:
                dlc_list.append('No run: ' + pj_path)

                self.lc_in_path[dlc_name]  = jobpath_list[i]
                self.lc_res_path[dlc_name] = dlc_path

        lc_num = len(dlc_list)
        content = '%s \n' % len(dlc_list)

        if lc_num:
            for dlc in dlc_list:
                content += dlc + '\n'
        else:
            content = 'Run successfully!'
            print('Run successfully!')

        # check result are outputted to check_joblist_list.txt
        txt_name  = 'check_joblist.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

        # write joblist
        if dlc_list:
            joblist_name = os.path.splitext(os.path.split(self.joblist)[1])[0]

            Create_Joblist(self.lc_in_path, self.lc_res_path, self.res_path, '_'.join([joblist_name,'failed']))
        print('Check is done!')

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check Joblist')
        self.setWindowIcon(QIcon('icon/check.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.label1 = QLabel("Result Path")
        self.label1.setFont(QFont("SimSun", 8))

        self.line1 = QLineEdit()
        self.line1.setFont(QFont("SimSun", 8))
        self.line1.setPlaceholderText("Choose result path")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("SimSun", 8))
        self.btn1.clicked.connect(self.load_result_path)

        self.label2 = QLabel("Joblist Path")
        self.label2.setFont(QFont("SimSun", 8))

        self.line2 = QLineEdit()
        self.line2.setFont(QFont("SimSun", 8))
        self.line2.setPlaceholderText("Pick joblist")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.load_joblist_path)

        self.btn3 = QPushButton("Run Check")
        self.btn3.setFont(QFont("SimSun", 8))
        self.btn3.clicked.connect(self.run_check)

        self.grid = QGridLayout()
        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.label2, 0, 0, 1, 1)
        self.grid.addWidget(self.line2, 0, 1, 1, 3)
        self.grid.addWidget(self.btn2, 0, 4, 1, 1)
        self.grid.addWidget(self.label1, 1, 0, 1, 1)
        self.grid.addWidget(self.line1, 1, 1, 1, 3)
        self.grid.addWidget(self.btn1, 1, 4, 1, 1)
        self.grid.addWidget(self.btn3, 2, 0, 1, 5)

        self.mywidget.setLayout(self.grid)
        self.setCentralWidget(self.mywidget)

    def load_result_path(self):

        self.file_name1 = QFileDialog.getExistingDirectory(self,
                                                           "Choose result path",
                                                           r".")

        if self.file_name1:

            self.line1.setText(self.file_name1)

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_joblist_path(self):


        self.file_name2, filetype1 = QFileDialog.getOpenFileName(self,
                                                                 "open project file dialog",
                                                                 r".",
                                                                 "project files(*.joblist)")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        joblist_path = self.line2.text().replace('\\', '/')
        res_path = self.line1.text().replace('\\', '/')
        print('joblist:', joblist_path)
        print('result path:', res_path)

        if joblist_path and res_path:

            check_run(joblist_path, res_path)

            QMessageBox.about(self, 'Window', 'Check joblist is done!')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = my_window()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())