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

class Check_DLC1x(object):
    '''
    check dlc1x to see if there is a 'shut down' during power production
    check result will be outputted to check_dlc1x.txt under the specified result path
    '''

    def __init__(self, out_time, run_path, res_path, all_shutdown=False):
        '''
        :param out_time: the time for Bladed to output result
        :param run_path: the path which contains all DLC1X simulation results
        :param res_path: the path to output check result
        :param all_shutdown: bool flag to decide how to record the check results, all or just shut down
        '''

        self.out_time = out_time
        self.run_path = run_path
        self.res_path = res_path
        self.all_shut = all_shutdown
        self.check_dlc1x()

    def write_txt(self, file, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''

        with open(file, 'w+') as f:

            f.write(content)

    def check_dlc1x(self):

        # print('Begin to Check DLC1X')
        content = ''

        dlc1x_list = [file for file in os.listdir(self.run_path) if file.startswith('DLC1')]

        for dlc in dlc1x_list:

            dlc_path = os.path.join(self.run_path, dlc)
            lc_list  = os.listdir(dlc_path)

            for lc in lc_list:

                lc_path = os.path.join(dlc_path, lc)

                if os.path.isdir(lc_path):

                    me_path = os.path.join(lc_path, lc+'.$ME')
                    # print(me_path)

                    time_list = []
                    warn_list = []

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            lines = f.readlines()
                            for line in lines:

                                # record all warnings occurring after the time for Bladed to output result
                                if 'Time:' in line:
                                    time = float(line.split()[1][:-1])

                                    if time > float(self.out_time):
                                        time_list.append(time)
                                        warn_list.append(line.split('-')[-1].strip())

                        if self.all_shut:
                            content += ''.join(warn_list)
                            content += '\n'

                        else:
                            warn_cont = [warn for warn in warn_list if 'Shutdown' in warn]

                            if warn_cont:
                                content += '%s: %s\n' %(lc, warn_cont[0].strip())
                                print('%s: %s' %(lc, warn_cont[0].strip()))

                    else:
                        content += '%s: Not finished!\n' %lc
                        print('%s: Not finished!' %lc)

        if not content:
            content += 'DLC1x finished successfully'
            print('DLC1x finished successfully')

        # check result are outputted to check_dlc1x.txt
        txt_name  = 'check_dlc1x.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check DLC1x')
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

        self.label2 = QLabel("DLC1x Path")
        self.label2.setFont(QFont("SimSun", 8))

        self.line2 = QLineEdit()
        self.line2.setFont(QFont("SimSun", 8))
        self.line2.setPlaceholderText("Pick dlc1x")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.load_dlc1x_path)

        self.label3 = QLabel("Simulation output time")
        self.label3.setFont(QFont("SimSun", 8))

        self.line3 = QLineEdit()
        self.line3.setFont(QFont("SimSun", 8))
        self.line3.setText("30")
        self.line3.setToolTip('The shutdown triggered before the output time will be ignored!')

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
        self.grid.addWidget(self.label3, 2, 0, 1, 1)
        self.grid.addWidget(self.line3, 2, 1, 1, 3)
        self.grid.addWidget(self.btn3, 3, 0, 1, 5)

        self.mywidget.setLayout(self.grid)
        self.setCentralWidget(self.mywidget)

    def load_result_path(self):

        self.file_name1 = QFileDialog.getExistingDirectory(self,
                                                           "Choose result path",
                                                           r".")

        if self.file_name1:

            self.line1.setText(self.file_name1)

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_dlc1x_path(self):

        self.file_name2 = QFileDialog.getExistingDirectory(self,
                                                           "Choose dlc1x path",
                                                           r".")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        dlc1x_path  = self.line2.text()
        res_path    = self.line1.text()
        output_time = float(self.line3.text())
        print('dlc1x path:', dlc1x_path)
        print('result path:', res_path)

        if dlc1x_path and res_path and output_time:

            Check_DLC1x(output_time, dlc1x_path, res_path)

            QMessageBox.about(self, 'Window', 'Check dlc1x is done!')

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