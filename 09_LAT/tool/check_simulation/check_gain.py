#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/8/19 8:28
# @Author  : CE
# @File    : dynamic_powercurve.py

import os
import sys

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

class check_gain:
    '''
    check load case simulation result under run to see if there is the optimal mode gain are the same
    check result will be outputted to check_gain.txt under the specified result path
    '''

    def __init__(self, run_path, res_path):
        '''
        :param run_path: run path which contains all load cases results
        :param res_path: the path to output the check result
        '''

        self.run_path = run_path
        self.res_path = res_path
        self.check_gain()

    def write_txt(self, file, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''

        with open(file, 'w+') as f:

            f.write(content)

    def check_gain(self):

        content  = ''

        for root, dir, names in os.walk(self.run_path):

            for name in names:

                if '.$PJ' in name:

                    name = name.split('.')[0]

                    pj_path = os.path.join(root, name) + '.$PJ'

                    pm_path = os.path.join(root, name) + '.$PM'

                    if os.path.exists(pm_path) and os.path.exists(pj_path):

                        with open(pj_path, 'r') as f1, open(pm_path, 'r') as f2:

                            gain1 = gain2 = gain3 = None

                            line_count = 0
                            line_flag  = 0

                            for line in f1.readlines():

                                # redd the gain specified under the control module interface
                                if 'GAIN_TSR' in line:

                                    gain1 = float(line.split()[1])
                                    # print(gain1)
                                    continue

                                # read the gain specified in the additional controller parameters box
                                # not xml
                                if 'P_OptimalModeGain' in line and line_count == 0:

                                    line_count += 1
                                    line_flag  = 1
                                    # print(line)

                                elif line_count < 2 and line_flag == 1:
                                    line_count += 1
                                    continue

                                elif line_count == 2 and line_flag == 1:
                                    gain2 = float(line.split(';')[2][:-3])
                                    # print(gain2)
                                    break

                            line_count = 0
                            line_flag  = 0

                            # read the gain recorded in the .$PM file
                            for line in f2.readlines():

                                if 'P_OptimalModeGain' in line and line_count == 0:

                                    line_count += 1
                                    line_flag  = 1
                                    # print(line)

                                elif line_count < 3 and line_flag == 1:
                                    line_count += 1
                                    continue

                                elif line_count == 3 and line_flag == 1:
                                    # print(line)
                                    gain3 = float(line.strip())
                                    # print(gain3)
                                    break

                            # eliminate the load case without external controller, such as idling or parked
                            if gain1 == None:

                                break

                            elif gain1 == gain2 and gain2 == gain3:

                                break

                            else:

                                content += '%s:\n' %name
                                content += 'GAIN_TSR: %s\n' % gain1
                                content += 'GAIN_ADD: %s\n' % gain2
                                content += 'GAIN_$PM: %s\n' % gain3
                                content += '\n'
                                print('%s' %name)
                                print('GAIN_TSR: %s' % gain1)
                                print('GAIN_ADD: %s' % gain2)
                                print('GAIN_$PM: %s' % gain3)
                                break

        if not content:

            content = 'The gains from model, parameter and $pm file are the same!'
            print('The gains from model, parameter and $pm file are the same!')

        # check result are outputted to check_gain_list.txt
        txt_name  = 'check_gain_list.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check Gain')
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

        self.label2 = QLabel("DLC Path")
        self.label2.setFont(QFont("SimSun", 8))

        self.line2 = QLineEdit()
        self.line2.setFont(QFont("SimSun", 8))
        self.line2.setPlaceholderText("Pick DLC result path")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.load_dlc_path)

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

    def load_dlc_path(self):

        self.file_name2 = QFileDialog.getExistingDirectory(self,
                                                           "Choose DLC path",
                                                           r".")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        run_path = self.line2.text().replace('\\', '/')
        res_path = self.line1.text().replace('\\', '/')
        print('run path:', run_path)
        print('result path:', res_path)

        if run_path and res_path:

            check_gain(run_path, res_path)

            QMessageBox.about(self, 'Window', 'Check run is done!')

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
    # check_run(r'E:\0_TASK\4.8-146-90\run',r'F:\A_My_Pycharm\35_Check_run\result')