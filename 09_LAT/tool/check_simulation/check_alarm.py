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

class Check_Alarm(object):
    '''
    to check the ME file if warning is triggered during the simulation
    check result will be outputted to check_alarm_list.txt under the specified result path
    '''

    def __init__(self, dlc_path, res_path, alarm_id, output_time, copy=False):
        '''
        :param run_path: run path which contains all load cases results
        :param res_path: the path to output the check result
        '''

        self.dlc_path  = dlc_path
        self.res_path  = res_path
        self.alarm_id  = alarm_id
        self.out_time  = output_time
        self.copy_flag = copy
        self.check_alarm()

    def write_txt(self, file, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''

        with open(file, 'w+') as f:
            f.write(content)

    def check_alarm(self):

        alarm = 'Alarm number ' + str(self.alarm_id)
        # print(alarm, self.out_time)

        dlc_list = []

        for root, dir, names in os.walk(self.dlc_path):

            for name in names:

                if '.$PJ' in name:

                    dlc_name = name.split('.')[0]
                    me_path  = os.path.join(root, dlc_name) + '.$ME'
                    pj_path  = os.path.join(root, dlc_name) + '.$PJ'

                    time_state = 0

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            for line in f.readlines():

                                if alarm in line and 'raised' in line:

                                    time_state = line.split(':')[1].split('s')[0]

                                    break

                                    # print(dlc_name, time_state)

                        if float(time_state) > float(self.out_time):

                            cont = '%s: Raised %s' %(dlc_name, alarm)
                            print(cont)
                            dlc_list.append(cont)

                            if self.copy_flag:

                                new_directory = root.replace(self.dlc_path, self.res_path)

                                isexist = os.path.exists(new_directory)

                                if not isexist:

                                    os.makedirs(new_directory)

                                shutil.copy(pj_path, new_directory)

                    else:

                        dlc_list.append('%s: Not finished' %dlc_name)
                        print('%s: Not finished' %dlc_name)

        content = ''

        if dlc_list:

            for dlc in dlc_list:

                content += dlc + '\n'
        else:

            content = 'No cases trigger ‘Alarm number %s’!' % str(self.alarm_id)
            print('No cases trigger ‘Alarm number %s’!' % str(self.alarm_id))

        # check result are outputted to check_alarm_list.txt
        txt_name  = 'check_alarm.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check Alarm')
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
        self.line2.setPlaceholderText("Pick dlc directory")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.load_dlc_path)

        self.label3 = QLabel("Simulation output time")
        self.label3.setFont(QFont("SimSun", 8))

        self.line3 = QLineEdit()
        self.line3.setFont(QFont("SimSun", 8))
        self.line3.setText("30")
        self.line3.setToolTip('The shutdown triggered before the output time will be ignored!')

        self.btn3 = QPushButton("Run Check")
        self.btn3.setFont(QFont("SimSun", 8))
        self.btn3.clicked.connect(self.run_check)

        self.label4 = QLabel("Alarm number")
        self.label4.setFont(QFont("SimSun", 8))

        self.line4 = QLineEdit()
        self.line4.setFont(QFont("SimSun", 8))
        # self.line4.setText("30")
        self.line4.setToolTip('The alarm number had been raised, other alarm number will not be checked!')

        self.grid = QGridLayout()
        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.label2, 0, 0, 1, 1)
        self.grid.addWidget(self.line2, 0, 1, 1, 3)
        self.grid.addWidget(self.btn2, 0, 4, 1, 1)
        self.grid.addWidget(self.label1, 1, 0, 1, 1)
        self.grid.addWidget(self.line1, 1, 1, 1, 3)
        self.grid.addWidget(self.btn1, 1, 4, 1, 1)
        self.grid.addWidget(self.label4, 2, 0, 1, 1)
        self.grid.addWidget(self.line4, 2, 1, 1, 3)
        self.grid.addWidget(self.label3, 3, 0, 1, 1)
        self.grid.addWidget(self.line3, 3, 1, 1, 3)
        self.grid.addWidget(self.btn3, 4, 0, 1, 5)

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
                                                           "Choose result path",
                                                           r".")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        dlc_path = self.line2.text().replace('\\', '/')
        res_path = self.line1.text().replace('\\', '/')
        alarm_id = self.line4.text()
        out_time = float(self.line3.text())
        print('dlc path:', dlc_path)
        print('result path:', res_path)
        print('alarm number:', alarm_id)

        if dlc_path and res_path:

            Check_Alarm(dlc_path, res_path, alarm_id, out_time)

            QMessageBox.about(self, 'Window', 'Check alarm is done!')

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