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

__version__ = "1.0.1"
'''
2020.4.15_v1.0.1 modify the logic for 'short circuit'
'''

class Check_DLC2x(object):
    '''
    check dlc2x to see if the fault is triggered during the simulation
    check result will be outputted to check_dlc2x.txt under the specified result path
    '''

    def __init__(self, output_time, run_path, res_path):
        '''
        :param run_path: path which contains dlc2x
        :param res_path: path to output check resutl
        '''

        self.out_time = output_time
        self.run_path = run_path
        self.res_path = res_path
        self.check_dlc2x()

    def write_txt(self, file, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''

        with open(file, 'w+') as f:

            f.write(content)

    def check_dlc2x(self):

        # print('Begin to Check DLC2X')
        # the minimum time for all load cases to trigger fault
        # initial_time = 20

        content = ''

        for root, dir, names in os.walk(self.run_path, topdown=False):

            for name in names:

                # DLC21
                if '.$PJ' in name and '21_' in name:

                    dlc_name = name.split('.')[0]

                    me_path = os.path.join(root, dlc_name) + '.$ME'

                    time_list = []
                    warn_list = []

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            for line in f.readlines():

                                # 记录初始化之后报的所有warning
                                if 'Time:' in line:

                                    t_temp = float(line.split(':')[1].split('s')[0])

                                    if t_temp > self.out_time:

                                        time_list.append(t_temp)
                                        warn_list.append(line.split('-')[-1])

                            if time_list:

                                warn_cont = ' '.join(warn_list)

                                # N4 Overspeed
                                if 'Controller fault number 7' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: N4 Overspeed, No Fast Shutdown\n' % dlc_name
                                    print('%s: N4 Overspeed, No Fast Shutdown' % dlc_name)

                                # Single Blade Runaway
                                elif 'A_PitchFollowingErrorVspr1' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: Single Blade Runaway, No Fast Shutdown\n' % dlc_name
                                    print('%s: Single Blade Runaway, No Fast Shutdown' % dlc_name)

                                # All Blade Runaway
                                elif 'Controller fault number 6' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: All Blade Runaway, No Fast Shutdown\n' % dlc_name
                                    print('%s: All Blade Runaway, No Fast Shutdown' % dlc_name)

                            else:

                                content += '%s: No fault\n' % dlc_name
                                print('%s: No fault' % dlc_name)

                    else:

                        content += '%s: Not finished!\n' % dlc_name
                        print('%s: Not finished!' % dlc_name)

                # DLC22
                if '.$PJ' in name and '22_' in name:

                    dlc_name = name.split('.')[0]

                    me_path = os.path.join(root, dlc_name) + '.$ME'

                    time_list = []
                    warn_list = []

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            for line in f.readlines():

                                # 记录初始化之后报的所有warning
                                if 'Time:' in line:

                                    t_temp = float(line.split(':')[1].split('s')[0])

                                    if t_temp > self.out_time:

                                        time_list.append(t_temp)
                                        warn_list.append(line.split('-')[-1])

                            if time_list:

                                warn_cont = ' '.join(warn_list)

                                # NA Overspeed
                                if 'Controller fault number 8' in warn_cont and 'Safety system circuit number' not in warn_cont:

                                    content += '%s: NA Overspeed, No Safety System Shutdown\n' % dlc_name
                                    print('%s: NA Overspeed, No Safety System Shutdown' % dlc_name)

                                # Single Blade Seizure
                                elif 'A_PitchFollowingErrorVspr1' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: Single Blade Seizure, No Fast Shutdown\n' % dlc_name
                                    print('%s: Single Blade Seizure, No Fast Shutdown' % dlc_name)

                                # Large Yaw Error
                                elif 'A_HighYawError' in warn_cont and 'VsprNormalShutdownEnter' not in warn_cont:

                                    content += '%s: Large Yaw Error, No Normal Shutdown\n' % dlc_name
                                    print('%s: Large Yaw Error, No Normal Shutdown' % dlc_name)

                                # Short Circuit
                                elif 'A_ElectricalTorqueFollowingError' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: Short Circuit, No Fast Shutdown\n' % dlc_name
                                    print('%s: Short Circuit, No Fast Shutdown' % dlc_name)

                            else:

                                content += '%s: No fault\n' % dlc_name
                                print('%s: No fault' % dlc_name)

                    else:

                        content += '%s: Not finished!\n' % dlc_name
                        print('%s: Not finished!' % dlc_name)

                # DLC23
                if '.$PJ' in name and '23_' in name:

                    dlc_name = name.split('.')[0]

                    me_path = os.path.join(root, dlc_name) + '.$ME'

                    time_list = []
                    warn_list = []

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            for line in f.readlines():

                                # 记录初始化之后报的所有warning
                                if 'Time:' in line:

                                    t_temp = float(line.split(':')[1].split('s')[0])

                                    if t_temp > self.out_time:

                                        time_list.append(t_temp)
                                        warn_list.append(line.split('-')[-1])

                            if time_list:

                                warn_cont = ' '.join(warn_list)

                                # Grid Loss
                                if 'Controller fault number 5' in warn_cont and 'VsprGridLossShutdownEnter' not in warn_cont:

                                    content += '%s: Grid Loss, No Grid Loss Shutdown\n' % dlc_name
                                    print('%s: Grid Loss, No Grid Loss Shutdown' % dlc_name)

                            else:

                                content += '%s: No fault\n' % dlc_name
                                print('%s: No fault' % dlc_name)

                    else:

                        content += '%s: Not finished!\n' % dlc_name
                        print('%s: Not finished!' % dlc_name)

                # DLC24
                if '.$PJ' in name and '24_' in name:

                    dlc_name = name.split('.')[0]

                    me_path = os.path.join(root, dlc_name) + '.$ME'

                    time_list = []
                    warn_list = []

                    if os.path.isfile(me_path):

                        with open(me_path, 'r') as f:

                            for line in f.readlines():

                                # 记录初始化之后报的所有warning
                                if 'Time:' in line:

                                    t_temp = float(line.split(':')[1].split('s')[0])

                                    if t_temp > self.out_time:

                                        time_list.append(t_temp)
                                        warn_list.append(line.split('-')[-1])

                            if time_list:

                                warn_cont = ' '.join(warn_list)

                                # N4 Overspeed
                                if 'Controller fault number 7' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                    content += '%s: N4 Overspeed, No Fast Shutdown\n' % dlc_name
                                    print('%s: N4 Overspeed, No Fast Shutdown' % dlc_name)

                                # Grid Loss
                                elif 'Controller fault number 5' in warn_cont and 'VsprGridLossShutdownEnter' not in warn_cont:

                                    content += '%s: Grid Loss, No Grid Loss Shutdown\n' % dlc_name
                                    print('%s: Grid Loss, No Grid Loss Shutdown' % dlc_name)

                                # Large Yaw Error
                                elif 'A_HighYawError' in warn_cont and 'VsprNormalShutdownEnter' not in warn_cont:

                                    content += '%s: Large Yaw Error, No Normal Shutdown\n' % dlc_name
                                    print('%s: Large Yaw Error, No Normal Shutdown' % dlc_name)

                            else:

                                content += '%s: No fault\n' % dlc_name
                                print('%s: No fault' % dlc_name)

                    else:

                        content += '%s: Not finished!\n' % dlc_name
                        print('%s: Not finished!' % dlc_name)

        if not content:

            content += 'DLC2x finished successfully'

        # check result are outputted to check_dlc2x_list.txt
        txt_name  = 'check_dlc2x.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check DLC2x')
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

        self.label2 = QLabel("DLC2x Path")
        self.label2.setFont(QFont("SimSun", 8))

        self.line2 = QLineEdit()
        self.line2.setFont(QFont("SimSun", 8))
        self.line2.setPlaceholderText("Pick dlc2x")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.load_dlc2x_path)

        self.label3 = QLabel("Simulation output time")
        self.label3.setFont(QFont("SimSun", 8))

        self.line3 = QLineEdit()
        self.line3.setFont(QFont("SimSun", 8))
        self.line3.setText("20")
        self.line3.setToolTip('The shutdown triggered before the output time will be ignored!\n'
                              'Input the minimum simulation output time for all load cases')

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

    def load_dlc2x_path(self):

        self.file_name2 = QFileDialog.getExistingDirectory(self,
                                                           "Choose dlc2x path",
                                                           r".")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        dlc2x_path  = self.line2.text()
        res_path    = self.line1.text()
        output_time = float(self.line3.text())
        print('dlc2x path:', dlc2x_path)
        print('result path:', res_path)

        if dlc2x_path and res_path and output_time:

            Check_DLC2x(output_time, dlc2x_path, res_path)

            QMessageBox.about(self, 'Window', 'Check dlc2x is done!')

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