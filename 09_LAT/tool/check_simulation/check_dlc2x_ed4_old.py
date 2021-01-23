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

        # DLC21
        dlc21_list = [file for file in os.listdir(self.run_path) if file.startswith('DLC21')]

        for dlc in dlc21_list:

            dlc_path = os.path.join(self.run_path, dlc)

            for root, dir, names in os.walk(dlc_path):

                for name in names:

                    if '.$PJ' in name.upper():

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

                                    # N4 Overspeed-dlc21a
                                    if 'Controller fault number 7' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: N4 Overspeed, No Fast Shutdown\n' % dlc_name
                                        print('%s: N4 Overspeed, No Fast Shutdown' % dlc_name)

                                    # Single Blade Runaway toward fine or feather-dlc21b/21c
                                    elif 'A_PitchFollowingErrorVspr1' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: Single Blade Runaway, No Fast Shutdown\n' % dlc_name
                                        print('%s: Single Blade Runaway, No Fast Shutdown' % dlc_name)

                                    # Yaw Runaway-dlc21d
                                    elif 'A_HighYawError' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: Large Yaw Error, No Fast Shutdown\n' % dlc_name
                                        print('%s: Large Yaw Error, No Fast Shutdown' % dlc_name)

                                else:

                                    content += '%s: No fault\n' % dlc_name
                                    print('%s: No fault' % dlc_name)

                        else:

                            content += '%s: Not finished!\n' % dlc_name
                            print('%s: Not finished!' % dlc_name)

        # DLC22
        dlc22_list = [file for file in os.listdir(self.run_path) if file.startswith('DLC22')]

        for dlc in dlc22_list:

            dlc_path = os.path.join(self.run_path, dlc)

            for root, dir, names in os.walk(dlc_path):

                for name in names:

                    if '.$PJ' in name.upper():

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

                                    # NA Overspeed-22a
                                    if 'Controller fault number 8' in warn_cont and 'Safety system circuit number' not in warn_cont:

                                        content += '%s: NA Overspeed, No Safety System Shutdown\n' % dlc_name
                                        print('%s: NA Overspeed, No Safety System Shutdown' % dlc_name)

                                    # Single Blade Seizure-22b
                                    elif 'A_PitchFollowingErrorVspr1' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: Single Blade Seizure, No Fast Shutdown\n' % dlc_name
                                        print('%s: Single Blade Seizure, No Fast Shutdown' % dlc_name)

                                    # All Blade Runaway-22c
                                    elif 'Controller fault number 6' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: All Blade Runaway, No Fast Shutdown\n' % dlc_name
                                        print('%s: All Blade Runaway, No Fast Shutdown' % dlc_name)

                                    # Short Circuit-22d
                                    elif 'A_ElectricalTorqueFollowingError' in warn_cont and 'VsprGridLossShutdownEnter' not in warn_cont:

                                        content += '%s: Short Circuit, No GridLoss Shutdown\n' % dlc_name
                                        print('%s: Short Circuit, No GridLoss Shutdown' % dlc_name)

                                else:

                                    content += '%s: No fault\n' % dlc_name
                                    print('%s: No fault' % dlc_name)

                        else:

                            content += '%s: Not finished!\n' % dlc_name
                            print('%s: Not finished!' % dlc_name)

        # DLC23
        dlc23_list = [file for file in os.listdir(self.run_path) if file.startswith('DLC23')]

        for dlc in dlc23_list:

            dlc_path = os.path.join(self.run_path, dlc)

            for root, dir, names in os.walk(dlc_path):

                for name in names:

                    if '.$PJ' in name.upper():

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

                                    # triggered by safety system
                                    if 'Safety system circuit' not in warn_cont:

                                        content += '%s: No Safety System Shutdown\n' % dlc_name
                                        print('%s: No Safety System Shutdown' % dlc_name)

                                else:

                                    content += '%s: No fault\n' % dlc_name
                                    print('%s: No fault' % dlc_name)

                        else:

                            content += '%s: Not finished!\n' % dlc_name
                            print('%s: Not finished!' % dlc_name)

        # DLC24
        dlc24_list = [file for file in os.listdir(self.run_path) if file.startswith('DLC24')]

        for dlc in dlc24_list:

            dlc_path = os.path.join(self.run_path, dlc)

            for root, dir, names in os.walk(dlc_path):

                for name in names:

                    if '.$PJ' in name.upper():

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

                                    # triggered by safety system
                                    elif 'Safety system circuit' not in warn_list and '24b_' in name:

                                        content += '%s: No Safety System Shutdown\n' % dlc_name
                                        print('%s: No Safety System Shutdown' % dlc_name)

                                    # Large Yaw Error
                                    elif 'A_HighYawError' in warn_cont and 'VsprNormalShutdownEnter' not in warn_cont:

                                        content += '%s: Large Yaw Error, No Fast Shutdown\n' % dlc_name
                                        print('%s: Large Yaw Error, No Fast Shutdown' % dlc_name)

                                    # pitch error
                                    elif 'A_PitchFollowingErrorVspr1' in warn_cont and 'VsprFastShutdownEnter' not in warn_cont:

                                        content += '%s: Pitch Error, No Fast Shutdown\n' % dlc_name
                                        print('%s: Pitch Error, No Fast Shutdown' % dlc_name)

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
        self.setWindowTitle('Check DLC2x for Ed4')
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
        self.line2.setPlaceholderText("Pick run path")

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