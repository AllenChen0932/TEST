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

class check_run(object):
    '''
    check load case simulation result under run to see if there is an error during the simulation
    check result will be outputted to check_run.txt under the specified result path
    '''

    def __init__(self, run_path, res_path):
        '''
        :param run_path: run path which contains all load cases results
        :param res_path: the path to output the check result
        '''

        self.run_path = run_path
        self.res_path = res_path
        self.check_run()

    def write_txt(self, file, content):
        '''
        output the content to a txt
        :param file: txt file path to output check result
        :param content: content to be written in the txt file
        :return: None
        '''

        with open(file, 'w+') as f:

            f.write(content)
            print('Output is done!')

    def check_run(self):

        prj_list = {}

        content = ''

        for root, dir, names in os.walk(self.run_path):
            print(root)
            print(not names, not dir, self.res_path not in root)

            # dir is empty, and not check result directory
            if (not names) and (not dir) and (self.res_path not in root):

                dlc_name = root.split(os.sep)[-1]

                prj_list[dlc_name] = root
                content += 'Empty case: %s\n' %root

            # file exists under root, not include result directory
            elif names and (self.res_path not in root):

                for name in names:

                    # get the name of load case through .$PJ file
                    if '.$PJ' in name:

                        dlc_name = name.split('.')[0]

                        te_path = os.path.join(root, dlc_name) + '.$TE'
                        me_path = os.path.join(root, dlc_name) + '.$ME'
                        pj_path = os.path.join(root, name)


                        if os.path.exists(te_path):

                            with open(te_path, 'r') as f:

                                te = f.read()

                                if 'terminated' in te:

                                    prj_list[name] = root
                                    content += 'Error case: %s\n' %pj_path
                                    # print(dlc_name,': error case')

                                elif 'in progress' in te:

                                    prj_list[name] = root
                                    content += 'Ongoing case: %s\n' %pj_path

                                elif 'completed' in te:

                                    if not os.path.isfile(me_path):

                                        prj_list[name] = root
                                        content += 'No result: %s\n' %pj_path

                                    else:

                                        with open(me_path, 'r') as f:

                                            me = f.read()
                                            if not me:
                                                prj_list[name] = root
                                                content += 'No result: %s\n' % pj_path

                        else:

                            prj_list[name] = root
                            content += 'Not finished: %s\n' % pj_path
                            # print(dlc_name,': not finished')

        # if content is empty, then write 'Run successfully' to the content,
        # which means all simulations run successfully
        if not content:

            content = 'Run successfully!'
            print('run successfully!')

        else:
            # failed jobs will be outputted to failed file
            result_path = os.path.join(self.res_path, 'failed')
            
            for prj, path in prj_list.items():

                dlc_path = path.replace(self.run_path, result_path)

                if not os.path.exists(dlc_path):

                    os.makedirs(dlc_path)

                # destination directory and project files exist, then remove the project file
                else:

                    shutil.rmtree(dlc_path)
                    os.makedirs(dlc_path)

                if os.path.isfile(os.path.join(path, prj)):
                    shutil.copy(os.path.join(path, prj), dlc_path)

        # check result are outputted to check_run_list.txt
        txt_name  = 'check_run_list.txt'
        file_path = os.path.join(self.res_path, txt_name)

        self.write_txt(file_path, content)

class my_window(QMainWindow):

    def __init__(self, parent=None):

        super(my_window, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 50)
        self.setWindowTitle('Check Run')
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
                                                           r"E:/0_TASK/4.8-146-90")

        if self.file_name2:

            self.line2.setText(self.file_name2)

    # @pysnooper.snoop()
    def run_check(self):

        run_path = self.line2.text().replace('\\', '/')
        res_path = self.line1.text().replace('\\', '/')
        print('run path:', run_path)
        print('result path:', res_path)

        if run_path and res_path:

            check_run(run_path, res_path)

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