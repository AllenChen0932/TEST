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
                             QGroupBox,
                             QFileDialog,
                             QGridLayout,
                             QVBoxLayout,
                             QHBoxLayout,
                             QWidget,
                             QAction,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

from tool.post_assistant.Write_Ultimate import Ultimate
from tool.post_assistant.Write_Bstats import Bstats
from tool.post_assistant.Write_Joblist import Create_Joblist

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

class Extrapolation(object):
    
    def __init__(self, run_path, joblist_name, post_path):
        
        self.run_path = run_path
        self.job_name = joblist_name
        self.pst_path = post_path

        self.dlc11_path = {}
        self.dlc13_path = {}
        self.bld_tip_sec = None

        self.get_dlc_var()
        self.write_pj()
        self.create_joblist()

    def get_dlc_var(self):

        dlc_list = os.listdir(self.run_path)
        print(dlc_list)

        if 'DLC11' not in dlc_list:
            raise Exception('DLC11 not exist under %s' %self.run_path)

        for dlc in dlc_list:
            dlc_path = os.path.join(self.run_path, dlc)
            # get all dlc13
            if dlc.startswith('DLC13') and os.path.isdir(dlc_path):
                self.dlc13_path[dlc] = dlc_path

        if not self.dlc13_path:
            raise Exception('DLC13 not exist under %s' %self.run_path)

        dlc11_path = os.path.join(self.run_path, 'DLC11')
        # first load case in DLC11
        lc_path    = os.path.join(dlc11_path, os.listdir(dlc11_path)[0])
        files_list = os.listdir(lc_path)

        pj_file  = [file for file in files_list if os.path.splitext(file)[-1].upper() == '.$PJ'][0]
        lc_name  = os.path.splitext(pj_file)[0]
        res_file = os.path.join(lc_path, lc_name + '.%18')
        if not res_file:
            raise Exception('No blade deflection results under:\n%s' %lc_path)

        with open(res_file, 'r') as f:
            lines = f.readlines()
            for line in lines:
                if line.startswith('AXIVAL'):
                    section_list = line.strip().split()[1:]
                    self.bld_tip_sec = section_list[-1]
                    break

    def write_pj(self):

        # DLC11
        # write Mx, My, x-deflection
        chan_dict = {'bstats': [['Blade root 1 Mx', '22'],
                                ['Blade root 2 Mx', '22'],
                                ['Blade root 3 Mx', '22'],
                                ['Blade root 1 My', '22'],
                                ['Blade root 2 My', '22'],
                                ['Blade root 3 My', '22'],
                                ['Blade 1 x-deflection (perpendicular to rotor plane)', '18', self.bld_tip_sec],
                                ['Blade 2 x-deflection (perpendicular to rotor plane)', '19', self.bld_tip_sec],
                                ['Blade 3 x-deflection (perpendicular to rotor plane)', '20', self.bld_tip_sec]]}

        post_path = os.path.join(self.pst_path, '01_Extrapolation\DLC11')
        Bstats(self.run_path, ['DLC11'], chan_dict, post_path)

        # wind speed
        chan_dict = {'bstats_1': [['Cup anemometer wind speed', '14']]}
        post_path = os.path.join(self.pst_path, '01_Extrapolation\DLC11\ws')
        Bstats(self.run_path, ['DLC11'], chan_dict, post_path, False)

        # DLC13
        ult_chan = dict({'ep': [['Blade root 1 Mx','Blade root 2 Mx','Blade root 3 Mx',
                                 'Blade root 1 My','Blade root 2 My','Blade root 3 My',
                                 'Blade 1 x-deflection (perpendicular to rotor plane)',
                                 'Blade 2 x-deflection (perpendicular to rotor plane)',
                                 'Blade 3 x-deflection (perpendicular to rotor plane)'], [self.bld_tip_sec]]})

        for dlc,path in self.dlc13_path.items():
            post_path = os.path.join(self.pst_path, os.sep.join(['01_Extrapolation', dlc]))
            Ultimate(self.run_path, post_path, [dlc], ult_chan, '1', False, '#', '+', True)

    def create_joblist(self):
        joblist_path = os.path.join(self.pst_path, '00_Joblist')
        post_path    = os.path.join(self.pst_path, '01_Extrapolation')

        Create_Joblist(post_path, self.job_name, joblist_path)

class ExtrapolationWindow(QMainWindow):

    def __init__(self, parent=None):

        super(ExtrapolationWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.setMinimumSize(650, 160)
        self.setWindowTitle('Extrapolation')
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

        # 菜单栏添加
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        helpMenu = menubar.addMenu('&Help')

        # 设置菜单栏、工具栏等动作并绑定槽函数
        file_exitAction = QAction(QIcon('icon/Exit.ico'), 'Exit', self)
        # file_exitAction.setShortcut('Ctrl+Q')
        file_exitAction.setStatusTip('exit application')
        file_exitAction.triggered.connect(self.myexit)
        fileMenu.addAction(file_exitAction)

        help_openAction = QAction(QIcon('icon/Text.ico'), 'User Manual', self)
        help_openAction.triggered.connect(self.user_manual)
        helpMenu.addAction(help_openAction)

        self.lbl1 = QLabel("Run Path")
        self.lbl1.setFont(self.cont_font)

        self.lbl2 = QLabel("Post Path")
        self.lbl2.setFont(self.cont_font)

        # self.lbl3 = QLabel("Joblist")
        # self.lbl3.setFont(self.cont_font)

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(self.cont_font)
        self.lin1.setPlaceholderText("Choose run path")
        
        self.lin2 = MyQLineEdit()
        self.lin2.setFont(self.cont_font)
        self.lin2.setPlaceholderText("Choose post path")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        self.btn1.clicked.connect(self.load_run_path)

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        self.btn2.clicked.connect(self.load_post_path)

        self.btn3 = QPushButton('Generate joblist')
        self.btn3.setFont(self.cont_font)
        self.btn3.clicked.connect(self.create_joblist)

        self.btn4 = QPushButton('Extrapolation Calculation')
        # self.btn4.setDisabled(True)
        self.btn4.setFont(self.cont_font)
        self.btn4.clicked.connect(self.check_extrapolation)

        self.group1 = QGroupBox('Post processing')
        self.group1.setFont(self.title_font)
        self.grid = QGridLayout()
        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 0, 0, 1, 1)
        self.grid.addWidget(self.lin1, 0, 1, 1, 4)
        self.grid.addWidget(self.btn1, 0, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 1, 0, 1, 1)
        self.grid.addWidget(self.lin2, 1, 1, 1, 4)
        self.grid.addWidget(self.btn2, 1, 5, 1, 1)
        self.grid.addWidget(self.btn3, 2, 0, 1, 6)
        self.group1.setLayout(self.grid)

        self.group2 = QGroupBox('Check')
        self.group2.setFont(self.title_font)
        self.hbox = QHBoxLayout()
        self.hbox.addWidget(self.btn4)
        self.group2.setLayout(self.hbox)

        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.group1)
        self.vbox.addWidget(self.group2)
        self.vbox.addStretch(1)

        self.mywidget.setLayout(self.vbox)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def myexit(self):
        self.close()

    def user_manual(self):
        os.startfile(os.getcwd() + '\\' + 'user manual\Extrapolation.pdf')

    # @pysnooper.snoop()     此处加这句之后程序直接退出
    def load_run_path(self):

        run_path = QFileDialog.getExistingDirectory(self, "Choose run path", r".")

        if run_path:

            self.lin1.setText(run_path.replace('/', '\\'))
            self.lin1.home(True)

    def load_post_path(self):

        post_path = QFileDialog.getExistingDirectory(self, "Choose post path", r".")

        if post_path:

            self.lin2.setText(post_path.replace('/', '\\'))
            self.lin2.home(True)

    def create_joblist(self):

        run_path = self.lin1.text()
        # run_path = r'\\172.20.0.4\fs02\CE\V3\loop05\run_0615'
        pst_path = self.lin2.text()
        # pst_path = r'\\172.20.0.4\fs02\CE\V3\post_test'

        if os.path.isdir(run_path) and os.path.isdir(pst_path):
            print('Begin to handle extrapolation')

            try:
                Extrapolation(run_path, 'ep', pst_path)

                joblist = os.path.join(pst_path, '00_Joblist\ep.joblist')
                if os.path.isfile(joblist):
                    QMessageBox.about(self, 'Window', 'Extrapolation is done!')
                else:
                    QMessageBox.about(self, 'Window', 'Extrapolation is not created correctly!')

            except Exception as e:
                QMessageBox.about(self, 'Warnning', '%s'%e)

        elif not os.path.isdir(run_path):
            QMessageBox.about(self, 'Warnning', 'Please choose run path first!')

        elif not os.path.isdir(pst_path):
            QMessageBox.about(self, 'Warnning', 'Please choose post path first!')

    def check_extrapolation(self):
        os.startfile(os.getcwd() + '\\' + r'subs\UlExtra_v2.1.xlsm')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExtrapolationWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())