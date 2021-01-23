#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/8/19 8:28
# @Author  : CE
# @File    : Plot4Dlc12.py

import sys, os

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QDesktopWidget,
                             QPushButton,
                             QGridLayout,
                             QWidget)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

class OneClickWindow(QMainWindow):

    def __init__(self, parent=None):

        super(OneClickWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        # self.resize(300, 100)
        self.setFixedSize(300,120)
        self.setWindowTitle('DLC Generator')
        self.setWindowIcon(QIcon('./Icon/program6.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.btn1 = QPushButton("Load Case Table(New name)")
        self.btn1.setFont(QFont("SimSun", 8))
        self.btn1.setFixedHeight(25)
        self.btn1.clicked.connect(self.open_loadcasetable)
        self.btn1.setStyleSheet("QPushButton{color:black}"
                                "QPushButton:hover{color:red;border-style:solid}"
                                "QPushButton{background-color:#CD8500}"
                                "QPushButton{border:2px}"
                                "QPushButton{border-radius:10px}"
                                "QPushButton{padding:2px 4px}")

        self.btn3 = QPushButton("Onshore_One-Click_v2.1.3")
        self.btn3.setFont(QFont("SimSun", 8))
        self.btn3.setFixedHeight(25)
        self.btn3.clicked.connect(self.open_onshore2)
        self.btn3.setStyleSheet("QPushButton{color:black}"
                                "QPushButton:hover{color:red;border-style:solid}"
                                "QPushButton{background-color:#CD8500}"
                                "QPushButton{border:2px}"
                                "QPushButton{border-radius:10px}"
                                "QPushButton{padding:2px 4px}")

        self.btn2 = QPushButton("Offshore_One-Click_v2.0.3")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.setFixedHeight(25)
        self.btn2.clicked.connect(self.open_offshore)
        self.btn2.setStyleSheet("QPushButton{color:black}"
                                "QPushButton:hover{color:red;border-style:solid}"
                                "QPushButton{background-color:#CD8500}"
                                "QPushButton{border:2px}"
                                "QPushButton{border-radius:10px}"
                                "QPushButton{padding:2px 4px}")

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.btn1, 0, 0, 1, 1)
        self.grid.addWidget(self.btn3, 1, 0, 1, 1)
        self.grid.addWidget(self.btn2, 2, 0, 1, 1)

        self.mywidget.setLayout(self.grid)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):

        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    # @pysnooper.snoop()     此处加这句之后程序直接退出
    def open_onshore1(self):
        os.startfile(os.getcwd() + '\subs\DLC_generator_ED4_v2.1.1.exe')

    def open_loadcasetable(self):
        os.startfile(os.getcwd() + '\subs\IEC61400-1_ed4_model_oneclick_new_name_V1.0.xlsm')

    def open_onshore2(self):
        os.startfile(os.getcwd() + '\subs\DLC_generator_ED4_v2.1.3.exe')

    def open_offshore(self):
        os.startfile(os.getcwd() + '\subs\DLC_generator_ED4_v2.0.3.exe')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OneClickWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())