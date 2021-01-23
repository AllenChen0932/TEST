#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020/4/4 10:37
# @Author  : CE
# @File    : Mainwindow.py

import os
from PyQt5.QtWidgets import (QMainWindow,
                             QAction,
                             QApplication,
                             QDesktopWidget,
                             QPushButton,
                             QVBoxLayout,
                             QWidget)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

import tool.check.check_run as check_run
import tool.check.check_joblist as check_joblist
import tool.check.check_dlc1x as check_dlc1x
import tool.check.check_dlc2x_ed4 as check_dlc2x
import tool.check.check_alarm as check_alarm
import tool.check.check_gain as check_gain

class myPushButton(QPushButton):

    def __init__(self, parent):
        super(myPushButton,self).__init__(parent)
        # self.setPalette(self.pedef())
        font = self.font()
        self.setFont(font)

    def font(self):
        # 设置字体
        font = QFont()
        font.setFamily("Calibri")
        # fontHeight = rect.height()/1
        font.setPixelSize(15)
        # font.setBold(True)
        font.setItalic(True)
        return font

class CheckWindow(QMainWindow):

    def __init__(self,parent = None):

        super(CheckWindow,self).__init__(parent)

        self.initUI()
        # 设置窗口大小
        # self.setGeometry(5,30,300,self.minimumHeight())
        self.setFixedSize(200,300)
        # self.resize(200,90)
        # self.resize(self.maximumWidth(),self.minimumHeight())
        # 设置标题、图标和是否
        self.setWindowTitle('Check Simulation')
        self.setWindowIcon(QIcon('icon/check.ico'))
        self.setAutoFillBackground(True)
        # self.center()

    def initUI(self):
        
        # 菜单栏添加
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        helpMenu = menubar.addMenu('&Help')
        
        # 设置菜单栏、工具栏等动作并绑定槽函数
        file_exitAction = QAction(QIcon('icon/Exit.ico'),'Exit',self)
        # file_exitAction.setShortcut('Ctrl+Q')
        file_exitAction.setStatusTip('exit application')
        file_exitAction.triggered.connect(self.myexit)
        fileMenu.addAction(file_exitAction)

        help_openAction = QAction(QIcon('icon/Text.ico'), 'User Manual', self)
        help_openAction.triggered.connect(self.user_manual)
        helpMenu.addAction(help_openAction)

        # 不同模块定义，使用自定义的MyPushButton组件（具有不同样式字体，后续还有按钮样式设置）
        self.check_run = myPushButton('Check Run')
        self.check_run.clicked.connect(self.check_run_show)
        self.check_run.setFixedHeight(30)
        
        self.check_joblist = myPushButton('Check Joblist')
        self.check_joblist.clicked.connect(self.check_joblist_show)
        self.check_joblist.setFixedHeight(30)
        
        self.check_dlc1x = myPushButton('Check DLC1x')
        self.check_dlc1x.clicked.connect(self.check_dlc1x_show)
        self.check_dlc1x.setFixedHeight(30)
        
        self.check_dlc2x = myPushButton('Check DLC2x')
        self.check_dlc2x.clicked.connect(self.check_dlc2x_show)
        self.check_dlc2x.setFixedHeight(30)
        
        self.check_alarm = myPushButton('Check Alarm')
        self.check_alarm.clicked.connect(self.check_alarm_show)
        self.check_alarm.setFixedHeight(30)
        
        self.check_gain = myPushButton('Check Gain')
        self.check_gain.clicked.connect(self.check_gain_show)
        self.check_gain.setFixedHeight(30)

        # 窗口布局采用垂直是窗口布局
        self.vbox1 = QVBoxLayout()
        self.vbox1.addWidget(self.check_run)
        self.vbox1.addWidget(self.check_joblist)
        self.vbox1.addWidget(self.check_dlc1x)
        self.vbox1.addWidget(self.check_dlc2x)
        self.vbox1.addWidget(self.check_alarm)
        self.vbox1.addWidget(self.check_gain)

        self.mywidget = QWidget()
        self.mywidget.setLayout(self.vbox1)
        self.setCentralWidget(self.mywidget)
        self.show()

    def myexit(self):
        self.close()

    def user_manual(self):
        os.startfile(os.getcwd() + '\\' + 'user manual\Check Simulation.docx')

    def keyPressEvent(self,e):
        if e.key()== QtCore.Qt.Key_Escape:
            self.close()

    def check_run_show(self):
        self.w0 = check_run.my_window()
        self.w0.show()

    def check_joblist_show(self):
        # 槽函数定义，打开bearing窗口
        self.w1 = check_joblist.my_window()
        self.w1.show()

    def check_dlc1x_show(self):
        # 槽函数定义，打开求解器窗口
        self.w2 = check_dlc1x.my_window()
        self.w2.show()

    def check_dlc2x_show(self):
        # 槽函数定义，打开时域扭矩扫频窗口
        self.w3 = check_dlc2x.my_window()
        self.w3.show()

    def check_gain_show(self):
        # 槽函数定义，打开坎贝尔图分析窗口
        self.w4 = check_gain.my_window()
        self.w4.show()

    def check_alarm_show(self):
        # 槽函数定义，打开模态分析窗口
        self.w5 = check_alarm.my_window()
        self.w5.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.center())

if __name__ == '__main__':
    #整个布局采用样式表定义
    import sys
    app = QApplication(sys.argv)

    app.setStyleSheet('''
        QPushButton{
            min-height: 20px;
        }
        QPushButton:pressed {
        background-color: rgb(224, 0, 0);
        border-style: inset;
        }
        QPushButton#cancel{
           background-color: red ;
        }
        myPushButton{
            background-color: #E0E0E0 ;
            height:20px;
            border-style: outset;
            border-width: 2px;
            border-radius: 10px;
            border-color: beige;
            font: Italic 13px;
            min-width: 10em;
            padding: 6px;
        }
        myPushButton:hover {
        background-color: #d0d0d0;
        border-style: inset;
        }
        ''')

    window = CheckWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())