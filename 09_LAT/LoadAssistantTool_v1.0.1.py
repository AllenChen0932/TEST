# -*- coding: utf-8 -*-
"""
Created on Wed Dec 28 14:24:22 2516

@author: 10700700
"""

import os
from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QStyleFactory,
                             QDesktopWidget,
                             QPushButton,
                             QVBoxLayout,
                             QWidget,
                             QAction,
                             qApp,
                             QGroupBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore
from PyQt5.QtCore import QThread, pyqtSignal

import multiprocessing as mp

import PostAssistant as post
import JoblistModify_v1 as joblist
import CheckSimulation as check
import LoadSummaryTable_v1 as lct
import JoblistTrim as jobtrim
import Extrapolation as ep
import DataTransfer as dt
import JoblistIn as jobin
import LoadReport as lr
import OneClick as oneclick
import JoblistOffshore as o2j
import PostExport as pe
import DeleteRun

class MyButton(QPushButton):

    def __init__(self, parent):

        super(MyButton,self).__init__(parent)
        # self.setPalette(self.pedef())
        font = self.font()
        self.setFont(font)
        self.setFixedHeight(30)
        self.setStyleSheet("QPushButton{color:black}"
                           "QPushButton:hover{color:red;border-style:solid}"
                           "QPushButton{background-color:#E0E0E0}"
                           "QPushButton{border:2px}"
                           "QPushButton{border-radius:10px}"
                           "QPushButton{padding:2px 4px}")

    def font(self):
        font = QFont()
        font.setFamily("Calibri")
        # fontHeight = rect.height()/1
        font.setPixelSize(12)
        # font.setBold(True)
        # font.setItalic(True)
        return font

class Mythread(QThread):

    _signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        # print('Thread trim joblist initialization...!')

    def run(self):
        print('Run thread trim joblist...')
        self.window = jobtrim.JobTrimWindow()
        self.window.show()
        self._signal.emit('joblist trim is done!')

class MainWindow(QMainWindow):

    def __init__(self,parent = None):

        super(MainWindow,self).__init__(parent)

        self.initUI()
        self.setGeometry(30,30,235,200)
        # self.resize(1000,61)
        # self.resize(self.maximumWidth(),self.minimumHeight())
        self.setWindowTitle('Load Assistant Tool')
        self.setWindowIcon(QIcon('icon/program1.ico'))
        self.setAutoFillBackground(True)

    def initUI(self):

        font1 = QFont()
        font1.setFamily("Times New Roman")
        font1.setPixelSize(10)
        # font1.setBold(True)
        # font1.setItalic(True)

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        toolMenu = menubar.addMenu('Tool')
        helpMenu = menubar.addMenu('Help')

        exitAction = QAction(QIcon('Icon/exit.ico'), 'Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.triggered.connect(qApp.quit)
        fileMenu.addAction(exitAction)

        toolAction = QAction(QIcon('Icon/backup.ico'), 'One click backup', self)
        toolAction.triggered.connect(self.back_up)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(all)', self)
        toolAction.triggered.connect(self.delete_run)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(except pj/in)', self)
        toolAction.triggered.connect(self.delete_run_except_pj)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(extension)', self)
        toolAction.triggered.connect(self.delete_run_extension)
        toolMenu.addAction(toolAction)

        # toolAction = QAction(QIcon('Icon/copy.ico'), 'Copy', self)
        # toolAction.triggered.connect(self.copy_file)
        # toolMenu.addAction(toolAction)

        helpAction = QAction(QIcon('Icon/doc.ico'), 'User Manual', self)
        helpAction.setShortcut('Ctrl+H')
        helpAction.triggered.connect(self.user_manual)
        helpMenu.addAction(helpAction)

        self.btn_dlc1 = MyButton('One Click')
        self.btn_dlc1.clicked.connect(self.dlc_generator)
        self.btn_dlc1.setFixedHeight(25)

        self.btn_dlc2 = MyButton('Modify Joblist')

        self.btn_dlc2.clicked.connect(self.joblist_modify)
        self.btn_dlc2.setFixedHeight(25)

        self.btn_jbl1 = MyButton('From IN')
        self.btn_jbl1.clicked.connect(self.joblist_from_in)
        self.btn_jbl1.setFixedHeight(25)

        self.btn_jbl2 = MyButton('To Ultimate/Fatigue')
        # self.btn_jbl2.clicked.connect(self.joblist_trim)
        self.btn_jbl2.clicked.connect(self.joblist_trim)
        self.btn_jbl2.setFixedHeight(25)

        self.btn_jbl3 = MyButton('Offshore')
        # self.btn_jbl3.setDisabled(True)
        # self.btn_jbl3.setStyleSheet("QPushButton{color:black}"
        #                             "QPushButton:hover{color:red;border-style:solid}"
        #                             "QPushButton{background-color:#6E6E6E}"
        #                             "QPushButton{border:2px}"
        #                             "QPushButton{border-radius:10px}"
        #                             "QPushButton{padding:2px 4px}")
        self.btn_jbl3.clicked.connect(self.joblist_offshore)
        self.btn_jbl3.setFixedHeight(25)

        self.btn_che1 = MyButton('Check Simulation')
        self.btn_che1.clicked.connect(self.check)
        self.btn_che1.setFixedHeight(25)

        self.btn_pst1 = MyButton('Post Assistant')
        self.btn_pst1.clicked.connect(self.post)
        self.btn_pst1.setFixedHeight(25)

        self.btn_pst2 = MyButton('Component Loads')
        self.btn_pst2.clicked.connect(self.loads_output)
        self.btn_pst2.setFixedHeight(25)

        self.btn_pst3 = MyButton('Extrapolation')
        self.btn_pst3.clicked.connect(self.extrapolation)
        self.btn_pst3.setFixedHeight(25)

        self.btn_lst1 = MyButton('Load Summary Table')
        self.btn_lst1.clicked.connect(self.loadsummary)
        self.btn_lst1.setFixedHeight(25)

        self.btn_com1 = MyButton('Data Transfer')
        self.btn_com1.clicked.connect(self.data_transfer)
        self.btn_com1.setFixedHeight(25)

        self.btn_com2 = MyButton('One Click')
        self.btn_com2.clicked.connect(self.data_exchange)
        self.btn_com2.setFixedHeight(25)

        self.btn_com3 = MyButton('Post Export')
        self.btn_com3.clicked.connect(self.loads_output)
        self.btn_com3.setFixedHeight(25)

        self.btn_lcr1 = MyButton('Load Report')
        self.btn_lcr1.clicked.connect(self.load_report)
        self.btn_lcr1.setFixedHeight(25)

        # dlc generator
        self.vbox1 = QVBoxLayout()
        self.vbox1.addWidget(self.btn_dlc1)
        self.vbox1.addWidget(self.btn_dlc2)
        self.vbox1.addStretch(1)
        self.group1 = QGroupBox('DLC Generator')
        self.group1.setLayout(self.vbox1)
        self.group1.setFont(font1)

        # joblist generator
        self.vbox2 = QVBoxLayout()
        self.vbox2.addWidget(self.btn_jbl1)
        self.vbox2.addWidget(self.btn_jbl2)
        self.vbox2.addWidget(self.btn_jbl3)
        self.group2 = QGroupBox('Joblist Generator')
        self.group2.setLayout(self.vbox2)
        self.group2.setFont(font1)

        # check simulation
        self.vbox3 = QVBoxLayout()
        self.vbox3.addWidget(self.btn_che1)
        self.group3 = QGroupBox('Check Simulation')
        self.group3.setLayout(self.vbox3)
        self.group3.setFont(font1)

        # post processing
        self.vbox4 = QVBoxLayout()
        self.vbox4.addWidget(self.btn_pst3)
        self.vbox4.addWidget(self.btn_pst1)
        # self.vbox4.addWidget(self.btn_pst2)
        self.group4 = QGroupBox('Post Processing')
        self.group4.setLayout(self.vbox4)
        self.group4.setFont(font1)

        # load summary
        self.vbox5 = QVBoxLayout()
        self.vbox5.addWidget(self.btn_lst1)
        self.group5 = QGroupBox('Load Summary')
        self.group5.setLayout(self.vbox5)
        self.group5.setFont(font1)

        # components load
        self.vbox6 = QVBoxLayout()
        self.vbox6.addWidget(self.btn_com2)
        self.vbox6.addWidget(self.btn_com1)
        self.vbox6.addWidget(self.btn_com3)
        self.group6 = QGroupBox('Components Load')
        self.group6.setLayout(self.vbox6)
        self.group6.setFont(font1)

        # load report
        self.vbox7 = QVBoxLayout()
        self.vbox7.addWidget(self.btn_lcr1)
        self.group7 = QGroupBox('Load Calculation Report')
        self.group7.setLayout(self.vbox7)
        self.group7.setFont(font1)

        self.vbox_all = QVBoxLayout()
        self.vbox_all.addWidget(self.group1)
        self.vbox_all.addWidget(self.group2)
        self.vbox_all.addWidget(self.group3)
        self.vbox_all.addWidget(self.group4)
        self.vbox_all.addWidget(self.group5)
        self.vbox_all.addWidget(self.group6)
        self.vbox_all.addWidget(self.group7)
        self.vbox_all.addStretch(1)

        self.mywidget = QWidget()
        self.mywidget.setLayout(self.vbox_all)
        self.setCentralWidget(self.mywidget)
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        # self.center()
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        # self.setWindowFlags(QtCore.Qt.WindowContextHelpButtonHint|QtCore.Qt.WindowSystemMenuHint)
        # self.setAttribute(QtCore.Qt.WA_TranslucentBackground,True)
        self.show()
        desktop = QApplication.desktop()
        rect = desktop.availableGeometry()
        # print(rect)
        # print(self.minimumHeight())
    # def about(self):
    def myexit(self):
        self.close()

    def back_up(self):
        os.startfile(os.getcwd() + '\subs\OneClickBackUp.exe')
        # pass

    def delete_run(self):
        DeleteRun.delete_run()

    def delete_run_except_pj(self):
        DeleteRun.delete_run_except_pj()

    def delete_run_extension(self):
        DeleteRun.delete_ext()

    def copy_file(self):
        pass

    def user_manual(self):
        os.startfile(os.getcwd()+os.sep+r'user manual\User Manual.docx')

    def keyPressEvent(self,e):
        if e.key()== QtCore.Qt.Key_Escape:
            self.close()

    def extrapolation(self):
        self.ep = ep.ExtrapolationWindow()
        self.ep.show()

    def dlc_generator(self):
        self.oneclick = oneclick.OneClickWindow()
        self.oneclick.show()
        # self.showMinimized()

    def data_exchange(self):
        os.startfile(os.getcwd() + '\subs\data_exchange_V1.1.exe')
        # self.showMinimized()

    def joblist_from_in(self):
        self.joblist_in = jobin.JoblistWindow()
        self.joblist_in.show()
        # self.showMinimized()

    def joblist_trim(self):

        self.jobtrim = jobtrim.JobTrimWindow()
        self.jobtrim.show()
        # self.showMinimized()
        # self.btn_jbl2.setDisabled(True)
        # self.jobtrim = Mythread()
        # self.jobtrim._signal.connect(self.set_btn_jbl2)
        # self.jobtrim.start()

    def set_btn_jbl2(self):
        self.btn_jbl2.setDisabled(False)

    def joblist_modify(self):
        self.joblist = joblist.JoblistWindow()
        self.joblist.show()
        # self.showMinimized()

    def joblist_offshore(self):
        self.offshore = o2j.JoblistWindow()
        self.offshore.show()
        # self.showMinimized()

    def check(self):
        self.check = check.CheckWindow()
        self.check.setStyleSheet('''
        # QPushButton{
        #     min-height: 20px;
        # }
        # QPushButton:pressed {
        # background-color: rgb(224, 0, 0);
        # border-style: inset;
        # }
        # QPushButton#cancel{
        #    background-color: red ;
        # }
        # myPushButton{
        #     background-color: #E0E0E0 ;
        #     height:20px;
        #     border-style: outset;
        #     border-width: 2px;
        #     border-radius: 10px;
        #     border-color: beige;
        #     font: Italic 13px;
        #     min-width: 10em;
        #     padding: 6px;
        # }
        # myPushButton:hover {
        # background-color: #d0d0d0;
        # border-style: inset;
        # }
        # ''')
        self.check.show()
        # self.showMinimized()

    def post(self):
        self.post = post.PostWindow()
        self.post.show()
        # self.showMinimized()

    def loadsummary(self):
        self.loadsummary = lct.LoadSumWindow()
        self.loadsummary.show()
        # self.showMinimized()

    def data_transfer(self):
        self.datatransfer = dt.DataTransferWindow()
        self.datatransfer.show()

    def loads_output(self):
        self.loadoutput = pe.LoadOutputWindow()
        self.loadoutput.show()

    def load_report(self):
        self.loadreport = lr.LoadReportWindow()
        self.loadreport.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    mp.freeze_support()

    import sys
    app = QApplication(sys.argv)
    # app.setStyleSheet('''
    #     # QPushButton{
    #     #     min-height: 10px;
    #     # }
    #     QPushButton:pressed {
    #     background-color: rgb(224, 0, 0);
    #     border-style: inset;
    #     }
    #     QPushButton#cancel{
    #        background-color: red ;
    #     }
    #     MyButton{
    #         background-color: #E0E0E0 ;
    #         height:5px;
    #         border-style: outset;
    #         border-width: 2px;
    #         border-radius: 10px;
    #         border-color: beige;
    #         font: Italic 12px;
    #         min-width: 10em;
    #         padding: 6px;
    #     }
    #     MyButton:hover {
    #     background-color: yellow;
    #     border-style: inset;
    #     }
    #     ''')

    window = MainWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())
