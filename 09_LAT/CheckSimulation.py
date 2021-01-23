#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020/4/4 10:37
# @Author  : CE
# @File    : Mainwindow.py

import os
from PyQt5.QtWidgets import (QMainWindow,
                             QLabel,
                             QAction,
                             QFileDialog,
                             QApplication,
                             QDesktopWidget,
                             QPushButton,
                             QVBoxLayout,
                             QGridLayout,
                             QGroupBox,
                             QLineEdit,
                             QComboBox,
                             QCheckBox,
                             QMessageBox,
                             QWidget)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

# import tool.check.check_run as check_run
# import tool.check.check_gain as check_gain
from tool.check_simulation.check_dlc2x_ed4 import Check_DLC2x as Check_DLC2x_ed4
from tool.check_simulation.check_dlc2x_ed3 import Check_DLC2x as Check_DLC2x_ed3
from tool.check_simulation.check_joblist import Check_Joblist
from tool.check_simulation.check_dlc1x import Check_DLC1x
from tool.check_simulation.check_alarm import Check_Alarm

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

        # 设置窗口大小
        # self.setGeometry(5,30,300,self.minimumHeight())
        self.setFixedSize(600,300)
        # self.resize(200,90)
        # self.resize(self.maximumWidth(),self.minimumHeight())
        # 设置标题、图标和是否
        self.setWindowTitle('Check Simulation')
        self.setWindowIcon(QIcon('icon/check.ico'))
        self.setAutoFillBackground(True)
        self.setAcceptDrops(True)

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.initUI()

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
        self.lbl1 = QLabel('Run Path')
        self.lbl1.setFont(self.cont_font)
        self.lbl2 = QLabel('Result')
        self.lbl2.setFont(self.cont_font)

        self.btn_run = QPushButton('Run Path')
        self.btn_run.setFont(self.cont_font)
        self.btn_run.setDisabled(True)
        self.btn_run.clicked.connect(self.choose_run)
        # self.btn_run.setFixedHeight(20)
        
        self.btn_joblist = QPushButton('Joblist')
        self.btn_joblist.setFont(self.cont_font)
        self.btn_joblist.setDisabled(True)
        self.btn_joblist.clicked.connect(self.choose_joblist)
        # self.btn_joblist.setFixedHeight(20)

        self.btn_result = QPushButton('Result Path')
        self.btn_result.setFont(self.cont_font)
        self.btn_result.clicked.connect(self.choose_result)
        # self.btn_result.setFixedHeight(20)

        self.btn_check = QPushButton('Check Simulation')
        self.btn_check.setFont(self.cont_font)
        self.btn_check.clicked.connect(self.run_check)

        self.lin1 = MyQLineEdit()         #run
        self.lin1.setFont(self.cont_font)
        self.lin1.setDisabled(True)
        self.lin2 = MyQLineEdit()         #joblist
        self.lin2.setFont(self.cont_font)
        self.lin2.setDisabled(True)
        self.lin3 = MyQLineEdit()         #result
        self.lin3.setFont(self.cont_font)
        self.lin4 = QLineEdit()           #alarm
        self.lin4.setFont(self.cont_font)
        self.lin4.setDisabled(True)
        self.lin5 = QLineEdit('30')       #time
        self.lin5.setFont(self.cont_font)
        self.lin5.setToolTip('Please define initialization time for contoller')
        self.lin5.setDisabled(True)

        self.cbx_run   = QCheckBox('Run')
        self.cbx_run.setFont(self.cont_font)
        self.cbx_run.setDisabled(True)
        self.cbx_job   = QCheckBox('Joblist')
        self.cbx_job.setFont(self.cont_font)
        self.cbx_job.clicked.connect(self.check_joblist)
        self.cbx_dlc1x = QCheckBox('DLC1x')
        self.cbx_dlc1x.setFont(self.cont_font)
        self.cbx_dlc1x.clicked.connect(self.check_dlc1x)
        self.cbx_dlc2x = QCheckBox('DLC2x')
        self.cbx_dlc2x.setFont(self.cont_font)
        self.cbx_dlc2x.clicked.connect(self.check_dlc2x)
        self.cbx_alarm = QCheckBox('Alarm')
        self.cbx_alarm.setFont(self.cont_font)
        self.cbx_alarm.clicked.connect(self.check_alarm)
        self.cbx_gain  = QCheckBox('Optimal Gain')
        self.cbx_gain.setFont(self.cont_font)

        self.cbx1 = QComboBox()
        self.cbx1.setFont(self.cont_font)
        self.cbx1.setDisabled(True)
        # self.cbx1.setFont(QFont("SimSun", 8))
        # self.cbx1.addItem("Please select")
        self.cbx1.addItem("IEC61400-1 Ed4")
        self.cbx1.addItem("IEC61400-1 Ed3/IEC61400-3")

        self.grid1 = QGridLayout()
        self.grid1.addWidget(self.lbl2, 0, 0, 1, 1)
        self.grid1.addWidget(self.lin3, 0, 1, 1, 5)
        self.grid1.addWidget(self.btn_result, 0, 6, 1, 1)
        self.grid1.addWidget(self.lbl1, 1, 0, 1, 1)
        self.grid1.addWidget(self.lin1, 1, 1, 1, 5)
        self.grid1.addWidget(self.btn_run, 1, 6, 1, 1)

        self.group1 = QGroupBox('Input/Output')
        self.group1.setFont(self.title_font)
        self.group1.setLayout(self.grid1)
        #
        self.grid2 = QGridLayout()
        self.grid2.addWidget(self.cbx_dlc1x,   2, 0, 1, 1)
        self.grid2.addWidget(self.lin5,        2, 1, 1, 5)
        self.grid2.addWidget(self.cbx_dlc2x,   3, 0, 1, 1)
        self.grid2.addWidget(self.cbx1       , 3, 1, 1, 5)
        self.grid2.addWidget(self.cbx_alarm,   1, 0, 1, 1)
        self.grid2.addWidget(self.lin4,        1, 1, 1, 5)
        self.grid2.addWidget(self.cbx_job,     0, 0, 1, 1)
        self.grid2.addWidget(self.lin2,        0, 1, 1, 5)
        self.grid2.addWidget(self.btn_joblist, 0, 6, 1, 1)

        self.group2 = QGroupBox('Check Option')
        self.group2.setFont(self.title_font)
        self.group2.setLayout(self.grid2)

        self.vbox2 = QVBoxLayout()
        self.vbox2.addWidget(self.group1)
        self.vbox2.addWidget(self.group2)
        self.vbox2.addWidget(self.btn_check)
        # self.vbox2.addWidget(self.group3)
        self.vbox2.addStretch(1)

        self.mywidget = QWidget()
        self.mywidget.setLayout(self.vbox2)
        self.setCentralWidget(self.mywidget)
        self.show()

    def myexit(self):
        self.close()

    def user_manual(self):
        os.startfile(os.getcwd() + '\\' + 'user manual\Check Simulation.docx')

    def keyPressEvent(self,e):
        if e.key()== QtCore.Qt.Key_Escape:
            self.close()

    def choose_run(self):
        run_path = QFileDialog.getExistingDirectory(self, "Choose path dialog", r".")

        if run_path:
            self.lin1.setText(run_path.replace('/', '\\'))
            self.lin1.home(True)

    def choose_result(self):
        res_path = QFileDialog.getExistingDirectory(self, "Choose path dialog", r".")

        if res_path:
            self.lin3.setText(res_path.replace('/', '\\'))
            self.lin3.home(True)

    def choose_joblist(self):
        joblist, type = QFileDialog.getOpenFileName(self,
                                                         "Choose joblist dialog",
                                                         r".",
                                                         "(*.joblist)")

        if joblist:
            self.lin2.setText(joblist.replace('/', '\\'))
            self.lin2.home(True)

    def check_joblist(self):
        if self.cbx_job.isChecked():
            self.lin2.setDisabled(False)
            self.btn_joblist.setDisabled(False)
        else:
            self.lin2.setDisabled(True)
            self.btn_joblist.setDisabled(True)

    def check_alarm(self):
        if self.cbx_alarm.isChecked():
            self.lin4.setDisabled(False)
            self.lin5.setDisabled(False)
            self.lin1.setDisabled(False)
            self.btn_run.setDisabled(False)
        elif self.cbx_dlc1x.isChecked() or self.cbx_dlc2x.isChecked():
            self.lin4.setDisabled(True)
            # self.btn_run.setDisabled(True)
        else:
            self.lin4.setDisabled(True)
            self.lin5.setDisabled(True)
            self.lin1.setDisabled(True)
            self.btn_run.setDisabled(True)

    def check_dlc1x(self):
        if self.cbx_dlc1x.isChecked():
            self.lin5.setDisabled(False)
            self.lin1.setDisabled(False)
            self.btn_run.setDisabled(False)
        elif not(self.cbx_alarm.isChecked() or self.cbx_dlc2x.isChecked()):
            self.lin5.setDisabled(True)
            self.lin1.setDisabled(True)
            self.btn_run.setDisabled(True)

    def check_dlc2x(self):
        if self.cbx_dlc2x.isChecked():
            self.lin5.setDisabled(False)
            self.lin1.setDisabled(False)
            self.cbx_run.setDisabled(False)
            self.cbx1.setDisabled(False)
        elif not(self.cbx_alarm.isChecked() or self.cbx_dlc1x.isChecked()):
            self.lin5.setDisabled(True)
            self.lin1.setDisabled(True)
            self.btn_run.setDisabled(True)
            self.cbx1.setDisabled(True)
        else:
            self.cbx1.setDisabled(True)

    def run_check(self):

        result_path = self.lin3.text()
        run_path    = self.lin1.text()
        alarm_cont  = self.lin4.text()
        init_time   = self.lin5.text()
        joblist     = self.lin2.text()
        # print(result_path)

        if not result_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a result path first!')

        elif self.cbx_job.isChecked() and not joblist:
            QMessageBox.about(self, 'Warnning', 'Please choose a joblist first!')

        elif self.cbx_job.isChecked() and not joblist.endswith('joblist'):
            QMessageBox.about(self, 'Warnning', 'Please choose a right joblist first!')

        elif (self.cbx_alarm.isChecked() or self.cbx_dlc1x.isChecked() or self.cbx_dlc2x.isChecked()) and not run_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a run path first!')

        elif (self.cbx_alarm.isChecked() or self.cbx_dlc1x.isChecked() or self.cbx_dlc2x.isChecked()) and \
                not (type(eval(init_time))==int or type(eval(init_time)==float)):
            QMessageBox.about(self, 'Warnning', 'Please define a valid initialization time first!')

        elif  not any((self.cbx_job.isChecked(),self.cbx_dlc1x.isChecked(),self.cbx_dlc2x.isChecked(),self.cbx_alarm.isChecked())):
            QMessageBox.about(self, 'Warnning', 'Please choose an option to check!')

        elif self.cbx_alarm.isChecked() and not alarm_cont:
            QMessageBox.about(self, 'Warnning', 'Please define alarm first!')

        else:
            try:
                print('begin to check...')
                if not os.path.exists(result_path):
                    os.makedirs(result_path)

                if self.cbx_job.isChecked():
                    Check_Joblist(joblist, result_path)

                if self.cbx_alarm.isChecked():
                    Check_Alarm(run_path, result_path, alarm_cont, float(init_time))

                if self.cbx_dlc1x.isChecked():
                    Check_DLC1x(float(init_time), run_path, result_path)

                if self.cbx_dlc2x.isChecked() and (self.cbx1.currentIndex()==0):
                    Check_DLC2x_ed4(float(init_time), run_path, result_path)

                if self.cbx_dlc2x.isChecked() and (self.cbx1.currentIndex()==1):
                    Check_DLC2x_ed3(float(init_time), run_path, result_path)

                cbx_list = [self.cbx_job, self.cbx_alarm, self.cbx_dlc1x, self.cbx_dlc2x]

                func_txt = {self.cbx_job: 'check_joblist.txt',
                            self.cbx_alarm: 'check_alarm.txt',
                            self.cbx_dlc1x: 'check_dlc1x.txt',
                            self.cbx_dlc2x: 'check_dlc2x.txt'}

                check_flag   = True
                file_list    = os.listdir(result_path)
                check_isok   = ''
                check_error  = ''

                for cbx in cbx_list:
                    if cbx.isChecked():
                        if func_txt[cbx] in file_list:
                            check_isok += '%s\n' %func_txt[cbx]
                        else:
                            check_flag = False
                            check_error += '%s\n' %func_txt[cbx]
                if check_flag:
                    QMessageBox.about(self, 'Window', 'Check is done!\n%s' %check_isok)
                else:
                    QMessageBox.about(self, 'Window', 'Check is unsucessfully!\n%s' %check_error)

            except Exception as e:
                QMessageBox.about(self, 'Warnning', 'Error occurrs when checking simulation\n%s' %e)

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