# ！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/10/26 11:14
# @Author  : CE
# @File    : joblist.py
'''
1 本程序只适用于通过additional controller parameters对话框来设置控制器参数，不适用于有多个控制器时的模型；
2 有多个控制器的模型需要在每个控制器中定义单独的参数文件；
'''

import sys, os
import configparser
import openpyxl

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QDesktopWidget,
                             QLabel,
                             QLineEdit,
                             QPushButton,
                             QFileDialog,
                             QGridLayout,
                             QVBoxLayout,
                             QWidget,
                             QMessageBox,
                             QTextBrowser,
                             QComboBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

from tool.post_export.writeCompLoads import writeCompLoads
from tool.post_export.writeRainflow_v1_1 import writeRainflow
from tool.post_export.writeUltimate_v2 import writeUltimate


class MyQLineEdit(QLineEdit):
    def __init__(self):
        super(MyQLineEdit, self).__init__()
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

class LoadOutputWindow(QMainWindow):

    def __init__(self, parent=None):

        super(LoadOutputWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.setMinimumSize(700, 150)
        self.setWindowTitle('Post Export')
        self.setWindowIcon(QIcon(r".\icon\excel1.ico"))
        # root = QFileInfo(__file__).absolutePath()
        # self.setWindowIcon(QIcon(root + "./icon/Text_Edit.ico"))

        self.job_list = None
        self.lis_name = None
        self.job_path = None
        self.res_path = None
        self.dll_path = None
        self.xml_path = None
        self.njl_path = None
        self.para_con = None
        self.othe_con = None

        self.initUI()
        self.load_setting()
        self.center()

    def initUI(self):

        self.label1 = QLabel("Post Path")
        self.label1.setFont(QFont("Calibri", 10))
        self.line1 = MyQLineEdit()
        self.line1.setFont(QFont("Calibri", 10))
        self.line1.setPlaceholderText("Pick post path")
        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("Calibri", 10))
        self.btn1.clicked.connect(self.load_post)

        self.label2 = QLabel("Resutl Path")
        self.label2.setFont(QFont("Calibri", 10))
        self.line2 = MyQLineEdit()
        self.line2.setFont(QFont("Calibri", 10))
        self.line2.setPlaceholderText("Pick result path")
        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("Calibri", 10))
        self.btn2.clicked.connect(self.load_result)

        self.label3 = QLabel("Export Type")
        self.label3.setFont(QFont("Calibri", 10))
        self.txt1 = QTextBrowser()

        self.cbx = QComboBox()
        self.cbx.setFont(QFont("Calibri", 10))
        # self.cbx.addItem('All components')
        self.items = ['All Components','Blade','Pitch Bearing','Pitch Lock','Pitch System','Hub','Main Bearing',
                      'Main Shaft','Gearbox_64','Gearbox_144','Nacelle Acc','Yaw Bearing','Tower','Fatigue','Ultimate']
        self.cbx.addItems(self.items)
        self.cbx.currentTextChanged.connect(self.type_action)

        self.label4 = QLabel('LC Table')
        self.label4.setFont(QFont("Calibri", 10))
        self.line3  = MyQLineEdit()
        self.line3.setFont(QFont("Calibri", 10))
        # self.line3.setDisabled(True)

        self.btn3 = QPushButton("...")
        self.btn3.setFont(QFont("Calibri", 10))
        self.btn3.clicked.connect(self.get_excel)
        # self.btn3.setDisabled(True)

        self.label5 = QLabel('Txt/Excel')
        self.label5.setFont(QFont("Calibri", 10))
        self.line4  = MyQLineEdit()
        self.line4.setFont(QFont("Calibri", 10))
        self.line4.setDisabled(True)

        self.label6 = QLabel('Time Step')
        self.label6.setFont(QFont("Calibri", 10))
        self.line5  = MyQLineEdit()
        # self.line5.setText('0.05')
        self.line5.setFont(QFont("Calibri", 10))
        # self.line5.setDisabled(True)

        self.btn4 = QPushButton("Write Excel")
        self.btn4.setFont(QFont("Calibri", 10))
        self.btn4.clicked.connect(self.generate_excel)

        self.btn5 = QPushButton("Save Setting")
        self.btn5.setFont(QFont("Calibri", 10))
        self.btn5.clicked.connect(self.save_setting)

        self.grid1 = QGridLayout()
        # 起始行，起始列，占用行，占用列
        self.grid1.addWidget(self.label1, 0, 0, 1, 1)
        self.grid1.addWidget(self.line1, 0, 1, 1, 5)
        self.grid1.addWidget(self.btn1, 0, 6, 1, 1)
        self.grid1.addWidget(self.label2, 1, 0, 1, 1)
        self.grid1.addWidget(self.line2, 1, 1, 1, 5)
        self.grid1.addWidget(self.btn2, 1, 6, 1, 1)
        self.grid1.addWidget(self.label3, 2, 0, 1, 1)
        self.grid1.addWidget(self.cbx, 2, 1, 1, 5)
        self.grid1.addWidget(self.label4, 3, 0, 1, 1)
        self.grid1.addWidget(self.line3, 3, 1, 1, 5)
        self.grid1.addWidget(self.btn3, 3, 6, 1, 1)
        self.grid1.addWidget(self.label5, 5, 0, 1, 1)
        self.grid1.addWidget(self.line4, 5, 1, 1, 5)
        self.grid1.addWidget(self.label6, 4, 0, 1, 1)
        self.grid1.addWidget(self.line5, 4, 1, 1, 5)
        self.grid1.addWidget(self.btn5, 6, 0, 1, 3)
        self.grid1.addWidget(self.btn4, 6, 4, 1, 3)

        self.main_layout = QVBoxLayout()
        self.main_layout.addLayout(self.grid1)
        self.main_layout.addStretch(1)

        self.mywidget.setLayout(self.main_layout)
        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def save_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if not config.has_section('Post Export'):
            config.add_section('Post Export')

        config['Post Export'] = {'Post Path':self.line1.text(),
                                 'Result Path':self.line2.text(),
                                 'LCT':self.line3.text(),
                                 'Time Step':self.line5.text(),
                                 'Xls Name':self.line4.text()}

        with open("config.ini",'w') as f:
            config.write(f)

    def load_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if config.has_section('Post Export'):
            self.line1.setText(config.get('Post Export','Post Path'))
            self.line2.setText(config.get('Post Export','Result Path'))
            self.line3.setText(config.get('Post Export','LCT'))
            self.line5.setText(config.get('Post Export','Time Step'))
            self.line4.setText(config.get('Post Export','Xls Name'))

    def clear_setting(self):
        self.line1.setText('')
        self.line2.setText('')

    def user_manual(self):
        os.startfile(os.getcwd()+os.sep+'user manual\Load Report.docx')

    def type_action(self):
        items = ['All Components', 'Pitch Bearing', 'Main Bearing', 'Main Shaft', 'Gearbox_64', 'Gearbox_144',
                 'Yaw Bearing', 'Tower']
        if self.cbx.currentText() not in items:
            self.line4.setDisabled(False)
            self.line5.setDisabled(True)
            self.line3.setDisabled(True)
            self.btn3.setDisabled(True)
        else:
            self.line3.setDisabled(False)
            self.line4.setDisabled(True)
            self.line5.setDisabled(False)
            self.btn3.setDisabled(False)

    def load_post(self):
        post_path = QFileDialog.getExistingDirectory(self, "Choose result path", ".")

        if post_path:
            self.line1.setText(post_path.replace('/', '\\'))

    def load_result(self):
        restult_path = QFileDialog.getExistingDirectory(self, "Choose result path", ".")

        if restult_path:
            self.line2.setText(restult_path.replace('/', '\\'))

    def get_excel(self):
        excel_path, extension = QFileDialog.getOpenFileName(self, "Choose result path", ".xlsx")

        if excel_path:
            self.line3.setText(excel_path.replace('/', '\\'))

    def generate_excel(self):

        post_path = self.line1.text()
        res_path  = self.line2.text()
        xls_path  = os.getcwd()+os.sep+'template'
        lct_path  = self.line3.text()
        xls_name  = self.line4.text()
        time_step = self.line5.text()

        if not os.path.isdir(res_path):
            os.makedirs(res_path)

        try:
            if post_path and os.path.isdir(post_path) and res_path and os.path.isdir(res_path):

                comp_flag = {'blade': {'fat': True, 'ult': True},
                             'bladeroot': {'fat': True},
                             'hub': {'fat': True, 'ult': True},
                             'gearbox_64': {'fat': True, 'ult': True},
                             'gearbox_144': {'fat': True, 'ult': True},
                             'pitchbearing': {'fat': True, 'ult': True},
                             'yawbearing': {'fat': True, 'ult': True},
                             'mainbearing': {'fat': True, 'ult': True},
                             'mainshaft': {'fat': True, 'ult': True},
                             'tower': {'fat': True, 'ult': True},
                             'nacacc': {'ult': True},
                             'pitchlock': {'ult': True},
                             'pitchsystem': {'ult': True}}

                if self.cbx.currentText() not in ['Fatigue', 'Ultimate']:
                    if os.path.isdir(xls_path) and time_step and os.path.isfile(lct_path):
                        comp_list = {}
                        if self.cbx.currentText() == 'All Components':
                            comp_list.update((k,v) for k,v in comp_flag.items())
                        elif self.cbx.currentText() == 'Blade':
                            comp_list.update((k,v) for k,v in comp_flag.items() if 'blade' in k)
                        elif self.cbx.currentText() == 'Pitch Bearing':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'pitchbearing' in k)
                        elif self.cbx.currentText() == 'Pitch Lock':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'pitchlock' in k)
                        elif self.cbx.currentText() == 'Pitch System':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'pitchsystem' in k)
                        elif self.cbx.currentText() == 'Hub':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'hub' in k)
                        elif self.cbx.currentText() == 'Yaw Bearing':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'yawbearing' in k)
                        elif self.cbx.currentText() == 'Main Bearing':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'mainbearing' in k)
                        elif self.cbx.currentText() == 'Main Shaft':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'mainshaft' in k)
                        elif self.cbx.currentText() == 'Gearbox_64':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'gearbox_64' in k)
                        elif self.cbx.currentText() == 'Gearbox_144':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'gearbox_144' in k)
                        elif self.cbx.currentText() == 'Nacelle Acc':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'nacacc' in k)
                        elif self.cbx.currentText() == 'Tower':
                            comp_list.update((k,v) for k, v in comp_flag.items() if 'tower' in k)

                        writeCompLoads(post_path, res_path, xls_path, comp_list, time_step, lct_path)
                        QMessageBox.about(self, 'Window', 'Post export is done!')

                    elif not os.path.isdir(xls_path):
                        QMessageBox.about(self, 'warnning', 'Please make sure the xls template path exist!\n%s'%xls_path)
                    elif not time_step:
                        QMessageBox.about(self, 'warnning', 'Please define time step first!')
                    elif not os.path.isfile(lct_path):
                        QMessageBox.about(self, 'warnning', 'Please choose a load case table first!')

                elif xls_name:
                    if self.cbx.currentText() == 'Fatigue':
                        try:
                            txt_path = os.path.join(res_path, xls_name+'.txt')
                            writeRainflow(post_path, content=('DEL',)).write2singletxt(txt_path)
                            QMessageBox.about(self, 'Window', 'Post export is done!')
                        except Exception as e:
                            QMessageBox.about(self, 'Warnning', 'Error occurs in generating fatigue loads!\n%s' %e)

                    elif self.cbx.currentText() == 'Ultimate':
                        try:
                            excel_path = os.path.join(res_path, xls_name+'.xlsx')
                            table = openpyxl.Workbook()
                            writeUltimate(post_path, table, sheetname='Ultimate', rowstart=1, colstart=1)
                            table.save(excel_path)
                            QMessageBox.about(self, 'Window', 'Post export is done!')
                        except Exception as e:
                            QMessageBox.about(self, 'Warnning', 'Error occurs in generating ultimate loads!\n%s' %e)

                else:
                    QMessageBox.about(self, 'warnning', 'Please define a excel name first!')

            elif not post_path:
                QMessageBox.about(self, 'warnning', 'Please choose a post path first!')
            elif not os.path.isdir(post_path):
                QMessageBox.about(self, 'warnning', 'Please make sure the post path is right!')
            elif not res_path:
                QMessageBox.about(self, 'warnning', 'Please choose a result path first!')
            elif not os.path.isdir(res_path):
                QMessageBox.about(self, 'warnning', 'Please make sure the result path is right!')

        except Exception as e:
            QMessageBox.about(self, 'Warnning', 'Error occurs in post export\n%s'%e)

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = LoadOutputWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())

