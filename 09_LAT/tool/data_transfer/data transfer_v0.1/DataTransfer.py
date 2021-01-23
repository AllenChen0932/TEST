#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/23/2020 9:32 AM
# @Author  : CE
# @File    : Data_Window.py

import sys, os
import configparser

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QTextEdit,
                             QCheckBox,
                             QTableView,
                             QDesktopWidget,
                             QLabel,
                             qApp,
                             QLineEdit,
                             QPushButton,
                             QFileDialog,
                             QGridLayout,
                             QVBoxLayout,
                             QWidget,
                             QAction,
                             QMessageBox,
                             QGroupBox,
                             QTableWidget,
                             QTableWidgetItem,
                             QHeaderView)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

# from Parse_Variable import Get_Variable
from Get_Variable import Get_Variable
from Data_Exchange import Get_Data

class DataTransferWindow(QMainWindow):

    def __init__(self, parent=None):

        super(DataTransferWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(900, 600)
        self.setWindowTitle('Data Transfer')
        self.setWindowIcon(QIcon(".\icon\exchange.ico"))
        # root = QFileInfo(__file__).absolutePath()
        # self.setWindowIcon(QIcon(root + "./icon/Text_Edit.ico"))

        self.run_path = None
        self.lis_name = None
        self.job_path = None
        self.pst_path = None
        self.dll_path = None
        self.xml_path = None
        self.njl_path = None
        self.para_con = None
        self.othe_con = None

        self.initUI()
        self.load_setting()
        self.center()

    def initUI(self):

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        helpMenu = menubar.addMenu('Help')

        saveAction = QAction(QIcon('./Icon/save.ico'), 'Save', self)
        saveAction.setShortcut('Ctrl+S')
        saveAction.triggered.connect(self.save_setting)
        fileMenu.addAction(saveAction)

        clearAction = QAction(QIcon('Icon/clear.ico'), 'Clear', self)
        clearAction.setShortcut('Ctrl+R')
        clearAction.triggered.connect(self.clear_setting)
        fileMenu.addAction(clearAction)

        exitAction = QAction(QIcon('Icon/exit.ico'), 'Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.triggered.connect(qApp.quit)
        fileMenu.addAction(exitAction)

        helpAction = QAction(QIcon('Icon/doc.ico'), 'User Manual', self)
        helpAction.setShortcut('Ctrl+H')
        helpAction.triggered.connect(self.user_manual)
        helpMenu.addAction(helpAction)

        self.label1 = QLabel("Run Path")
        self.label1.setFont(QFont("Calibri", 10))

        self.line1 = QLineEdit()
        self.line1.setFont(QFont("Calibri", 10))
        self.line1.setPlaceholderText("Pick run path")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("Calibri", 10))
        self.btn1.clicked.connect(self.load_run_path)

        self.label2 = QLabel("Output Path")
        self.label2.setFont(QFont("Calibri", 10))

        self.line2 = QLineEdit()
        self.line2.setFont(QFont("Calibri", 10))
        self.line2.setPlaceholderText("Pick output path")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(QFont("Calibri", 10))
        self.btn2.clicked.connect(self.load_post_path)
        
        self.label3 = QLabel("Variable Path")
        self.label3.setFont(QFont("Calibri", 10))

        self.line3 = QLineEdit()
        self.line3.setFont(QFont("Calibri", 10))
        self.line3.setPlaceholderText("Choose variable path")

        self.btn3 = QPushButton('...')
        self.btn3.setFont(QFont("Calibri", 10))
        self.btn3.clicked.connect(self.load_variable_path)

        self.btn4 = QPushButton("Get Variable")
        self.btn4.setFont(QFont("Calibri", 10))
        self.btn4.clicked.connect(self.load_variable)

        self.btn5 = QPushButton("Data Transfer")
        self.btn5.setFont(QFont("Calibri", 10))
        self.btn5.clicked.connect(self.handle)
        self.btn5.setDisabled(True)

        self.table = QTableWidget(10,2)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        # self.table.setColumnWidth(0,200)
        # self.table.setHorizontalHeaderLabels(['File name', 'Variables', 'DLCs', 'Units'])
        self.table.setHorizontalHeaderLabels(['File name', 'Design load cases'])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableView.NoEditTriggers)

        # self.cbx = QCheckBox('File name')

        self.group1 = QGroupBox('Input/Output')
        self.group1.setFont(QFont("Calibri", 8))
        self.grid1 = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid1.addWidget(self.label1, 0, 0, 1, 1)
        self.grid1.addWidget(self.line1, 0, 1, 1, 5)
        self.grid1.addWidget(self.btn1, 0, 6, 1, 1)
        self.grid1.addWidget(self.label2, 1, 0, 1, 1)
        self.grid1.addWidget(self.line2, 1, 1, 1, 5)
        self.grid1.addWidget(self.btn2, 1, 6, 1, 1)
        self.grid1.addWidget(self.label3, 2, 0, 1, 1)
        self.grid1.addWidget(self.line3, 2, 1, 1, 5)
        self.grid1.addWidget(self.btn3, 2, 6, 1, 1)
        self.group1.setLayout(self.grid1)

        self.group2 = QGroupBox('Export information')
        self.group1.setFont(QFont("Calibri", 8))
        self.grid2 = QVBoxLayout()
        self.grid2.addWidget(self.table)
        self.group2.setLayout(self.grid2)

        # self.grid.setHorizontalSpacing(0)
        self.main_layout = QVBoxLayout()
        self.main_layout.addWidget(self.group1)
        # self.main_layout.addStretch(0)
        self.main_layout.addWidget(self.group2)
        self.main_layout.addWidget(self.btn4)
        self.main_layout.addWidget(self.btn5)
        self.mywidget.setLayout(self.main_layout)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def save_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if not config.has_section('Data Transfer'):
            config.add_section('Data Transfer')

        config['Data Transfer'] = {'Run path':self.line1.text(),
                                   'Result path':self.line2.text(),
                                   'Variable path':self.line3.text()}

        with open("config.ini",'w') as f:
            config.write(f)

    def load_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if config.has_section('Data Transfer'):
            self.line1.setText(config.get('Data Transfer','Run path'))
            # self.line1.home(True)
            self.line2.setText(config.get('Data Transfer','Result path'))
            # self.line2.home(True)
            self.line3.setText(config.get('Data Transfer','Variable path'))
            # self.line3.home(True)

    def clear_setting(self):
        self.line1.setText('')
        self.line2.setText('')
        self.line3.setText('')

    def user_manual(self):
        os.startfile(os.getcwd()+os.sep+'user manual\Data Transfer.docx')

    def load_variable_path(self):
        var_path, filetype = QFileDialog.getOpenFileName(self, "Choose joblist dialog", r".")

        if var_path:
            self.line3.setText(var_path.replace('/', '\\'))
            self.line3.home(True)

    def load_post_path(self):
        post_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose result path",
                                                         r".")

        if post_path:
            self.line2.setText(post_path.replace('/', '\\'))
            self.line2.home(True)

    def load_run_path(self):

        run_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose job path",
                                                         r".")

        if run_path:
            self.line1.setText(run_path.replace('/', '\\'))
            self.line1.home(True)

    def load_variable(self):
        for i in range(self.table.rowCount())[::-1]:
            self.table.removeRow(i)

        self.run_path = self.line1.text()
        self.var_path = self.line3.text()

        if self.var_path and self.run_path:

            try:
                self.result = Get_Variable(self.run_path, self.var_path)
                file_num    = 1

                for comp in self.result.component_list:

                    row_num = self.table.rowCount()

                    if file_num > row_num:
                        self.table.insertRow(row_num)

                    component = QTableWidgetItem(comp)
                    component.setCheckState(QtCore.Qt.Checked)

                    self.table.setItem(file_num-1, 0, component)

                    con = ','.join(self.result.component_dlcs[comp])
                    dlc = QTableWidgetItem(con)
                    # print(dlc, type(dlc))
                    self.table.setItem(file_num-1, 1, dlc)

                    file_num += 1

                # self.btn4.setDisabled(True)
                self.btn5.setDisabled(False)

            except Exception as e:
                QMessageBox.about(self, 'Warnning', '%s' %e)

        elif not self.run_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a run file first!')

        elif not self.var_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a variable file first!')

    def handle(self):
        self.pst_path = self.line2.text()
        comp_checked = []

        for i in range(self.table.rowCount()):

            item = self.table.item(i, 0)

            if item.checkState() == QtCore.Qt.Checked:
                comp_checked.append(item.text())

        if comp_checked and self.pst_path:

            try:
                for comp in comp_checked:

                    result_path = os.path.join(self.pst_path, comp)

                    dlc_list   = self.result.component_dlcs[comp]
                    unit_list  = self.result.component_unit[comp]
                    var_list   = self.result.component_vars[comp]
                    index_list = [self.result.var_ext_index[var][1] for var in var_list]
                    exten_list = [self.result.var_ext_index[var][0] for var in var_list]
                    comp_type  = self.result.component_type[comp]
                    step_num   = ''.join(list(filter(str.isdigit,comp_type[2]))) if len(comp_type)==3 else 1
                    trans_type = 'flex' if (len(comp_type)>1) and comp_type[0].lower().startswith('f') else 'bladed'
                    data_form  = 'binary' if (len(comp_type)>1) and comp_type[1].lower().startswith('b') else 'ascii'

                    if ('blade' in comp.lower()) and ('section' in comp.lower()):

                        if self.result.component_secs[comp][0].lower().startswith('a'):
                            sec_index = list(range(1, len(self.result.comp_section['blade'])+1))
                            sec_list  = self.result.comp_section['blade']
                        else:
                            sec_index = self.result.component_secs[comp]
                            sec_list  = [self.result.comp_section['blade'][int(i)-1] for i in sec_index]

                        Get_Data(step_num=step_num,
                                 transfer_type=trans_type,
                                 data_format=data_form,
                                 run_path=self.run_path,
                                 res_path=result_path,
                                 dlc_list=dlc_list,
                                 unit_list=unit_list,
                                 index_list=index_list,
                                 variable_list=var_list,
                                 extension_list=exten_list,
                                 node_list=sec_index,
                                 section_list=sec_list,
                                 matrix_type='blade').run()

                    elif ('tower' in comp.lower()) and ('section' in comp.lower()):

                        tower_type  = self.result.component_secs[comp][0]
                        model_type  = self.result.component_secs[comp][1]
                        print(tower_type, model_type)

                        # section
                        if tower_type.lower().startswith('s'):
                            matrix_type = 'b2t_onshore'
                            sec_index   = self.result.component_secs[comp][2]
                            sec_list    = [self.result.comp_section['tower_section'][int(i)-1] for i in sec_index]

                            Get_Data(step_num=step_num,
                                     transfer_type=trans_type,
                                     data_format=data_form,
                                     run_path=self.run_path,
                                     res_path=result_path,
                                     dlc_list=dlc_list,
                                     unit_list=unit_list,
                                     index_list=index_list,
                                     variable_list=var_list,
                                     extension_list=exten_list,
                                     node_list=sec_index,
                                     section_list=sec_list,
                                     matrix_type=matrix_type).run()
                        # mbr
                        elif tower_type.lower().startswith('m'):
                            if model_type.lower().startswith('b'):
                                matrix_type = 'b2t_offshore'
                                mbr_list    = self.result.component_secs[comp][2]
                                end_index   = [(int(mbr)*2-2) if mbr != mbr_list[-1] else (int(mbr)*2-1) for mbr in mbr_list]
                                bot_index   = self.result.comp_section['tower_node'][0][0]
                                tow_height  = self.result.comp_section['tower_node'][1]
                                sec_list    = [tow_height[int(i)-1]-tow_height[bot_index-1] for i in mbr_list]

                                Get_Data(step_num=step_num,
                                         transfer_type=trans_type,
                                         data_format=data_form,
                                         run_path=self.run_path,
                                         res_path=result_path,
                                         dlc_list=dlc_list,
                                         unit_list=unit_list,
                                         index_list=index_list,
                                         variable_list=var_list,
                                         extension_list=exten_list,
                                         node_list=end_index,
                                         section_list=sec_list,
                                         matrix_type=matrix_type).run()

                            elif model_type.lower().startswith('t'):
                                matrix_type = 't2b_offshore'
                                mbr_list    = self.result.component_secs[comp][2]
                                print(mbr_list)
                                end_list    = [(int(mbr)*2-2) if mbr == mbr_list[0] else (int(mbr)*2-1) for mbr in mbr_list]
                                print(end_list)
                                bot_index   = self.result.comp_section['tower_node'][0][1]
                                print(bot_index)
                                # tow_height  = self.result.comp_section['tower_node'][1]
                                # print(tow_height)
                                # sec_list    = [tow_height[int(i)-1]-tow_height[bot_index-1] for i in mbr_list]
                                # print(sec_list)
                                print(var_list)
                                print(index_list)

                                Get_Data(step_num=step_num,
                                         transfer_type=trans_type,
                                         data_format=data_form,
                                         run_path=self.run_path,
                                         res_path=result_path,
                                         dlc_list=dlc_list,
                                         unit_list=unit_list,
                                         index_list=index_list,
                                         variable_list=var_list,
                                         extension_list=exten_list,
                                         node_list=end_list,
                                         section_list=mbr_list,
                                         matrix_type=matrix_type).run()

                    else:

                        Get_Data(step_num=step_num,
                                 transfer_type=trans_type,
                                 data_format=data_form,
                                 run_path=self.run_path,
                                 res_path=result_path,
                                 dlc_list=dlc_list,
                                 unit_list=unit_list,
                                 index_list=index_list,
                                 variable_list=var_list,
                                 extension_list=exten_list,
                                 matrix_type='ones').run()

            except Exception as e:
                QMessageBox.about(self, 'Window', '%s' %e)
            finally:
                QMessageBox.about(self, 'Window', 'Data transfet is done!')

        else:
            QMessageBox.about(self, 'Window', 'Contents to modify are empty, nothing would be done!')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    app    = QApplication(sys.argv)
    window = DataTransferWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())
