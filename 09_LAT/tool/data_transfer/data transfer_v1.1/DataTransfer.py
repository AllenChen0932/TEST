#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/23/2020 9:32 AM
# @Author  : CE
# @File    : Data_Window.py

import sys, os
import configparser

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QTableView,
                             QDesktopWidget,
                             QLabel,
                             qApp,
                             QLineEdit,
                             QPushButton,
                             QFileDialog,
                             QGridLayout,
                             QVBoxLayout,
                             QHBoxLayout,
                             QWidget,
                             QAction,
                             QMessageBox,
                             QGroupBox,
                             QTableWidget,
                             QTableWidgetItem,
                             QHeaderView)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

import multiprocessing

# from Parse_Variable import Get_Variable
from tool.data_transfer.Get_Variable import Get_Variable
from tool.data_transfer.Data_Exchange import Get_Data

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

class DataTransferWindow(QMainWindow):

    def __init__(self, parent=None):

        super(DataTransferWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(900, 670)
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

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

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
        self.label1.setFont(self.cont_font)

        self.line1 = MyQLineEdit()
        self.line1.setFont(self.cont_font)
        self.line1.setPlaceholderText("Pick run path")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        self.btn1.clicked.connect(self.load_run_path)

        self.label2 = QLabel("Output Path")
        self.label2.setFont(self.cont_font)

        self.line2 = MyQLineEdit()
        self.line2.setFont(self.cont_font)
        self.line2.setPlaceholderText("Pick output path")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        self.btn2.clicked.connect(self.load_post_path)
        
        self.label3 = QLabel("Variable Path")
        self.label3.setFont(self.cont_font)

        self.line3 = MyQLineEdit()
        self.line3.setFont(self.cont_font)
        self.line3.setPlaceholderText("Choose variable path")
        self.line3.textChanged.connect(self.text_changed)

        self.btn3 = QPushButton('...')
        self.btn3.setFont(self.cont_font)
        self.btn3.clicked.connect(self.load_variable_path)

        self.btn4 = QPushButton("Get Variable")
        self.btn4.setFont(self.cont_font)
        self.btn4.clicked.connect(self.load_variable)

        self.btn5 = QPushButton("Data Transfer")
        self.btn5.setFont(self.cont_font)
        self.btn5.clicked.connect(self.handle)
        self.btn5.setDisabled(True)

        self.btn6 = QPushButton("Check Result")
        self.btn6.setFont(self.cont_font)
        self.btn6.clicked.connect(self.check_result)
        # self.btn6.setDisabled(True)

        self.btn7 = QPushButton('Deselect All Items')
        self.btn7.setFont(self.cont_font)
        self.btn7.clicked.connect(self.deselect_all)
        self.btn7.setDisabled(True)

        self.table = QTableWidget(10,2)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        # self.table.setColumnWidth(0,200)
        # self.table.setHorizontalHeaderLabels(['File name', 'Variables', 'DLCs', 'Units'])
        self.table.setHorizontalHeaderLabels(['File name', 'Design load cases'])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableView.NoEditTriggers)

        self.group1 = QGroupBox('Input/Output')
        self.group1.setFont(self.title_font)
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
        self.group2.setFont(self.title_font)
        self.grid2 = QVBoxLayout()
        self.grid2.addWidget(self.table)
        self.grid3 = QHBoxLayout()
        self.grid3.addWidget(self.btn4)
        self.grid3.addWidget(self.btn7)
        self.grid2.addLayout(self.grid3)

        self.group2.setLayout(self.grid2)

        # self.grid.setHorizontalSpacing(0)
        self.main_layout = QVBoxLayout()
        self.main_layout.addWidget(self.group1)
        # self.main_layout.addStretch(0)
        self.main_layout.addWidget(self.group2)
        # self.main_layout.addWidget(self.btn4)
        self.main_layout.addWidget(self.btn5)
        self.main_layout.addWidget(self.btn6)
        # self.main_layout.addStretch(0)
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

    def text_changed(self):
        self.table.clearContents()

    def load_variable(self):
        # reset variable setting
        # self.result   = None

        # self.table.clear()
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
                    self.table.setItem(file_num-1, 1, dlc)

                    file_num += 1

                # self.btn4.setDisabled(True)
                self.btn5.setDisabled(False)
                self.btn7.setDisabled(False)

            except Exception as e:
                QMessageBox.about(self, 'Warnning', '%s' %e)

        elif not self.run_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a run file first!')

        elif not self.var_path:
            QMessageBox.about(self, 'Warnning', 'Please choose a variable file first!')

    def deselect_all(self):

        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item.checkState() == QtCore.Qt.Checked:
                item.setCheckState(QtCore.Qt.Unchecked)

    def handle(self):

        self.pst_path = self.line2.text()
        comp_checked  = []

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
                            sec_list  = [self.result.comp_section['blade_all'][int(i)-1] for i in self.result.comp_section['blade_node']]
                            sec_index = [int(i)-1 for i in self.result.comp_section['blade_node']]
                            node_list = self.result.comp_section['blade_node']

                        else:
                            sec_list  = [self.result.comp_section['blade_all'][int(i)-1] for i in self.result.component_secs[comp]]
                            sec_index = [int(i)-1 for i in self.result.component_secs[comp]]
                            node_list = self.result.component_secs[comp]

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
                                 node_list=node_list,
                                 section_list=sec_list,
                                 section_index=sec_index,
                                 matrix_type='blade').run()

                    elif ('tower' in comp.lower()) and ('section' in comp.lower()):

                        tower_type  = self.result.component_secs[comp][0]
                        model_type  = self.result.component_secs[comp][1]
                        print(tower_type, model_type)

                        # section
                        if tower_type.lower().startswith('s'):

                            if self.result.component_secs[comp][1].startswith('t'):
                                raise Exception('Please define b2t for onshore wind turbine!')

                            matrix_type = 'b2t_onshore'
                            # according to 'TLOADS_ELS' list in $PJ
                            if not self.result.component_secs[comp][2][0].lower().startswith('a'):
                                node_list   = self.result.component_secs[comp][2]
                                sec_index   = [self.result.comp_section['tower_node'].index(i) for i in node_list]
                                sec_list    = [self.result.tower_node[i] for i in sec_index]
                            else:
                                node_list   = self.result.comp_section['tower_node']
                                sec_index   = [self.result.comp_section['tower_node'].index(i) for i in node_list]
                                sec_list    = [self.result.tower_node[i] for i in sec_index]

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
                                     node_list=node_list,
                                     section_list=sec_list,
                                     section_index=sec_index,
                                     matrix_type=matrix_type).run()
                        # mbr
                        elif tower_type.lower().startswith('m'):
                            tower_type = self.result.component_secs[comp][1].lower()
                            elem_list  = self.result.component_secs[comp][2]
                            node_list  = []
                            mbr_list   = []
                            mbr_index  = [self.result.comp_section['tower_node'].index(str(i)) for i in elem_list]
                            end_index  = []
                            print(tower_type)

                            # top to bottom
                            if tower_type.startswith('t'):
                                for i, mbr in enumerate(elem_list):
                                    mbr_list.append('Mbr %s End %s' % (mbr, ('1' if mbr == elem_list[0] else '2')))
                                    node_list.append(int(mbr) if mbr==elem_list[0] else int(mbr)+1)
                                    end_index.append(2*mbr_index[i] if mbr==elem_list[0] else 2*mbr_index[i]+1)
                            # bottom to top
                            elif tower_type.startswith('b'):
                                for i, mbr in enumerate(elem_list):
                                    mbr_list.append('Mbr %s End %s' % (mbr, ('2' if mbr == elem_list[-1] else '1')))
                                    node_list.append(int(mbr)-int(elem_list[0])+2 if mbr==elem_list[-1] else int(mbr)-int(elem_list[0])+1)
                                    end_index.append(2*mbr_index[i]+1 if mbr==elem_list[-1] else 2*mbr_index[i])

                            if model_type.lower().startswith('b'):
                                matrix_type = 'b2t_offshore'
                                # end_index   = [(int(mbr)*2-2) if mbr != mbr_list[-1] else (int(mbr)*2-1) for mbr in mbr_list]
                                # bot_index   = self.result.comp_section['tower_node'][0][0]
                                # tow_height  = self.result.comp_section['tower_node'][1]
                                # sec_list    = [tow_height[int(i)-1]-tow_height[bot_index-1] for i in mbr_list]

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
                                         node_list=node_list,
                                         section_list=mbr_list,
                                         section_index=end_index,
                                         matrix_type=matrix_type).run()

                            elif model_type.lower().startswith('t'):
                                matrix_type = 't2b_offshore'
                                # mbr_list    = self.result.component_secs[comp][2]
                                # end_list    = [(int(mbr)*2-2) if mbr == mbr_list[0] else (int(mbr)*2-1) for mbr in mbr_list]
                                # bot_index   = self.result.comp_section['tower_node'][0][1]
                                # tow_height  = self.result.comp_section['tower_node'][1]
                                # sec_list    = [tow_height[int(i)-1]-tow_height[bot_index-1] for i in mbr_list]

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
                                         node_list=node_list,
                                         section_list=mbr_list,
                                         section_index=end_index,
                                         matrix_type=matrix_type).run()

                    else:
                        # if comp in self.result.component_secs.keys():
                        #     node_list = self.result.component_secs[comp]
                        # else:
                        #     node_list = None
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

                QMessageBox.about(self, 'Window', 'Data transfer is done!')
                self.btn6.setDisabled(False)

            except Exception as e:
                QMessageBox.about(self, 'Window', '%s' %e)

        else:
            QMessageBox.about(self, 'Window', 'Contents to modify are empty, nothing would be done!')

    def check_result(self):
        try:
            content = ''

            # get load case number for each DLC
            dlc_list = [dlc for dlc in os.listdir(self.run_path) if os.path.isdir(os.path.join(self.run_path, dlc))]

            if not dlc_list:
                raise Exception('No results in %s!' %self.run_path)

            dlc_case = {}
            for dlc in dlc_list:
                dlc_path = os.path.join(self.run_path, dlc)
                lc_list  = [lc for lc in os.listdir(dlc_path) if os.path.isdir(os.path.join(dlc_path, lc))]

                dlc_case[dlc] = len(lc_list)

            pst_path = self.line2.text()

            comp_checked = []

            for i in range(self.table.rowCount()):
                item = self.table.item(i, 0)

                if item.checkState() == QtCore.Qt.Checked:
                    comp_checked.append(item.text())

            comp_flag = {}
            for component in comp_checked:
                # print(content)
                comp_path = os.path.join(pst_path, component)
                res_file  = [file for file in os.listdir(comp_path) if 'BLADED' in file or 'FLEX' in file][0]
                ts_path   = os.path.join(comp_path, res_file)
                # print(comp_path, ts_path)

                if os.path.exists(ts_path):
                    DLC_list = self.result.component_dlcs[component]
                    change_flag = True

                    for DLC in DLC_list:
                        lc_num = len(os.listdir(os.path.join(ts_path, DLC)))

                        if lc_num != dlc_case[DLC]:
                            content += '%s_%s: %s/%s\n' %(component, DLC, lc_num, dlc_case[DLC])
                            change_flag = False
                        else:
                            self.result.component_dlcs[component].remove(DLC)

                    comp_flag[component] = change_flag
                else:
                    content += '%s: No results\n' %component

            for i in range(self.table.rowCount()):
                item = self.table.item(i, 0)
                # print(item.text())
                comp_name = item.text()
                if comp_name in comp_flag.keys():
                    if comp_flag[item.text()]:
                        item.setCheckState(QtCore.Qt.Unchecked)

            if content:
                QMessageBox.about(self, 'Window', '%s' %content)
            else:
                QMessageBox.about(self, 'Window', 'Data transfer successfully!')
        except Exception as e:
            QMessageBox.about(self, 'Warnning', 'Error occurs: %s' %e)

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    multiprocessing.freeze_support()
    app    = QApplication(sys.argv)
    window = DataTransferWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())
