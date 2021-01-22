# -*- coding: utf-8 -*-
# @Time    : 2020/11/23 15:51
# @Author  : CE
# @File    : LoadAssistantTool_v3.py

# -*- coding: utf-8 -*-
from random import randint
import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtWidgets

import JoblistModify_v2 as joblist


top=['Menu', 'DLC Generator', 'Joblist Generator', 'Check Simulations', 'Post Processing', 'Load Summary Table',
     'Components Load', 'Load Report']
tab=[['Load Case Table', 'One-Click', 'By Joblist'],
     ['Extrapolation', 'Combination', 'Post Assistant'],
     ['One-click', 'Data Transfer', 'Post Export']]

index_list = []
hide_menu = ['DLC Generator','Post Processing','Components Load']

for menu in top:
    index_list.append(menu)
    if menu in hide_menu:
        sub_menu = tab[hide_menu.index(menu)]
        for sub in sub_menu:
            index_list.append(sub)
print(index_list)

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


class JoblistModifyThread(QThread):
    def __init__(self, *args):
        super(JoblistModifyThread,self).__init__()
        self.args = args
        print(self.args)

    def run(self):
        self.func = joblist.modify_joblist(self.args[0],self.args[1],self.args[2])
        self.func.get_jobpath()
        self.func.create_jobs()


class LeftWidget(QWidget):
    update_ = pyqtSignal(str)
    def __init__(self, item, factor, parent=None):
        super(LeftWidget, self).__init__(parent)
        self.item = item
        layout = QFormLayout(self)

        for name in factor:
            button = QPushButton(name)
            layout.addRow(button)
            button.clicked.connect(self.onClick)

    def onClick(self):
        txt = self.sender().text()#获取发送信号的控件文本
        self.update_.emit(txt)

    def resizeEvent(self, event):
        # 解决item的高度问题
        super(LeftWidget, self).resizeEvent(event)
        self.item.setSizeHint(QSize(self.minimumWidth(), self.height()))


class TabButton(QPushButton):
    # 按钮作为开关
    update_ = pyqtSignal(str)
    def __init__(self, item,name,parent=None):
        super(TabButton, self).__init__(parent)
        self.item = item
        # self.setCheckable(True)  # 设置可选中
        self.setText(name)

    def onClick(self):
        txt = self.sender().text()#获取发送信号的控件文本
        self.update_.emit(txt)

    def resizeEvent(self, event):
        # 解决item的高度问题
        super(TabButton, self).resizeEvent(event)
        self.item.setSizeHint(QSize(self.minimumWidth(), self.height()))


class LeftWindow(QListWidget):
    '''left window'''
    def __init__(self, *args, **kwargs):
        super(LeftWindow, self).__init__(*args, **kwargs)
        # menu button
        menu = QListWidgetItem(self)
        self.menu_btn = TabButton(menu, top[0])
        self.setItemWidget(menu, self.menu_btn)

        # dlc button
        dlc = QListWidgetItem(self)
        self.dlc_btn = TabButton(dlc,top[1])
        self.setItemWidget(dlc, self.dlc_btn)
        # hidden item
        dlc_item = QListWidgetItem(self)
        # show/hide items through button click
        self.dlc_btn.clicked.connect(lambda :self.set_hidden(dlc_item))
        self.dlc_widget = LeftWidget(dlc_item, tab[0])
        self.setItemWidget(dlc_item,self.dlc_widget)
        # dlc_item.setHidden(True)

        # joblist button
        joblist = QListWidgetItem(self)
        self.jbl_btn = TabButton(joblist, top[2])
        self.setItemWidget(joblist, self.jbl_btn)

        # check button
        check = QListWidgetItem(self)
        self.check_btn = TabButton(check, top[3])
        self.setItemWidget(check, self.check_btn)

        # post button
        post = QListWidgetItem(self)
        self.post_btn = TabButton(post, top[4])
        self.setItemWidget(post, self.post_btn)
        # hidden item
        post_item = QListWidgetItem(self)
        # show/hide items through button click
        self.post_btn.clicked.connect(lambda :self.set_hidden(post_item))
        self.post_widget = LeftWidget(post_item, tab[1])
        self.setItemWidget(post_item,self.post_widget)
        # post_item.setHidden(True)

        # load summary button
        load = QListWidgetItem(self)
        self.lst_btn = TabButton(load, top[5])
        self.setItemWidget(load, self.lst_btn)

        # component button
        comp = QListWidgetItem(self)
        self.comp_btn = TabButton(comp, top[6])
        self.setItemWidget(comp, self.comp_btn)
        # hidden item
        comp_item = QListWidgetItem(self)
        # show/hide items through button click
        self.comp_btn.clicked.connect(lambda :self.set_hidden(comp_item))
        self.comp_widget = LeftWidget(comp_item, tab[2])
        self.setItemWidget(comp_item,self.comp_widget)
        # comp_item.setHidden(True)

        # load report button
        load = QListWidgetItem(self)
        self.load_btn = TabButton(load, top[7])
        self.setItemWidget(load, self.load_btn)

    def set_hidden(self, item):
        '''set hidden state for QListWidgetItem'''
        if not item.isHidden():
            item.setHidden(True)
        else:
            item.setHidden(False)


class JoblistModifyWidget(QWidget):

    def __init__(self, parent=None):
        super(JoblistModifyWidget, self).__init__(parent)

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.joblist()

    def joblist(self):
        # ********* joblist parameters group *********
        self.label1 = QLabel("Origninal joblist")
        self.label1.setFont(self.cont_font)

        self.line1 = MyQLineEdit()
        self.line1.setFont(self.cont_font)

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        # self.btn1.clicked.connect(self.load_job_list)

        self.group1 = QGroupBox('Joblist parameters')
        self.group1.setFont(self.title_font)

        self.grid1 = QGridLayout(self.group1)
        self.grid1.addWidget(self.label1, 0, 0, 1, 1)
        self.grid1.addWidget(self.line1, 0, 1, 1, 5)
        self.grid1.addWidget(self.btn1, 0, 6, 1, 1)

        # result parameters group
        self.label2 = QLabel("New result path")
        self.label2.setFont(self.cont_font)

        self.line2 = MyQLineEdit()
        self.line2.setFont(self.cont_font)
        self.line2.setToolTip('new run path for the new joblist')

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        # self.btn2.clicked.connect(self.load_res_path)

        self.label3 = QLabel("New joblist name")
        self.label3.setFont(self.cont_font)

        self.line3 = MyQLineEdit()
        self.line3.setFont(self.cont_font)
        self.line3.setToolTip('new joblist name, default as the old one')

        self.group2 = QGroupBox('Result parameters')
        self.group2.setFont(self.title_font)

        self.grid2 = QGridLayout(self.group2)
        self.grid2.addWidget(self.label2, 1, 0, 1, 1)
        self.grid2.addWidget(self.line2, 1, 1, 1, 5)
        self.grid2.addWidget(self.btn2, 1, 6, 1, 1)
        self.grid2.addWidget(self.label3, 2, 0, 1, 1)
        self.grid2.addWidget(self.line3, 2, 1, 1, 5)

        # ********* project parameters group *********
        # controller parameters
        self.label5 = QLabel("DLL Path")
        self.label5.setFont(self.cont_font)

        self.line5 = MyQLineEdit()
        self.line5.setFont(self.cont_font)
        # self.line5.setPlaceholderText("Pick DLL file")

        self.btn5 = QPushButton("...")
        self.btn5.setFont(self.cont_font)
        # self.btn5.clicked.connect(self.load_dll_path)

        self.label6 = QLabel("XML Path")
        self.label6.setFont(self.cont_font)

        self.line6 = MyQLineEdit()
        self.line6.setFont(self.cont_font)
        # self.line6.setPlaceholderText("Pick XML file")

        self.btn6 = QPushButton("...")
        self.btn6.setFont(self.cont_font)
        # self.btn6.clicked.connect(self.load_xml_path)

        self.sub1 = QGroupBox('Controller')
        self.sub1.setFont(self.title_font)

        self.grid3 = QGridLayout(self.sub1)
        self.grid3.addWidget(self.label5, 0, 0, 1, 1)
        self.grid3.addWidget(self.line5, 0, 1, 1, 5)
        self.grid3.addWidget(self.btn5, 0, 6, 1, 1)
        self.grid3.addWidget(self.label6, 1, 0, 1, 1)
        self.grid3.addWidget(self.line6, 1, 1, 1, 5)
        self.grid3.addWidget(self.btn6, 1, 6, 1, 1)
        self.sub1.setLayout(self.grid3)

        # controller parameters
        self.text1 = QTextEdit()
        self.text1.setFont(self.cont_font)
        self.text1.setPlaceholderText("Additional controller parameters")

        self.sub2 = QGroupBox('Add/Modify controller parameters')
        self.sub2.setFont(self.title_font)
        self.vbox2 = QVBoxLayout(self.sub2)
        self.vbox2.addWidget(self.text1)

        # replacement
        self.text2 = QTextEdit()
        self.text2.setFont(self.cont_font)
        self.text2.setPlaceholderText("format: old, new\n"
                                      "e.g. WDIR	 0, WDIR	 1")

        self.sub3 = QGroupBox('Replacement')
        self.sub3.setFont(self.title_font)
        self.vbox3= QVBoxLayout(self.sub3)
        self.vbox3.addWidget(self.text2)

        self.btn8 = QPushButton("Create jobs and joblist")
        self.btn8.setFont(self.cont_font)
        # self.btn8.clicked.connect(self.run_joblist_modify)

        self.group3 = QGroupBox('Project parameters')
        self.group3.setFont(self.title_font)

        self.hbox = QHBoxLayout()
        self.hbox.addWidget(self.sub2)
        self.hbox.addWidget(self.sub3)
        self.vbox = QVBoxLayout(self.group3)
        self.vbox.addWidget(self.sub1)
        self.vbox.addLayout(self.hbox)

        self.text3 = QTextEdit()
        self.text3.setFont(self.cont_font)
        self.text3.setPlaceholderText("output info...")

        self.label0 = QLabel('Joblist Modification')
        self.label0.setFont(self.title_font)

        self.group0 = QGroupBox()
        self.group0.setStyleSheet("QGroupBox{border:None}")
        self.vbox0 = QHBoxLayout(self.group0)
        self.vbox0.addStretch()
        self.vbox0.addWidget(self.label0)
        self.vbox0.addStretch()

        self.main_layout = QVBoxLayout()
        # self.main_layout.addWidget(self.group0)
        self.main_layout.addWidget(self.group1)
        self.main_layout.addWidget(self.group2)
        self.main_layout.addWidget(self.group3)
        self.main_layout.addWidget(self.btn8)
        self.main_layout.addWidget(self.text3)
        self.main_layout.addStretch(1)

        self.setLayout(self.main_layout)


class DataTransferWidget(QWidget):
    def __init__(self, parent=None):
        super(DataTransferWidget, self).__init__(parent)

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.data_transfer()

    def data_transfer(self):
        self.label1 = QLabel("Run Path")
        self.label1.setFont(self.cont_font)

        self.line1 = MyQLineEdit()
        self.line1.setFont(self.cont_font)
        self.line1.setPlaceholderText("Pick run path")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        # self.btn1.clicked.connect(self.load_run_path)

        self.label2 = QLabel("Output Path")
        self.label2.setFont(self.cont_font)

        self.line2 = MyQLineEdit()
        self.line2.setFont(self.cont_font)
        self.line2.setPlaceholderText("Pick output path")

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        # self.btn2.clicked.connect(self.load_post_path)

        self.label3 = QLabel("Variable Path")
        self.label3.setFont(self.cont_font)

        self.line3 = MyQLineEdit()
        self.line3.setFont(self.cont_font)
        self.line3.setPlaceholderText("Choose variable path")
        # self.line3.textChanged.connect(self.text_changed)

        self.btn3 = QPushButton('...')
        self.btn3.setFont(self.cont_font)
        # self.btn3.clicked.connect(self.load_variable_path)

        self.btn4 = QPushButton("Get Variable")
        self.btn4.setFont(self.cont_font)
        # self.btn4.clicked.connect(self.load_variable)

        self.btn5 = QPushButton("Data Transfer")
        self.btn5.setFont(self.cont_font)
        # self.btn5.clicked.connect(self.handle)
        self.btn5.setDisabled(True)

        self.btn6 = QPushButton("Check Result")
        self.btn6.setFont(self.cont_font)
        # self.btn6.clicked.connect(self.check_result)
        # self.btn6.setDisabled(True)

        self.btn7 = QPushButton('Deselect All Items')
        self.btn7.setFont(self.cont_font)
        # self.btn7.clicked.connect(self.deselect_all)
        self.btn7.setDisabled(True)

        self.table = QTableWidget(10, 2)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
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
        self.setLayout(self.main_layout)


class JoblistGeneratorWidget(QWidget):
    def __init__(self, parent=None):
        super(JoblistGeneratorWidget, self).__init__(parent)

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.data_transfer()

    def data_transfer(self):

        self.group1 = QGroupBox('Generate Joblist from IN')
        self.grid = QGridLayout(self.group1)

        self.lbl1 = QLabel("Run/Post Path")
        self.lbl1.setFont(self.cont_font)

        self.lbl2 = QLabel("Joblist Path")
        self.lbl2.setFont(self.cont_font)

        self.lbl3 = QLabel("Joblist Name")
        self.lbl3.setFont(self.cont_font)

        self.lbl4 = QLabel('Joblist Type')
        self.lbl4.setFont(self.cont_font)

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(self.cont_font)
        self.lin1.setPlaceholderText("Choose run or post path")

        self.lin2 = MyQLineEdit()
        self.lin2.setFont(self.cont_font)
        self.lin2.setPlaceholderText("Choose joblist path")

        self.lin3 = QLineEdit()
        self.lin3.setFont(self.cont_font)
        # self.lin3.setDisabled(True)
        self.lin3.setPlaceholderText("Define joblist name")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        # self.btn1.clicked.connect(self.choose_run_path)

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        # self.btn2.clicked.connect(self.choose_joblist_path)

        self.btn3 = QPushButton("Generate Joblist")
        self.btn3.setFont(self.cont_font)
        # self.btn3.clicked.connect(self.generate_joblist)

        self.lbl5 = QLabel('DLC list')
        self.lbl5.setFont(self.cont_font)
        self.tbl1 = QTableWidget(0, 2)
        self.tbl1.setDisabled(True)
        self.tbl1.setFont(self.cont_font)
        self.tbl1.setHorizontalHeaderLabels(['Ultimate Load Case', 'Fatigue Load Case'])

        self.tbl1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # self.tbl1.horizontalHeader().setSectionResizeMode(0,QHeaderView.ResizeToContents)
        self.tbl1.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # self.tbl1.verticalHeader().setVisible(False)
        self.tbl1.setEditTriggers(QTableView.NoEditTriggers)

        self.btn4 = QPushButton('Get DLC')
        self.btn4.setDisabled(True)
        self.btn4.setFont(self.cont_font)
        # self.btn4.clicked.connect(self.get_dlc)

        self.cbx1 = QComboBox()
        self.cbx1.setFont(self.cont_font)
        # self.cbx1.setDisabled(True)
        # self.cbx1.currentTextChanged.connect(self.type_action)
        self.cbx1.addItem("Please select")
        self.cbx1.addItem("Run")
        self.cbx1.addItem("Post")

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 3, 0, 1, 1)
        self.grid.addWidget(self.lin1, 3, 1, 1, 4)
        self.grid.addWidget(self.btn1, 3, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 2, 0, 1, 1)
        self.grid.addWidget(self.lin2, 2, 1, 1, 4)
        self.grid.addWidget(self.btn2, 2, 5, 1, 1)
        self.grid.addWidget(self.lbl3, 1, 0, 1, 1)
        self.grid.addWidget(self.lin3, 1, 1, 1, 4)
        self.grid.addWidget(self.lbl4, 0, 0, 1, 1)
        self.grid.addWidget(self.cbx1, 0, 1, 1, 4)
        self.grid.addWidget(self.lbl5, 4, 0, 1, 1)
        self.grid.addWidget(self.tbl1, 4, 1, 10, 4)
        self.grid.addWidget(self.btn4, 4, 5, 1, 1)
        self.grid.addWidget(self.btn3, 14, 0, 1, 6)

        self.group2 = QGroupBox('Trim Joblist')
        self.grid2 = QGridLayout(self.group2)

        self.lbl21 = QLabel("Original Joblist")
        self.lbl21.setFont(QFont("Calibri", 9))

        self.lbl22 = QLabel("New Joblist Name")
        self.lbl22.setFont(QFont("Calibri", 9))

        self.lbl23 = QLabel("Joblist Type")
        self.lbl23.setFont(QFont("Calibri", 9))

        self.lin21 = MyQLineEdit()
        self.lin21.setFont(QFont("Calibri", 9))
        self.lin21.setPlaceholderText("Choose original joblist")

        self.lin22 = QLineEdit()
        self.lin22.setFont(QFont("Calibri", 9))
        self.lin22.setPlaceholderText("Define new joblist name")

        self.lin23 = QLineEdit()
        self.lin23.setFont(QFont("Calibri", 9))
        self.lin23.setDisabled(True)
        self.lin23.setPlaceholderText("DLC list, separated by ','")

        self.btn21 = QPushButton("...")
        self.btn21.setFont(QFont("Calibri", 9))
        # self.btn21.clicked.connect(self.load_joblist)

        self.btn22 = QPushButton("Run")
        self.btn22.setFont(QFont("Calibri", 9))
        # self.btn22.clicked.connect(self.run)
        #
        self.cbx21 = QComboBox()
        self.cbx21.setFont(QFont("Calibri", 9))
        # self.cbx1.setDisabled(True)
        # self.cbx21.currentTextChanged.connect(self.type_action)
        self.cbx21.addItem("Please select")
        self.cbx21.addItem("Fatigue")
        self.cbx21.addItem("Ultimate")
        self.cbx21.addItem("User define")

        # 起始行，起始列，占用行，占用列
        self.grid2.addWidget(self.lbl21, 0, 0, 1, 1)
        self.grid2.addWidget(self.lin21, 0, 1, 1, 4)
        self.grid2.addWidget(self.btn21, 0, 5, 1, 1)
        self.grid2.addWidget(self.lbl22, 1, 0, 1, 1)
        self.grid2.addWidget(self.lin22, 1, 1, 1, 4)
        self.grid2.addWidget(self.lbl23, 2, 0, 1, 1)
        self.grid2.addWidget(self.cbx21, 2, 1, 1, 4)
        self.grid2.addWidget(self.lin23, 3, 1, 1, 4)
        self.grid2.addWidget(self.btn22, 4, 0, 1, 6)

        vbox = QVBoxLayout()
        vbox.addWidget(self.group1)
        vbox.addWidget(self.group2)
        vbox.addStretch()
        self.setLayout(vbox)


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.setWindowTitle('Load Assistant Tool_v3')
        self.setWindowIcon(QIcon('icon/program1.ico'))
        self.setAutoFillBackground(True)

        self.title_font = QFont("Calibri", 10)
        self.title_font.setItalic(True)
        self.title_font.setBold(True)

        self.cont_font = QFont("Calibri", 9)
        self.cont_font.setItalic(False)
        self.cont_font.setBold(False)

        self.setupUI()

    def setupUI(self):

        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        toolMenu = menubar.addMenu('Tool')
        helpMenu = menubar.addMenu('Help')

        exitAction = QAction(QIcon('Icon/exit.ico'), 'Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        # exitAction.triggered.connect(qApp.quit)
        fileMenu.addAction(exitAction)

        toolAction = QAction(QIcon('Icon/backup.ico'), 'One click backup', self)
        # toolAction.triggered.connect(self.back_up)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(all)', self)
        # toolAction.triggered.connect(self.delete_run)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(except pj/in)', self)
        # toolAction.triggered.connect(self.delete_run_except_pj)
        toolMenu.addAction(toolAction)

        toolAction = QAction(QIcon('Icon/delete_file.ico'), 'Delete run(extension)', self)
        # toolAction.triggered.connect(self.delete_run_extension)
        toolMenu.addAction(toolAction)

        # toolAction = QAction(QIcon('Icon/copy.ico'), 'Copy', self)
        # toolAction.triggered.connect(self.copy_file)
        # toolMenu.addAction(toolAction)

        helpAction = QAction(QIcon('Icon/doc.ico'), 'User Manual', self)
        helpAction.setShortcut('Ctrl+H')
        # helpAction.triggered.connect(self.user_manual)
        helpMenu.addAction(helpAction)

        # self.toolbar = self.addToolBar('Exit')
        # self.toolbar.addAction(QIcon('Icon/exit.ico'), 'Exit')

        _backup = self.addToolBar('Backup')
        _backup.addAction(QAction(QIcon('Icon/backup.ico'), 'Backup', self))
        _backup.setIconSize(QSize(18,18))

        _delete = self.addToolBar('Delete')
        _delete.addAction(QIcon('Icon/delete_file.ico'), 'Delete')
        _delete.setIconSize(QSize(18,18))

        # toolbar = QtWidgets.QToolBar(self)
        # toolbar.setStyleSheet("QToolBar{spacing:8px")
        # toolbar.setIconSize(QSize=(18,18))
        # print(toolbar.iconSize())
        # _delete.setIconSize(QSize=(12,12))

        self.available_geometry = QDesktopWidget().availableGeometry()
        # init_width = self.available_geometry.width() * 0.85
        # init_height = self.available_geometry.height() * 0.85
        # self.setWindowTitle(app_name)
        # self.resize(init_width, init_height)
        self.resize(1200, 800)


        # 实例化状态栏，设置状态栏
        self.statusBar().showMessage('Ready')
        # self.statusBar = QStatusBar()
        # self.setStatusBar(self.statusBar)


        ###### 创建界面 ######
        self.centralwidget = QWidget()
        self.setCentralWidget(self.centralwidget)
        self.mainLayout = QHBoxLayout(self.centralwidget)#全局横向

        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.mainLayout.setSpacing(0)  # 去除控件间的间隙
        #################
        self.listWidget = LeftWindow()  # 左侧列表
        self.listWidget.setMaximumWidth(150)
        self.listWidget.setMinimumWidth(150)
        # 去掉边框
        self.listWidget.setFrameShape(QListWidget.NoFrame)
        # 隐藏滚动条
        self.listWidget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.listWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.StackDataDisplay = QStackedWidget()  # 右侧层叠窗口
        # welcome menu
        label = QLabel('Welcome to use Load Assistant Tool', self)
        label.setFont(QFont("Calibri", 32))
        label.setStyleSheet('background: rgb(%d, %d, %d);margin: 1px;'%(
            randint(0, 255), randint(0, 255), randint(0, 255)))
        label.setAlignment(Qt.AlignTop)
        label.setAlignment(Qt.AlignCenter)
        self.StackDataDisplay.addWidget(label)

        # DLC generator menu
        label = QLabel('Welcome to use DLC Generator', self)
        label.setFont(QFont("Calibri", 32))
        label.setStyleSheet('background: rgb(%d, %d, %d);margin: 1px;'%(
            randint(0, 255), randint(0, 255), randint(0, 255)))
        self.StackDataDisplay.addWidget(label)
        label.setAlignment(Qt.AlignHCenter)

        # load case table tab
        label = QLabel('Welcome to use Load Case Table', self)
        label.setFont(QFont("Calibri", 32))
        label.setStyleSheet('background: rgb(%d, %d, %d);margin: 1px;'%(
            randint(0, 255), randint(0, 255), randint(0, 255)))
        self.StackDataDisplay.addWidget(label)
        label.setAlignment(Qt.AlignCenter)

        # one-click tab
        label = QLabel('Welcome to use One-Click', self)
        label.setFont(QFont("Calibri", 32))
        label.setStyleSheet('background: rgb(%d, %d, %d);margin: 1px;'%(
            randint(0, 255), randint(0, 255), randint(0, 255)))
        self.StackDataDisplay.addWidget(label)
        label.setAlignment(Qt.AlignCenter)

        # joblist tab
        self.StackDataDisplay.addWidget(JoblistModifyWidget())

        # joblist generator
        self.StackDataDisplay.addWidget(JoblistGeneratorWidget())

        # data transfer tab
        self.StackDataDisplay.addWidget(DataTransferWidget())


        # change page for each menu and tab
        self.listWidget.menu_btn.clicked.connect(lambda :self.update_btn(self.listWidget.menu_btn))
        self.listWidget.dlc_btn.clicked.connect(lambda :self.update_btn(self.listWidget.dlc_btn))
        self.listWidget.dlc_widget.update_.connect(self.update_tab)
        self.listWidget.jbl_btn.clicked.connect(lambda :self.update_btn(self.listWidget.jbl_btn))
        self.listWidget.check_btn.clicked.connect(lambda :self.update_btn(self.listWidget.check_btn))
        self.listWidget.post_btn.clicked.connect(lambda :self.update_btn(self.listWidget.post_btn))
        self.listWidget.post_widget.update_.connect(self.update_tab)
        self.listWidget.lst_btn.clicked.connect(lambda :self.update_btn(self.listWidget.load_btn))
        self.listWidget.comp_btn.clicked.connect(lambda :self.update_btn(self.listWidget.load_btn))
        self.listWidget.comp_widget.update_.connect(self.update_tab)
        self.listWidget.load_btn.clicked.connect(lambda :self.update_btn(self.listWidget.load_btn))

        ###################
        self.mainLayout.addWidget(self.listWidget)
        self.mainLayout.addWidget(self.StackDataDisplay)

    def update_btn(self, btn):
        # print(self.StackDataDisplay.CurrentIndex)
        # self.StackDataDisplay.setCurrentIndex(index_list.index(text))#根据文本设置不同的页面
        self.statusBar().showMessage('current module: %s' %btn.text())
        self.StackDataDisplay.setCurrentIndex(index_list.index(btn.text()))

    def update_tab(self, text):
        # print(self.StackDataDisplay.CurrentIndex)
        # self.StackDataDisplay.setCurrentIndex(index_list.index(text))#根据文本设置不同的页面
        self.statusBar().showMessage('current module: %s' %text)
        self.StackDataDisplay.setCurrentIndex(index_list.index(text))

    def run_joblist_modify(self):
        job_list = self.line1.text().replace('\\', '/')
        dll_path = self.line5.text().replace('\\', '/')
        xml_path = self.line6.text().replace('\\', '/')
        self.thread1 = JoblistModifyThread(job_list,dll_path,xml_path)
        self.thread1.start()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())