# -*- coding: utf-8 -*-
# @Time    : 2020/5/1 22:04
# @Author  : CE
# @File    : Post_UI_v2.0.py

import sys, os, time
import shutil
import configparser

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QRadioButton,
                             QListWidget,
                             QListWidgetItem,
                             QDesktopWidget,
                             QButtonGroup,
                             QCheckBox,
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
                             QGroupBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

from tool.post_assistant.Write_Joblist import Create_Joblist
from tool.post_assistant.Get_Variable import Get_Variables
from tool.post_assistant.Get_LoadCase import Get_DLC
from tool.post_assistant.Write_Ultimate import Ultimate
from tool.post_assistant.Write_Fatigue import Fatigue
from tool.post_assistant.Write_Bstats import Bstats
from tool.post_assistant.Write_Probability import Probability
from tool.post_assistant.Write_Steadytc import Steadytc
from tool.post_assistant.Write_Combination import Combination

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

class PostWindow(QMainWindow):

    def __init__(self, parent=None):

        super(PostWindow, self).__init__(parent)
        self.mywidget = QWidget(self)
        # self.mywidget.setFont(QFont("Calibri", 6))

        self.resize(850,600)
        self.setWindowTitle('Post Assistant')
        self.setWindowIcon(QIcon(".\icon\Text_Edit.ico"))
        # root = QFileInfo(__file__).absolutePath()
        # self.setWindowIcon(QIcon(root + "./icon/Text_Edit.ico"))

        # font
        self.title_font = QFont("Microsoft YaHei", 7)
        self.title_font.setItalic(True)

        self.cont_font = QFont("Microsoft YaHei", 7)
        # component under groupbox will be changed to italic
        self.cont_font.setItalic(False)

        self.proj_path = None
        self.post_path = None
        self.run_path  = None
        self.lc_table  = None
        self.sf_flag   = "2"      # IEC61400-1 ed4 in default
        self.etm_opt   = False

        @property
        def run_path(self):
            return self._run_path

        @run_path.setter
        def run_path(self, run_path):
            if not os.path.isdir(run_path):
                raise ValueError('Not valid post path!')
            else:
                self._run_path = run_path

        self.var_result = None

        self.initUI()
        self.load_setting()
        self.center()

    def initUI(self):

        # *************** menu bar ***************
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

        # *************** main panel ***************
        # Input and output group
        self.group1 = QGroupBox('Input and Output')
        self.group1.setFont(self.title_font)
        #
        self.lin1 = MyQLineEdit()
        self.lin1.setFont(self.cont_font)
        self.lin1.textChanged.connect(self.clear_var_list)
        self.btn1 = QPushButton("Prj Path")
        self.btn1.setFont(self.cont_font)
        self.btn1.clicked.connect(self.load_project)

        self.lin2 = MyQLineEdit()
        self.lin2.setFont(self.cont_font)
        self.btn2 = QPushButton("LC Table")
        self.btn2.setFont(self.cont_font)
        self.btn2.clicked.connect(self.load_lct)

        self.lin3 = MyQLineEdit()
        self.lin3.setFont(self.cont_font)
        self.lin3.textChanged.connect(self.clear_dlc_list)
        self.btn3 = QPushButton("Run Path")
        self.btn3.setFont(self.cont_font)
        self.btn3.clicked.connect(self.load_run_path)

        self.lin4 = MyQLineEdit()
        self.lin4.setFont(self.cont_font)
        self.btn4 = QPushButton("Post Path")
        self.btn4.setFont(self.cont_font)
        self.btn4.clicked.connect(self.load_post_path)

        self.lin15 = QLineEdit()
        self.lin15.setFont(self.cont_font)
        self.btn15 = QPushButton('Batch Path')
        self.btn15.setFont(self.cont_font)
        self.btn15.clicked.connect(self.load_batch_path)

        # group 1: input
        self.grid1 = QGridLayout()
        self.grid1.addWidget(self.lin1, 0, 1, 1, 5)
        self.grid1.addWidget(self.btn1, 0, 0, 1, 1)
        self.grid1.addWidget(self.lin2, 1, 1, 1, 5)
        self.grid1.addWidget(self.btn2, 1, 0, 1, 1)
        self.grid1.addWidget(self.lin3, 2, 1, 1, 5)
        self.grid1.addWidget(self.btn3, 2, 0, 1, 1)
        self.grid1.addWidget(self.lin4, 3, 1, 1, 5)
        self.grid1.addWidget(self.btn4, 3, 0, 1, 1)
        self.group1.setLayout(self.grid1)

        # group variable
        self.group2 = QGroupBox('Variable Selection')
        self.group2.setFont(self.title_font)

        self.btn5 =QPushButton('Get Variables')
        self.btn5.setFont(self.cont_font)
        self.btn5.clicked.connect(self.get_variable)
        self.cbx1 = QCheckBox('Blade_root')
        self.cbx1.setFont(self.cont_font)
        self.cbx1.setDisabled(True)
        self.cbx2 = QCheckBox('Hub_rotating')
        self.cbx2.setFont(self.cont_font)
        self.cbx2.setDisabled(True)
        self.cbx3 = QCheckBox('Hub_stationary')
        self.cbx3.setFont(self.cont_font)
        self.cbx3.setDisabled(True)
        self.cbx4 = QCheckBox('Yaw_bearing')
        self.cbx4.setFont(self.cont_font)
        self.cbx4.setDisabled(True)
        self.cbx5 = QCheckBox('Tower_sections')
        self.cbx5.setFont(self.cont_font)
        self.cbx5.setDisabled(True)
        self.cbx5.stateChanged.connect(self.tower_checked)
        self.cbx6 = QCheckBox('Blade_sections')
        self.cbx6.setFont(self.cont_font)
        self.cbx6.setDisabled(True)
        self.cbx7 = QCheckBox('Principal')
        self.cbx7.setFont(self.cont_font)
        self.cbx7.setDisabled(True)
        self.cbx7.stateChanged.connect(self.principle_checked)
        self.cbx8 = QCheckBox('Root')
        self.cbx8.setFont(self.cont_font)
        self.cbx8.setDisabled(True)
        self.cbx8.stateChanged.connect(self.root_checked)
        self.cbx9 = QCheckBox('Aerodynamic')
        self.cbx9.setFont(self.cont_font)
        self.cbx9.setDisabled(True)
        self.cbx9.stateChanged.connect(self.aero_checked)
        self.cbx10 = QCheckBox('User')
        self.cbx10.setFont(self.cont_font)
        self.cbx10.setDisabled(True)
        self.cbx10.stateChanged.connect(self.user_checked)
        self.lbl0 = QLabel()

        self.cbx11 = QCheckBox('Basic statistics')
        self.cbx11.setDisabled(True)
        self.cbx12 = QCheckBox('Combination')
        self.cbx12.setDisabled(True)
        self.cbx13 = QCheckBox('Clearance')
        self.cbx13.setDisabled(True)
        self.cbx14 = QCheckBox('LRD')
        self.cbx14.setDisabled(True)
        self.cbx15 = QCheckBox('LDD')
        self.cbx15.setDisabled(True)
        self.cbx16 = QCheckBox('Blade fatigue')
        self.cbx16.setDisabled(True)
        self.cbx17 = QCheckBox('Tower fatigue')
        self.cbx17.setDisabled(True)
        self.cbx18 = QCheckBox('Foundation fatigue')
        self.cbx18.setDisabled(True)
        self.cbx19 = QCheckBox('Blade ultimate')
        self.cbx19.setDisabled(True)
        self.cbx20 = QCheckBox('Hub ultimate')
        self.cbx20.setDisabled(True)
        self.cbx21 = QCheckBox('Tower ultimate')
        self.cbx21.setDisabled(True)
        self.cbx22 = QCheckBox('Nacelle Acc')
        self.cbx22.setDisabled(True)
        # self.cbx23 = QCheckBox('Main ultimate')
        # self.cbx23.setDisabled(True)

        self.subgroup1 = QGroupBox('blade_sections')
        self.subgroup1.setFont(self.title_font)
        self.list1 = QListWidget()
        self.rad_blade_all = QRadioButton('All')
        self.rad_blade_all.clicked.connect(self.blade_selection)
        self.rad_blade_no  = QRadioButton('None')
        self.rad_blade_no.clicked.connect(self.blade_selection)
        self.rad_blade_no.setChecked(True)
        self.btngroup1 = QButtonGroup()
        self.btngroup1.addButton(self.rad_blade_all)
        self.btngroup1.addButton(self.rad_blade_no)
        self.sublayout1 = QGridLayout()
        self.sublayout1.addWidget(self.list1, 0, 0, 10, 2)
        # self.sublayout1.addWidget(self.rad_blade_all, 11, 0, 1, 1)
        self.sublayout1.addWidget(self.rad_blade_all, 11, 0, 1, 1)
        self.sublayout1.addWidget(self.rad_blade_no, 11, 1, 1, 1)
        self.subgroup1.setLayout(self.sublayout1)

        self.subgroup2 = QGroupBox('tower_sections')
        self.subgroup2.setFont(self.title_font)
        self.list2 = QListWidget()
        self.rad_tower_all = QRadioButton('All')
        self.rad_tower_all.clicked.connect(self.tower_selection)
        self.rad_tower_no  = QRadioButton('None')
        self.rad_tower_no.clicked.connect(self.tower_selection)
        self.rad_tower_no.setChecked(True)
        self.btngroup2 = QButtonGroup()
        self.btngroup2.addButton(self.rad_tower_all)
        self.btngroup2.addButton(self.rad_tower_no)
        self.sublayout2 = QGridLayout()
        self.sublayout2.addWidget(self.list2, 0, 0, 10, 2)
        self.sublayout2.addWidget(self.rad_tower_all, 11, 0, 1, 1)
        self.sublayout2.addWidget(self.rad_tower_no,  11, 1, 1, 1)
        self.subgroup2.setLayout(self.sublayout2)

        # group-br/hr/hs/yb/ts
        self.sub00 = QGroupBox()
        self.subgrid00 = QVBoxLayout()
        self.subgrid00.addWidget(self.cbx1)
        self.subgrid00.addWidget(self.cbx2)
        self.subgrid00.addWidget(self.cbx3)
        self.subgrid00.addWidget(self.cbx4)
        self.subgrid00.addWidget(self.cbx5)
        self.sub00.setLayout(self.subgrid00)

        # group-brs/bps/bas/bus
        self.sub01 = QGroupBox()
        self.subgrid01 = QGridLayout()
        self.subgrid01.addWidget(self.cbx6, 0, 0, 1, 4)
        # self.subgrid01.addWidget(self.lbl0, 1, 0, 1, 1)
        self.subgrid01.addWidget(self.cbx7, 1, 1, 1, 3)
        # self.subgrid01.addWidget(self.lbl0, 2, 0, 1, 1)
        self.subgrid01.addWidget(self.cbx8, 2, 1, 1, 3)
        # self.subgrid01.addWidget(self.lbl0, 3, 0, 1, 1)
        self.subgrid01.addWidget(self.cbx9, 3, 1, 1, 3)
        # self.subgrid01.addWidget(self.lbl0, 4, 0, 1, 1)
        self.subgrid01.addWidget(self.cbx10, 4, 1, 1, 3)
        self.sub01.setLayout(self.subgrid01)

        # other
        self.sub02 = QGroupBox()
        self.subgrid02 = QGridLayout()
        self.subgrid02.addWidget(self.cbx11, 0, 0, 1, 4)
        self.subgrid02.addWidget(self.cbx13, 1, 0, 1, 4)
        self.subgrid02.addWidget(self.cbx12, 2, 0, 1, 4)
        self.subgrid02.addWidget(self.cbx15, 3, 0, 1, 4)
        self.subgrid02.addWidget(self.cbx14, 4, 0, 1, 4)
        self.sub02.setLayout(self.subgrid02)

        self.sub03 = QGroupBox()
        self.subgrid03 = QGridLayout()
        self.subgrid03.addWidget(self.cbx16, 0, 0, 1, 4)
        # self.subgrid03.addWidget(self.cbx22, 1, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx17, 2, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx18, 3, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx19, 4, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx20, 5, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx21, 7, 0, 1, 4)
        self.subgrid03.addWidget(self.cbx22, 8, 0, 1, 4)
        self.sub03.setLayout(self.subgrid03)

        # group2 grid
        self.grid2_1 = QVBoxLayout()
        self.grid2_1.addWidget(self.btn5)
        self.grid2_1.addWidget(self.sub02)
        self.grid2_1.addWidget(self.sub03)
        self.grid2_1.addStretch(1)
        # self.grid2_1.addWidget(self.sub01)
        # self.grid2_1.addWidget(self.sub02)
        self.grid2_2 = QVBoxLayout()
        self.grid2_2.addWidget(self.sub00)
        self.grid2_2.addWidget(self.subgroup2)
        self.grid2_1.addStretch(1)

        self.grid2_3 = QVBoxLayout()
        self.grid2_3.addWidget(self.sub01)
        self.grid2_3.addWidget(self.subgroup1)
        self.grid2_1.addStretch(1)

        self.grid2 = QHBoxLayout()
        self.grid2.addLayout(self.grid2_2)
        self.grid2.addLayout(self.grid2_3)
        self.grid2.addLayout(self.grid2_1)
        self.grid2.setStretch(0, 4)
        self.grid2.setStretch(1, 4)
        self.grid2.setStretch(2, 4)
        # self.grid2.addWidget(self.subgroup1)
        # self.grid2.addWidget(self.subgroup2)
        # self.grid2_2.addWidget(self.subgroup1, 0, 5, 13, 1)
        self.group2.setLayout(self.grid2)

        # group 3: get load case
        self.group3 = QGroupBox('DLC selection')
        self.group3.setFont(self.title_font)

        self.rad_ult_all = QRadioButton('All')
        self.rad_ult_all.clicked.connect(self.ultimate_selection)
        self.rad_ult_no  = QRadioButton('None')
        self.rad_ult_no.clicked.connect(self.ultimate_selection)
        self.rad_ult_no.setChecked(True)
        self.btngroup3 = QButtonGroup()
        self.btngroup3.addButton(self.rad_ult_all)
        self.btngroup3.addButton(self.rad_ult_no)
        self.grid3_1 = QHBoxLayout()
        self.grid3_1.addWidget(self.rad_ult_all)
        self.grid3_1.addWidget(self.rad_ult_no)

        self.rad_fat_all = QRadioButton('All')
        self.rad_fat_all.clicked.connect(self.fatigue_selection)
        self.rad_fat_no  = QRadioButton('None')
        self.rad_fat_no.clicked.connect(self.fatigue_selection)
        self.rad_fat_no.setChecked(True)
        self.btngroup4 = QButtonGroup()
        self.btngroup4.addButton(self.rad_fat_all)
        self.btngroup4.addButton(self.rad_fat_no)
        self.grid3_2 = QHBoxLayout()
        self.grid3_2.addWidget(self.rad_fat_all)
        self.grid3_2.addWidget(self.rad_fat_no)

        self.subgroup3= QGroupBox('Ultimate DLC')
        self.subgroup3.setFont(self.title_font)
        self.list3 = QListWidget()
        self.list3.setFont(self.cont_font)
        self.sublayout3 = QVBoxLayout()
        self.sublayout3.addWidget(self.list3)
        self.sublayout3.addLayout(self.grid3_1)
        self.subgroup3.setLayout(self.sublayout3)

        self.subgroup4 = QGroupBox('Fatigue DLC')
        self.subgroup4.setFont(self.title_font)
        self.list4 = QListWidget()
        self.list4.setFont(self.cont_font)
        self.sublayout4 = QVBoxLayout()
        self.sublayout4.addWidget(self.list4)
        self.sublayout4.addLayout(self.grid3_2)
        self.subgroup4.setLayout(self.sublayout4)

        self.btn6 =QPushButton('Get Load Case')
        self.btn6.setFont(self.cont_font)
        self.btn6.clicked.connect(self.get_dlc)

        # group3 grid
        self.grid3 = QVBoxLayout()
        self.grid3.addWidget(self.btn6)
        self.grid3_3  = QHBoxLayout()
        self.grid3_3.addWidget(self.subgroup3)
        self.grid3_3.addWidget(self.subgroup4)
        # self.grid3_1.addStretch(1)
        self.grid3.addLayout(self.grid3_3)
        self.group3.setLayout(self.grid3)

        # group ultimate parameters
        self.group4 = QGroupBox('Ultimate parameters')
        self.group4.setFont(self.title_font)
        self.lbl5 = QLabel('Mean subgroup')
        self.lbl5.setFont(self.cont_font)
        self.lin5 = QLineEdit()
        self.lin5.setFont(self.cont_font)
        self.lin5.setText('#')
        self.lbl6 = QLabel('Half subgroup')
        self.lbl6.setAlignment(QtCore.Qt.AlignRight)
        self.lbl6.setFont(self.cont_font)
        self.lin6 = QLineEdit()
        self.lin6.setFont(self.cont_font)
        self.lin6.setText('+')
        self.cbx_etm = QCheckBox('ETM for DLC23')
        self.cbx_etm.setFont(self.cont_font)
        # self.cbx_etm.setDisabled(True)
        self.cbx_etm.clicked.connect(self.sf_selection)

        self.group8 = QGroupBox('Safety factor')
        self.group8.setFont(self.title_font)
        # self.rad_sf = QRadioButton('None')
        # self.rad_sf.setFont(self.cont_font)
        # self.rad_sf.setChecked(True)
        # self.rad_sf.clicked.connect(self.sf_selection)
        self.rad_iec_1 = QRadioButton('IEC61400-1_Ed3/-3')
        self.rad_iec_1.setFont(self.cont_font)
        # self.rad_iec_1.setChecked(True)
        self.rad_iec_1.clicked.connect(self.sf_selection)
        self.rad_iec_2 = QRadioButton('IEC61400-1_Ed4/-3.1')
        self.rad_iec_2.setFont(self.cont_font)
        self.rad_iec_2.setChecked(True)
        self.rad_iec_2.clicked.connect(self.sf_selection)
        # self.rad_user  = QRadioButton('User defined')
        # self.rad_user.clicked.connect(self.sf_selection)
        # self.rad_user.setFont(self.cont_font)
        self.btngroup3 = QButtonGroup()
        # self.btngroup3.addButton(self.rad_sf)
        # self.btngroup3.addButton(self.rad_user)
        self.btngroup3.addButton(self.rad_iec_1)
        self.btngroup3.addButton(self.rad_iec_2)
        self.subv5 = QGridLayout()
        # self.subv5.addWidget(self.rad_sf, 0, 0, 1, 1)
        self.subv5.addWidget(self.rad_iec_1, 1, 0, 1, 2)
        self.subv5.addWidget(self.rad_iec_2, 2, 0, 1, 2)
        # self.subv5.addWidget(self.rad_user, 1, 0, 1, 1)
        self.subv5.addWidget(self.cbx_etm, 3, 0, 1, 2)
        self.group8.setLayout(self.subv5)

        self.sub1 = QGroupBox()
        self.subgrid1 = QGridLayout()
        self.subgrid1.addWidget(self.lbl5, 0, 0, 1, 1)
        self.subgrid1.addWidget(self.lin5, 0, 1, 1, 1)
        self.subgrid1.addWidget(self.lbl6, 1, 0, 1, 1)
        self.subgrid1.addWidget(self.lin6, 1, 1, 1, 1)
        self.sub1.setLayout(self.subgrid1)

        self.subv1 = QVBoxLayout()
        # self.subv1.addWidget(self.cbx11)
        # self.subv1.addWidget(self.cbx16)
        self.subv1.addWidget(self.group8)
        self.subv1.addWidget(self.sub1)
        # self.subgroup5.addLayout(self.subgrid2)
        self.subv1.addStretch(1)
        self.group4.setLayout(self.subv1)

        # group fatigue parameters
        self.group5 = QGroupBox('Fatigue parameters')
        self.group5.setFont(self.title_font)
        self.sub2 = QGroupBox()
        self.lbl7 = QLabel('Bins Number')
        self.lbl7.setAlignment(QtCore.Qt.AlignRight)
        self.lbl7.setFont(self.cont_font)
        self.lin7 = QLineEdit()
        self.lin7.setFont(self.cont_font)
        self.lin7.setText('64')
        self.lbl8 = QLabel('Design Life\n(year)')
        self.lbl8.setAlignment(QtCore.Qt.AlignRight)
        self.lbl8.setFont(self.cont_font)
        self.lin8 = QLineEdit()
        self.lin8.setFont(self.cont_font)
        self.lin8.setText('25')
        self.lbl9 = QLabel('Eq_cycles\n(1E+7)')
        self.lbl9.setAlignment(QtCore.Qt.AlignRight)
        self.lbl9.setFont(self.cont_font)
        self.lin9 = QLineEdit()
        self.lin9.setFont(self.cont_font)
        self.lin9.setText('1')

        self.subgrid2 = QGridLayout()
        self.subgrid2.addWidget(self.lbl7, 0, 0, 1, 1)
        self.subgrid2.addWidget(self.lin7, 0, 1, 1, 1)
        self.subgrid2.addWidget(self.lbl8, 1, 0, 1, 1)
        self.subgrid2.addWidget(self.lin8, 1, 1, 1, 1)
        self.subgrid2.addWidget(self.lbl9, 2, 0, 1, 1)
        self.subgrid2.addWidget(self.lin9, 2, 1, 1, 1)
        self.group5.setLayout(self.subgrid2)

        self.group6 = QGroupBox('Output Identifier')
        self.group6.setFont(self.title_font)
        self.sub3 = QGroupBox()
        self.lbl10 = QLabel('Ultimate')
        self.lbl10.setFont(self.cont_font)
        self.lin10 = QLineEdit()
        self.lin10.setFont(self.cont_font)
        self.lbl11 = QLabel('Fatigue')
        self.lbl11.setFont(self.cont_font)
        self.lin11 = QLineEdit()
        self.lin11.setFont(self.cont_font)

        self.subgrid3 = QGridLayout()
        self.subgrid3.addWidget(self.lbl10, 0, 0, 1, 1)
        self.subgrid3.addWidget(self.lin10, 0, 1, 1, 1)
        self.subgrid3.addWidget(self.lbl11, 1, 0, 1, 1)
        self.subgrid3.addWidget(self.lin11, 1, 1, 1, 1)
        self.sub3.setLayout(self.subgrid3)
        self.group6.setLayout(self.subgrid3)

        self.subv3 = QVBoxLayout()
        self.subv3.addWidget(self.sub3)

        # post type
        self.group7 = QGroupBox('Post type')
        self.group7.setFont(self.title_font)
        self.cb_ult = QCheckBox('Ultmate')
        self.cb_ult.setFont(self.cont_font)
        self.cb_ult.setDisabled(True)
        # self.cb_ult.stateChanged.connect(self.ultimate_choosed)
        self.cb_fat = QCheckBox('Fatigue')
        self.cb_fat.setFont(self.cont_font)
        self.cb_fat.setDisabled(True)
        # self.cb_fat.stateChanged.connect(self.fatigue_choosed)
        self.cb_pst = QCheckBox('Post')
        self.cb_pst.setFont(self.cont_font)
        self.cb_pst.setDisabled(True)
        self.cb_pst.clicked.connect(self.post_checked)
        self.cb_cmp = QCheckBox('Components')
        self.cb_cmp.clicked.connect(self.component_checked)
        self.cb_cmp.setFont(self.cont_font)
        self.cb_cmp.setDisabled(True)
        self.cb_main = QCheckBox('Main')
        # self.cb_main.setObjectName('main')
        self.cb_main.setFont(self.cont_font)
        self.cb_main.setDisabled(True)
        # self.cb_main.setChecked(True)
        self.cb_main.clicked.connect(self.main_checked)


        self.subv4 = QGridLayout()
        self.subv4.addWidget(self.cb_ult, 0, 0, 1, 1)
        self.subv4.addWidget(self.cb_main, 0, 1, 1, 1)
        self.subv4.addWidget(self.cb_fat, 1, 0, 1, 1)
        # self.subv4.addWidget(self.cb_fat_main, 1, 1, 1, 1)
        self.subv4.addWidget(self.cb_pst, 2, 0, 1, 1)
        self.subv4.addWidget(self.cb_cmp, 2, 1, 1, 1)
        self.group7.setLayout(self.subv4)

        self.btn7 = QPushButton('Run')
        self.btn7.setFont(self.cont_font)
        self.btn7.clicked.connect(self.write_joblist)

        # main layout
        # input/output-variable selection
        self.main_layout1 = QVBoxLayout()
        self.main_layout1.addWidget(self.group1)
        self.main_layout1.addWidget(self.group2)
        self.main_layout1.addStretch(0)

        # dlc selection
        self.main_layout2 = QVBoxLayout()
        self.main_layout2.addWidget(self.group3)
        # self.main_layout2.addStretch(0)

        # ultimate parameters
        self.main_layout3 = QVBoxLayout()
        self.main_layout3.addWidget(self.group4)
        self.main_layout3.addWidget(self.group5)
        self.main_layout3.addWidget(self.group6)
        # self.main_layout3.addWidget(self.group8)
        self.main_layout3.addWidget(self.group7)
        self.main_layout3.addWidget(self.btn7)
        self.main_layout3.addStretch(0)

        # input/output/variable/dlc
        self.main_layout4 = QGridLayout()
        self.main_layout4.addLayout(self.main_layout1, 0, 0, 1, 1)
        self.main_layout4.addLayout(self.main_layout2, 0, 1, 1, 1)
        self.main_layout4.addLayout(self.main_layout3, 0, 2, 1, 1)
        self.main_layout4.setColumnMinimumWidth(0, 430)
        self.main_layout4.setColumnMinimumWidth(1, 270)
        self.main_layout4.setColumnMinimumWidth(2, 150)

        self.main_layout = QVBoxLayout()
        self.main_layout.addLayout(self.main_layout4)
        self.main_layout.addStretch(0)

        self.mywidget.setLayout(self.main_layout)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def my_exit(self):
        self.close()

    def validate(narg):
        def decorate(func):
            def funcDecorated(*args):
                try:
                    os.path.isdir(args[narg])
                except Exception:
                    # wrong argument! do some logging here and re-raise
                    raise Exception("Invalid path: {}".format(args[0]))
                else:
                    return func(*args)
            return funcDecorated
        return decorate

    def clear_var_list(self):

        self.cbx1.setChecked(False)
        self.cbx1.setDisabled(True)
        self.cbx2.setChecked(False)
        self.cbx2.setDisabled(True)
        self.cbx3.setChecked(False)
        self.cbx3.setDisabled(True)
        self.cbx4.setChecked(False)
        self.cbx4.setDisabled(True)
        self.cbx5.setChecked(False)
        self.cbx5.setDisabled(True)
        self.cbx6.setChecked(False)
        self.cbx6.setDisabled(True)
        self.cbx7.setChecked(False)
        self.cbx7.setDisabled(True)
        self.cbx8.setChecked(False)
        self.cbx8.setDisabled(True)
        self.cbx9.setChecked(False)
        self.cbx9.setDisabled(True)
        self.cbx10.setChecked(False)
        self.cbx10.setDisabled(True)
        self.cbx11.setChecked(False)
        self.cbx11.setDisabled(True)
        self.cbx12.setChecked(False)
        self.cbx12.setDisabled(True)
        self.cbx13.setChecked(False)
        self.cbx13.setDisabled(True)
        self.cbx14.setChecked(False)
        self.cbx14.setDisabled(True)
        self.cbx15.setChecked(False)
        self.cbx15.setDisabled(True)
        self.cbx16.setChecked(False)
        self.cbx16.setDisabled(True)
        self.cbx17.setChecked(False)
        self.cbx17.setDisabled(True)
        self.cbx18.setChecked(False)
        self.cbx18.setDisabled(True)
        self.cbx19.setChecked(False)
        self.cbx19.setDisabled(True)
        self.cbx20.setChecked(False)
        self.cbx20.setDisabled(True)
        self.cbx21.setChecked(False)
        self.cbx21.setDisabled(True)
        self.cbx22.setChecked(False)
        self.cbx22.setDisabled(True)
        self.cb_ult.setChecked(False)
        self.cb_ult.setDisabled(True)
        self.cb_fat.setChecked(False)
        self.cb_fat.setDisabled(True)
        self.cb_main.setChecked(False)
        self.cb_main.setDisabled(True)
        self.cb_pst.setChecked(False)
        self.cb_pst.setDisabled(True)
        self.cb_cmp.setChecked(False)
        self.cb_cmp.setDisabled(True)
        self.list1.clear()
        self.list2.clear()
        self.rad_blade_no.setChecked(True)
        self.rad_tower_no.setChecked(True)

    def clear_dlc_list(self):

        self.list3.clear()
        self.list4.clear()
        self.rad_ult_no.setChecked(True)
        self.rad_fat_no.setChecked(True)

    def save_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if not config.has_section('Post'):
            config.add_section('Post')

        config['Post'] = {'Prj path':self.lin1.text(),
                          'LCT path':self.lin2.text(),
                          'Run path':self.lin3.text(),
                          'Post path':self.lin4.text(),
                          'Ultimate':self.lin10.text(),
                          'Fatigue':self.lin11.text()}

        with open("config.ini",'w') as f:
            config.write(f)

    def load_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if config.has_section('Post'):

            self.lin1.setText(config.get('Post','Prj path'))
            # self.lin1.home(False)
            self.lin2.setText(config.get('Post','LCT path'))
            # self.lin2.home(False)
            self.lin3.setText(config.get('Post','Run path'))
            # self.lin3.home(False)
            self.lin4.setText(config.get('Post','Post path'))
            # self.lin4.home(False)
            self.lin10.setText(config.get('Post','Ultimate'))
            # self.lin10.home(False)
            self.lin11.setText(config.get('Post','Fatigue'))
            # self.lin11.home(False)

    def clear_setting(self):
        self.lin1.setText('')
        self.lin2.setText('')
        self.lin3.setText('')
        self.lin4.setText('')

    def user_manual(self):
        os.startfile(os.getcwd()+r'\user manual\Post Assistant.docx')

    def load_project(self):

        self.project, filetype1 = QFileDialog.getOpenFileName(self,
                                                               "Choose project dialog",
                                                               r".",
                                                               "project(*.$PJ)")

        if self.project:
            self.lin1.setText(self.project.replace('/', '\\'))
            self.lin1.home(True)

            # self.loadcase = os.path.basename(self.project).split('.')[0]

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_run_path(self):

        self.run_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose run path",
                                                         r".")

        if self.run_path:
            self.lin3.setText(self.run_path.replace('/', '\\'))
            self.lin3.home(True)

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_post_path(self):

        self.post_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose post path",
                                                         r".")

        if self.post_path:
            self.lin4.setText(self.post_path.replace('/', '\\'))
            self.lin4.home(True)

    def load_batch_path(self):
        self.batch_path = QFileDialog.getExistingDirectory(self,
                                                           "Choose batch path",
                                                           r".")

        if self.batch_path:
            self.lin15.setText(self.batch_path.replace('/', '\\'))
            self.lin15.home(True)

    def load_lct(self):

        self.lct, filetype1 = QFileDialog.getOpenFileName(self,
                                                               "Choose load case table dialog",
                                                               r".",
                                                               "Excel(*.xlsm)")

        if self.lct:
            self.lin2.setText(self.lct.replace('/', '\\'))
            self.lin2.home(True)

    def get_variable(self):

        if self.lin1.text().upper().endswith('$PJ') or self.lin1.text().lower().endswith('prj'):
            # print(self.lin1.text())

            try:
                self.var_result = Get_Variables(self.lin1.text())
                # print(self.var_result)

                if self.var_result.br_flag:
                    self.cbx1.setDisabled(False)
                    # self.cbx1.setChecked(True)
                if self.var_result.hr_flag:
                    self.cbx2.setDisabled(False)
                    # self.cbx2.setChecked(True)
                if self.var_result.hs_flag:
                    self.cbx3.setDisabled(False)
                    # self.cbx3.setChecked(True)
                if self.var_result.yb_flag:
                    self.cbx4.setDisabled(False)
                    # self.cbx4.setChecked(True)
                if self.var_result.tr_flag:
                    self.cbx5.setDisabled(False)
                    # self.cbx5.setChecked(True)
                if self.var_result.bp_flag:
                    self.cbx7.setDisabled(False)
                    # self.cbx7.setChecked(True)
                    self.cbx6.setChecked(True)
                if self.var_result.bs_flag:
                    self.cbx8.setDisabled(False)
                    # self.cbx8.setChecked(True)
                    self.cbx6.setChecked(True)
                if self.var_result.ba_flag:
                    self.cbx9.setDisabled(False)
                    self.cbx9.setChecked(False)
                    self.cbx6.setChecked(True)
                if self.var_result.bu_flag:
                    self.cbx10.setDisabled(False)
                    self.cbx10.setChecked(False)
                    self.cbx6.setChecked(True)

                # basic statistics
                if any((self.var_result.yb_flag, self.var_result.br_flag, self.var_result.mb_flag)):
                    self.cbx11.setDisabled(False)
                    # self.cbx11.setTristate(True)

                # combination
                if self.var_result.br_flag or self.var_result.bs_flag:
                    self.cbx12.setDisabled(False)
                    # self.cbx12.setTristate(True)

                # clearance
                if self.var_result.tc_flag:
                    self.cbx13.setDisabled(False)

                # ldd
                if self.var_result.hs_flag or self.var_result.yb_flag or self.var_result.hs_flag or self.var_result.br_flag:
                    self.cbx15.setDisabled(False)
                    # self.cbx14.setTristate(True)

                # lrd
                # ws_flag for wind speed and mean pitch angle
                if (self.var_result.br_flag and self.var_result.ws_flag) or (self.var_result.hs_flag and self.var_result.rs_flag):
                    self.cbx14.setDisabled(False)
                    # self.cbx15.setTristate(True)

                # fatigue-brs/bus/br_mxy/brs_mxy_angle
                if any((self.var_result.bs_flag,self.var_result.bu_flag,self.var_result.br_mxy_flag,
                        self.var_result.brs1_mxy_flag,self.var_result.brs2_mxy_flag,self.var_result.brs3_mxy_flag)):
                    self.cbx16.setDisabled(False)
                    # self.cbx16.setTristate(True)

                # fatigue-tr/ts_0_40*40
                if self.var_result.tr_flag:
                    self.cbx17.setDisabled(False)
                    self.cbx18.setDisabled(False)

                # ultimate-blade
                if self.var_result.bs_flag or self.var_result.bu_flag:
                    self.cbx19.setDisabled(False)
                    # self.cbx19.setTristate(True)

                # ultimate-hub
                if self.var_result.hs_flag or self.var_result.hr_flag:
                    self.cbx20.setDisabled(False)

                # ultiamte-tower
                if self.var_result.tr_flag:
                    self.cbx21.setDisabled(False)

                # ultimate-nacelle acceleration
                if self.var_result.nac_flag:
                    self.cbx22.setDisabled(False)

                self.rad_blade_no.setChecked(True)
                self.rad_tower_no.setChecked(True)
                self.list1.clear()
                self.list2.clear()

                warning = ''
                if not self.var_result.br_flag:
                    warning += 'Blade root result not exist!\n'
                if not self.var_result.hr_flag:
                    warning += 'Hub rotating result not exist!\n'
                if not self.var_result.hs_flag:
                    warning += 'Hub stationary result not exist!\n'
                if not self.var_result.yb_flag:
                    warning += 'Yaw bearing result not exist!\n'
                if not self.var_result.tr_flag:
                    warning += 'Tower/foundation result not exist!\n'
                if not self.var_result.bp_flag:
                    warning += 'Blade principle result not exist!\n'
                if not self.var_result.bs_flag:
                    warning += 'Blade root result not exist!\n'
                if not self.var_result.bu_flag:
                    warning += 'Blade user axes result not exist!\n'
                if not self.var_result.tc_flag:
                    warning += 'Tower clearance result not exist!\n'
                if not self.var_result.ba_flag:
                    warning += 'Blade aerodynamic axes result not exist!\n'
                if not self.var_result.br_mxy_flag:
                    warning += 'Blade root combination result not exist!\n'
                if not self.var_result.brs1_mxy_flag:
                    warning += 'Blade 1 root axes combination result not exist!\n'
                if not self.var_result.brs2_mxy_flag:
                    warning += 'Blade 2 root axes combination result not exist!\n'
                if not self.var_result.brs3_mxy_flag:
                    warning += 'Blade 3 root axes combination result not exist!\n'
                if not self.var_result.nac_flag:
                    warning += 'Nacelle acceleration result not exist!\n'

                if warning:
                    QMessageBox.about(self, 'Warning', warning)

                # all variable
                cb_list     = [self.cbx1, self.cbx2, self.cbx3, self.cbx4, self.cbx5, self.cbx6, self.cbx7, self.cbx8, self.cbx9, self.cbx10]
                var_enabled = [True if cb.isEnabled() else False for cb in cb_list]

                if any(var_enabled):
                    self.cb_ult.setDisabled(False)
                    self.cb_fat.setDisabled(False)

                # main variable
                cb_list     = [self.cbx1, self.cbx2, self.cbx3, self.cbx4, self.cbx5, self.cbx8, self.cbx10]
                var_enabled = [True if cb.isEnabled() else False for cb in cb_list]

                if all(var_enabled):
                    self.cb_main.setDisabled(False)

                # post variable
                cb_list     = [self.cbx11, self.cbx12, self.cbx13, self.cbx14, self.cbx15]
                var_enabled = [True if cb.isEnabled() else False for cb in cb_list]

                if any(var_enabled):
                    self.cb_pst.setDisabled(False)

                # component variable
                cb_list     = [self.cbx16, self.cbx17, self.cbx18, self.cbx19, self.cbx20, self.cbx21, self.cbx22]
                var_enabled = [True if cb.isEnabled() else False for cb in cb_list]

                if any(var_enabled):
                    self.cb_cmp.setDisabled(False)

            except IOError:

                QMessageBox.about(self, 'Warning', 'Not invalid project path!')
                # print('Not invalid project path!')
        elif not self.lin1.text():
            QMessageBox.about(self, 'Warning', 'Project file has not been defined!')
        elif not(self.lin1.text().upper().endswith('$PJ') or self.lin1.text().lower().endswith('prj')):
            QMessageBox.about(self, 'Warning', 'Please choose a right $PJ or prj file!')

    def get_dlc(self):

        self.list3.clear()
        self.list4.clear()
        self.rad_ult_no.setChecked(True)
        self.rad_fat_no.setChecked(True)

        if os.path.isdir(self.lin3.text()):

            try:
                dlc_list = Get_DLC(self.lin3.text())
                fatigue_list = dlc_list.fatigue_list
                print('Fatigue load case: ',fatigue_list)
                for i in fatigue_list:
                    box  = QCheckBox(i)
                    # box.setChecked(True)
                    item = QListWidgetItem()
                    # item.setFont(self.cont_font)
                    self.list4.addItem(item)
                    self.list4.setItemWidget(item, box)

                ultimate_list = dlc_list.ultimate_list
                print('Ultimate load case:',ultimate_list)
                for i in ultimate_list:
                    box  = QCheckBox(i)
                    box.setFont(self.cont_font)
                    # box.setChecked(True)
                    item = QListWidgetItem()
                    # item.setFont(self.cont_font)
                    self.list3.addItem(item)
                    self.list3.setItemWidget(item, box)

            except Exception as e:
                QMessageBox.about(self, 'Warning', 'Error occurs when reading run path\n%s!' %e)

        else:
            QMessageBox.about(self, 'Warning', 'Please choose a right run path!')

    def tower_checked(self, state):

        if self.var_result and state == QtCore.Qt.Checked:
            twr_sec_list = []
            self.list2.clear()
            if self.var_result.tr_flag:
                if self.var_result.twr_model == '2':
                    twr_sec_list = [self.var_result.twr_height[int(i)-1] for i in self.var_result.twr_output]

                    for i in twr_sec_list:
                        box  = QCheckBox(i)
                        box.setFont(self.cont_font)
                        item = QListWidgetItem()
                        self.list2.addItem(item)
                        self.list2.setItemWidget(item, box)
                elif self.var_result.twr_model == '3':

                    for i in self.var_result.twr_output:

                        twr_sec_list.append('Mbr %s End 1' %i)
                        twr_sec_list.append('Mbr %s End 2' %i)

                    for i in twr_sec_list:
                        box = QCheckBox(i)
                        box.setFont(self.cont_font)
                        item = QListWidgetItem()
                        self.list2.addItem(item)
                        self.list2.setItemWidget(item, box)
        else:
            self.list2.clear()
            self.rad_tower_no.setChecked(True)

    def principle_checked(self, state):

        if self.var_result and state == QtCore.Qt.Checked:

            self.list1.clear()
            self.cbx6.setChecked(True)
            self.cbx8.setChecked(False)
            self.cbx9.setChecked(False)
            self.cbx10.setChecked(False)

            if self.var_result.bp_flag:

                bld_sec_list = []
                if self.var_result.bpa_out_type == '1':
                    bld_sec_list = self.var_result.bld_radius[0]
                elif self.var_result.bpa_out_type == 'A':
                    bld_sec_list = [self.var_result.bld_radius[int(i)-1] for i in self.var_result.bpa_output]
                # print(bld_sec_list)
                for i in bld_sec_list:
                    box  = QCheckBox(i)
                    box.setFont(self.cont_font)
                    item = QListWidgetItem()
                    self.list1.addItem(item)
                    self.list1.setItemWidget(item, box)
            else:
                QMessageBox.about(self, 'Warning', 'Principle axes results are not outputted!')
        else:
            self.list1.clear()
            self.rad_blade_no.setChecked(True)

    def root_checked(self, state):

        if self.var_result:

            if self.cbx9.isChecked() or self.cbx10.isChecked():
                pass

            elif state == QtCore.Qt.Checked:

                self.list1.clear()
                self.cbx6.setChecked(True)
                self.cbx7.setChecked(False)

                if self.var_result.bs_flag:

                    bld_sec_list = []
                    if self.var_result.bra_out_type == '1':
                        bld_sec_list = self.var_result.bld_radius[0]
                    elif self.var_result.bra_out_type == 'A':
                        bld_sec_list = self.var_result.bld_radius

                    for i in bld_sec_list:
                        box = QCheckBox(i)
                        box.setFont(self.cont_font)
                        item = QListWidgetItem()
                        self.list1.addItem(item)
                        self.list1.setItemWidget(item, box)
                else:
                    QMessageBox.about(self, 'Warning', 'Root axes results are not outputted!')
            else:
                self.list1.clear()
                self.rad_blade_no.setChecked(True)

    def aero_checked(self, state):

        if self.var_result:

            if self.cbx8.isChecked() or self.cbx10.isChecked():
                pass

            elif state == QtCore.Qt.Checked:

                self.list1.clear()
                self.cbx6.setChecked(True)
                self.cbx7.setChecked(False)

                if self.var_result.ba_flag:

                    bld_sec_list = []
                    if self.var_result.baa_out_type == '1':
                        bld_sec_list = self.var_result.bld_radius[0]
                    elif self.var_result.baa_out_type == 'A':
                        bld_sec_list = self.var_result.bld_radius

                    for i in bld_sec_list:
                        box = QCheckBox(i)
                        box.setFont(self.cont_font)
                        item = QListWidgetItem()
                        self.list1.addItem(item)
                        self.list1.setItemWidget(item, box)
                else:
                    QMessageBox.about(self, 'Warning', 'Aerodynamic axes results are not outputted!')
            else:
                self.list1.clear()
                self.rad_blade_no.setChecked(True)

    def user_checked(self, state):

        if self.var_result:

            if self.cbx8.isChecked() or self.cbx9.isChecked():
                pass

            elif state == QtCore.Qt.Checked:

                self.list1.clear()
                self.cbx6.setChecked(True)
                self.cbx7.setChecked(False)

                if self.var_result.bu_flag:

                    bld_sec_list = []
                    if self.var_result.bua_out_type == '1':
                        bld_sec_list = self.var_result.bld_radius[0]
                    elif self.var_result.bua_out_type == 'A':
                        bld_sec_list = self.var_result.bld_radius

                    for i in bld_sec_list:
                        box = QCheckBox(i)
                        box.setFont(self.cont_font)
                        item = QListWidgetItem()
                        self.list1.addItem(item)
                        self.list1.setItemWidget(item, box)
                else:
                    QMessageBox.about(self, 'Warning', 'User axes results are not outputted!')
            else:
                self.list1.clear()
                self.rad_blade_no.setChecked(True)

    def sf_selection(self, state):

        # if  self.rad_sf.isChecked():
        #     self.sf_flag = '0'
        #     self.cbx_etm.setChecked(False)
        #     self.cbx_etm.setDisabled(True)

        # IEC61400-1 ED3
        if self.rad_iec_1.isChecked():
            self.sf_flag = '1'
            self.cbx_etm.setChecked(False)
            self.cbx_etm.setDisabled(True)
        # IEC61400-1 ED4
        elif self.rad_iec_2.isChecked():
            self.sf_flag = '2'
            self.cbx_etm.setDisabled(False)

        if self.cbx_etm.isChecked():
            self.etm_opt = False
        else:
            self.etm_opt = False

    def tower_selection(self):
        count = self.list2.count()
        # print(count)
        if self.rad_tower_all.isChecked():
            for i in range(count):
                cb = self.list2.itemWidget(self.list2.item(i))
                # print(cb)
                cb.setChecked(True)
        elif self.rad_tower_no.isChecked():
            for i in range(count):
                cb = self.list2.itemWidget(self.list2.item(i))
                # print(cb)
                cb.setChecked(False)

    def blade_selection(self):
        count = self.list1.count()
        if self.rad_blade_all.isChecked():
            for i in range(count):
                cb = self.list1.itemWidget(self.list1.item(i))
                cb.setChecked(True)
        if self.rad_blade_no.isChecked():
            for i in range(count):
                cb = self.list1.itemWidget(self.list1.item(i))
                cb.setChecked(False)

    def ultimate_selection(self):
        count = self.list3.count()
        if self.rad_ult_all.isChecked():
            for i in range(count):
                cb = self.list3.itemWidget(self.list3.item(i))
                cb.setChecked(True)
        if self.rad_ult_no.isChecked():
            for i in range(count):
                cb = self.list3.itemWidget(self.list3.item(i))
                cb.setChecked(False)

    def fatigue_selection(self):
        count = self.list4.count()
        if self.rad_fat_all.isChecked():
            for i in range(count):
                cb = self.list4.itemWidget(self.list4.item(i))
                cb.setChecked(True)
        if self.rad_fat_no.isChecked():
            for i in range(count):
                cb = self.list4.itemWidget(self.list4.item(i))
                cb.setChecked(False)

    def get_choosed(self):
        '''
        get choosed optons
        :return: variable, blade sections, tower sections, ultimate load case, fatigue load case
        '''

        # main variable
        cb_list = [self.cbx1, self.cbx2, self.cbx3, self.cbx4, self.cbx5, self.cbx7, self.cbx8, self.cbx9, self.cbx10,
                   self.cbx11, self.cbx12, self.cbx13, self.cbx14, self.cbx15, self.cbx16, self.cbx17, self.cbx18,
                   self.cbx19, self.cbx20, self.cbx21, self.cbx22]
        var_choosed = [cb.text() for cb in cb_list if cb.isChecked()]
        # print(var_choosed)

        count   = self.list1.count()
        cb_list = [self.list1.itemWidget(self.list1.item(i)) for i in range(count)]
        bld_choosed = [cb.text() for cb in cb_list if cb.isChecked()]
        # print(bld_choosed)

        count   = self.list2.count()
        cb_list = [self.list2.itemWidget(self.list2.item(i)) for i in range(count)]
        twr_choosed = [cb.text() for cb in cb_list if cb.isChecked()]
        # print(twr_choosed)

        count   = self.list3.count()
        cb_list = [self.list3.itemWidget(self.list3.item(i)) for i in range(count)]
        ult_choosed = [cb.text() for cb in cb_list if cb.isChecked()]
        # print(ult_choosed)

        count   = self.list4.count()
        cb_list = [self.list4.itemWidget(self.list4.item(i)) for i in range(count)]
        fat_choosed = [cb.text() for cb in cb_list if cb.isChecked()]
        # print(fat_choosed)

        return var_choosed, bld_choosed, twr_choosed, ult_choosed, fat_choosed

    def main_checked(self):

        # main variable
        cb_list     = [self.cbx1, self.cbx2, self.cbx3, self.cbx4, self.cbx5, self.cbx8, self.cbx10]
        var_enabled = [True if cb.isEnabled() else False for cb in cb_list ]
        # print(var_enabled)
        # print(self.list1.count(), self.list2.count(), self.cb_main.checkState())

        if all(var_enabled):

            if self.cb_main.isChecked():

                self.cbx1.setChecked(True)
                self.cbx2.setChecked(True)
                self.cbx3.setChecked(True)
                self.cbx4.setChecked(True)
                self.cbx5.setChecked(True)
                self.cbx8.setChecked(True)
                self.cbx10.setChecked(True)

                count   = self.list1.count()
                if not count:
                    self.root_checked(QtCore.Qt.Checked)
                cb_list = [self.list1.itemWidget(self.list1.item(i)) for i in range(count)]
                cb_list[0].setChecked(True)
                # cb_list[-1].setChecked(True)

                count   = self.list2.count()
                if not count:
                    self.tower_checked(QtCore.Qt.Checked)
                cb_list = [self.list2.itemWidget(self.list2.item(i)) for i in range(count)]
                cb_list[0].setChecked(True)
                cb_list[-1].setChecked(True)

                self.lin10.setDisabled(True)
                self.lin11.setDisabled(True)

                self.cb_ult.setChecked(True)
                self.cb_fat.setChecked(True)

            else:

                self.cbx1.setChecked(False)
                self.cbx2.setChecked(False)
                self.cbx3.setChecked(False)
                self.cbx4.setChecked(False)
                self.cbx5.setChecked(False)
                self.cbx8.setChecked(False)
                self.cbx10.setChecked(False)

                # self.cb_main.setChecked(False)
                self.lin10.setDisabled(False)
                self.lin11.setDisabled(False)

                self.cb_ult.setChecked(False)
                self.cb_fat.setChecked(False)

        elif not all(var_enabled):
            self.cb_main.setChecked(False)
            # self.cb_main.setCheckState(QtCore.Qt.Unchecked)
            QMessageBox.about(self, 'window', 'Please make sure all variables are enabled!')

        elif not self.list1.count():
            self.cb_main.setChecked(False)
            # self.cb_main.setCheckState(QtCore.Qt.Unchecked)
            QMessageBox.about(self, 'window', 'Please make sure blade sections are enabled!')

        elif not self.list2.count():
            self.cb_main.setChecked(False)
            # self.cb_main.setCheckState(QtCore.Qt.Unchecked)
            QMessageBox.about(self, 'window', 'Please make sure tower sections are enabled!')
        # print(self.cb_main.checkState())
        # print(bool(state))

    def post_checked(self):

        # cb_list     = [self.cbx11, self.cbx12, self.cbx13, self.cbx14, self.cbx15]
        # cmp_choosed = [True if cb.isEnabled() else False for cb in cb_list]
        # print(cmp_choosed)

        # if all(cmp_choosed):

        if self.cb_pst.isChecked():

            self.cbx11.setChecked(True)
            self.cbx12.setChecked(True)
            self.cbx13.setChecked(True)
            self.cbx14.setChecked(True)
            self.cbx15.setChecked(True)

        else:

            self.cbx11.setChecked(False)
            self.cbx12.setChecked(False)
            self.cbx13.setChecked(False)
            self.cbx14.setChecked(False)
            self.cbx15.setChecked(False)

        # else:
        #     self.cb_cmp.setChecked(False)
        #     QMessageBox.about(self, 'window', 'Please make sure all components are enabled!')
        # print(self.cb_cmp.checkState())

    def component_checked(self):

        cb_list     = [self.cbx16, self.cbx17, self.cbx18, self.cbx19, self.cbx20, self.cbx21, self.cbx22]
        cmp_choosed = [True if cb.isEnabled() else False for cb in cb_list]
        # print(cmp_choosed)

        if all(cmp_choosed):

            if self.cb_cmp.isChecked():

                self.cbx16.setChecked(True)
                self.cbx17.setChecked(True)
                self.cbx18.setChecked(True)
                self.cbx19.setChecked(True)
                self.cbx20.setChecked(True)
                self.cbx21.setChecked(True)
                self.cbx22.setChecked(True)

            else:

                self.cbx16.setChecked(False)
                self.cbx17.setChecked(False)
                self.cbx18.setChecked(False)
                self.cbx19.setChecked(False)
                self.cbx20.setChecked(False)
                self.cbx21.setChecked(False)
                self.cbx22.setChecked(False)

        else:
            self.cb_cmp.setChecked(False)
            QMessageBox.about(self, 'window', 'Please make sure all components are enabled!')
        # print(self.cb_cmp.checkState())

    def handle(self):

        self.check_input = False
        self.post_list   = []       # post path for generating joblist

        #
        path_run = self.lin3.text()
        path_lct = self.lin2.text()
        
        result   = self.get_choosed()
        ult_list = result[3]
        fat_list = result[4]

        dlc12_flag = True if ('DLC12' in ' '.join(ult_list)) else False
        dlc13_list = [lc for lc in ult_list if 'DLC13' in lc] if ult_list else []

        blade_flag = ['Principal', 'Root', 'Aerodynamic', 'User']

        # check input
        if self.cb_ult.isChecked() or self.cb_fat.isChecked() or self.cb_pst.isChecked() or self.cb_cmp.isChecked():
            # print(any([self.cb_ult.isChecked(),self.cb_fat.isChecked(),self.cb_cmp.isChecked()]))

            # channel
            if (self.cb_fat.isChecked() or self.cb_ult.isChecked() or self.cb_main.isChecked()
                or self.cb_pst.isChecked() or self.cb_cmp.isChecked()) and not result[0]:
                QMessageBox.about(self, 'Warning', 'Please choose channel to output for fatigue or ultimate!')

            # blade section
            elif (not result[1]) and any(x in result[0] for x in blade_flag) and (self.cb_fat.isChecked() or self.cb_ult.isChecked()):
                QMessageBox.about(self, 'Warning', 'Please choose blade section to output!')

            # tower section
            elif (not result[2]) and 'Tower_sections' in result[0] and (self.cb_fat.isChecked() or self.cb_ult.isChecked()):
                QMessageBox.about(self, 'Warning', 'Please choose tower section to be output!')

            # output path
            elif not self.lin4.text():
                QMessageBox.about(self, 'Warning', 'Please choose a post path!')

            # ultimate load case for ultimate
            elif self.cb_ult.isChecked() and (not ult_list):
                QMessageBox.about(self, 'Warning', 'Please choose ultimate load case first!')

            # ultimate load case for component
            elif self.cb_cmp.isChecked() and (self.cbx19.isChecked() or self.cbx20.isChecked() or
                                                  self.cbx21.isChecked() or self.cbx22.isChecked()) and (not ult_list):
                QMessageBox.about(self, 'Warning', 'Please choose ultimate load case first!')

            # ultimate load case for post
            elif self.cb_pst.isChecked() and (self.cbx11.isChecked() or self.cbx13.isChecked()) and (not ult_list):
                QMessageBox.about(self, 'Warning', 'Please choose ultimate load case first!')

            # dlc12 for post and component ultimate option but not for ultimate function
            elif (self.cbx11.isChecked() or self.cbx13.isChecked() or self.cbx19.isChecked() or self.cbx20.isChecked()
                  or self.cbx21.isChecked() or self.cbx22.isChecked()) and dlc12_flag:
                QMessageBox.about(self, 'Warning', 'DLC12 is choosed in ultimate DLC!')

            # dlc13
            elif (self.cb_ult.isChecked() or self.cbx13.isChecked() or self.cbx19.isChecked() or self.cbx20.isChecked()
                  or self.cbx21.isChecked() or self.cbx22.isChecked()) and (len(dlc13_list)>1):
                QMessageBox.about(self, 'Warning', 'More than one dlc13 choosed in ultimate DLC!')

            # IEC61400-1 and ult_list need to be the same
            elif (self.cb_ult.isChecked() or self.cbx13.isChecked() or self.cbx19.isChecked() or self.cbx20.isChecked()
                  or self.cbx21.isChecked() or self.cbx22.isChecked()) and ('DLC25' in ult_list) and (self.sf_flag!='2'):
                QMessageBox.about(self, 'Warning', 'DLC25 not in IEC61400-1 ed3\n'
                                                   'Please choose a right safety factor!')

            # fatigue load case for fatigue function
            elif self.cb_fat.isChecked() and (not fat_list):
                QMessageBox.about(self, 'Warning', 'Please choose fatigue load case first!')

            # fatigue load case for post function
            elif (self.cbx11.isChecked() or self.cbx12.isChecked() or self.cbx14.isChecked() or self.cbx15.isChecked())\
                    and (not fat_list):
                QMessageBox.about(self, 'Warning', 'Please choose fatigue load case first!')

            # fatigue load case for component function
            elif (self.cbx16.isChecked() or self.cbx17.isChecked() or self.cbx18.isChecked()) and (not fat_list):
                QMessageBox.about(self, 'Warning', 'Please choose fatigue load case first!')

            elif (self.cb_fat.isChecked() or self.cb_cmp.isChecked() or self.cb_pst.isChecked()) and (not path_lct):
                QMessageBox.about(self, 'Warning', 'Please choose load case table!')

            elif (self.cb_fat.isChecked() or self.cb_cmp.isChecked() or self.cb_pst.isChecked())\
                    and (not os.path.isfile(path_lct)):
                QMessageBox.about(self, 'Warning', 'Please make sure load case table is right!')

            elif self.cbx21.isChecked() and ('DLC12' not in fat_list):
                QMessageBox.about(self, 'Warning', 'Please make choose DLC12 in fatigue list!!')
            else:
                self.check_input = True
            # print(self.check_input)

        else:
            QMessageBox.about(self, 'Warning', 'Please choose post type first!')

        try:
            if self.check_input:
                # ultimate
                if self.cb_ult.isChecked():

                    # sf_factor = False if self.rad_sf.isChecked() else True
                    mean_ind  = self.lin5.text().strip()
                    half_ind  = self.lin6.text().strip()

                    # ultimate channel dict
                    ult_chan = dict()
                    for chan in result[0]:

                        if chan == 'Blade_root':
                            ult_chan['br'] = self.var_result.br_var_list
                        if chan == 'Hub_rotating':
                            ult_chan['hr'] = self.var_result.hr_var_list
                        if chan == 'Hub_stationary':
                            ult_chan['hs'] = self.var_result.hs_var_list
                        if chan == 'Yaw_bearing':
                            ult_chan['yb'] = self.var_result.yb_var_list
                        if chan == 'Tower_sections':
                            ult_chan['tr']  = [self.var_result.tr_var_list,  result[2]]
                        if chan == 'Principal':
                            ult_chan['bps'] = [self.var_result.bps_var_list, result[1]]
                        if chan == 'Root':
                            ult_chan['brs'] = [self.var_result.brs_var_list, result[1]]
                        if chan == 'Aerodynamic':
                            ult_chan['bas'] = [self.var_result.bas_var_list, result[1]]
                        if chan == 'User':
                            ult_chan['bus'] = [self.var_result.bus_var_list, result[1]]

                    path = os.path.join(self.lin4.text(), '07_Ultimate')

                    print('began to handle main ultimate')
                    # include safety factor
                    if self.lin10.text():
                        path_ult = os.path.join(path, self.lin10.text()+'_Inclsf')
                    else:
                        path_ult = os.path.join(path, '08_Main_Inclsf')

                    # print(path_ult)
                    # print(ult_chan)
                    ult_res1 = Ultimate(run_path=path_run, ult_path=path_ult, dlc_list=ult_list, cha_dict=ult_chan,
                                        iec_stdard=self.sf_flag, etm_option=self.etm_opt, mean_index=mean_ind,
                                        half_index=half_ind,include_sf=True)
                    self.post_list.append(path_ult)

                    # exclude safety factor
                    if self.lin10.text() and (not self.cb_main.isChecked()):
                        path_ult = os.path.join(path, self.lin10.text()+'_Exclsf')
                    else:
                        path_ult = os.path.join(path, '09_Main_Exclsf')
                    # print(path_ult)
                    ult_res2 = Ultimate(run_path=path_run, ult_path=path_ult, dlc_list=ult_list, cha_dict=ult_chan,
                                        iec_stdard=self.sf_flag, etm_option=self.etm_opt, mean_index=mean_ind,
                                        half_index=half_ind,include_sf=False)
                    self.post_list.append(path_ult)

                # fatigue
                if self.cb_fat.isChecked():

                    num_bins = self.lin7.text()
                    des_life = self.lin8.text()
                    eq_cycle = self.lin9.text()+'E+07'

                    # fatigue channel dict
                    fat_chan = dict()
                    for chan in result[0]:
                        if chan == 'Blade_root':
                            fat_chan['br'] = [self.var_result.br_fat_list]
                        if chan == 'Hub_rotating':
                            fat_chan['hr'] = [self.var_result.hr_fat_list]
                        if chan == 'Hub_stationary':
                            fat_chan['hs'] = [self.var_result.hs_fat_list]
                        if chan == 'Yaw_bearing':
                            fat_chan['yb'] = [self.var_result.yb_fat_list]
                        if chan == 'Tower_sections':
                            fat_chan['tr']  = [self.var_result.tr_fat_list,  result[2]]
                        if chan == 'Principal':
                            fat_chan['bps'] = [self.var_result.bps_fat_list, result[1]]
                        if chan == 'Root':
                            fat_chan['brs'] = [self.var_result.brs_fat_list, result[1]]
                        if chan == 'Aerodynamic':
                            fat_chan['bas'] = [self.var_result.bas_fat_list, result[1]]
                        if chan == 'User':
                            fat_chan['bus'] = [self.var_result.bus_fat_list, result[1]]

                    path = os.path.join(self.lin4.text(), '08_Fatigue')

                    print('began to handle main fatigue')
                    if self.lin10.text() and (not self.cb_main.isChecked()):
                        path_fat = os.path.join(path, self.lin10.text())
                    else:
                        path_fat = os.path.join(path, '05_Main')

                    memo = os.path.join(self.lin4.text(), '08_Fatigue'+os.sep+'00_Life_%s_%s' %(des_life,eq_cycle))
                    if not os.path.exists(memo):
                        os.makedirs(memo)
                    # print(fat_chan)
                    fat_res = Fatigue(run_path=path_run,
                                      fat_path=path_fat,
                                      lct_path=path_lct,
                                      dlc_list=fat_list,
                                      num_bins=num_bins,
                                      des_life=des_life,
                                      eq_cycle=eq_cycle,
                                      fat_chan=fat_chan)
                    self.post_list.append(path_fat)
                    # run_result[0].append(fat_res.finish_fat)

                # post
                if self.cb_pst.isChecked():
                    #
                    mean_ind = self.lin5.text().strip()
                    half_ind = self.lin6.text().strip()

                    # ldd/lrd
                    num_bins = self.lin7.text()
                    des_life = self.lin8.text()
                    eq_cycle = self.lin9.text()+'E+07'

                    # cb_list = [self.cbx11, self.cbx12, self.cbx13, self.cbx14, self.cbx15]
                    # pst_choosed = [True for cb in cb_list if cb.isChecked()]

                    print('begin to handle post')
                    # 01_Extrapolation
                    # 02_Bstats
                    if self.cbx11.isChecked():
                        print('begin to handle basic statistics')
                        path_bs = os.path.join(self.lin4.text(), '02_Bstats')
                        if self.var_result.mb_flag:
                            chan_dict = {'mb':[['Hub wind speed magnitude','07'],
                                               ['Rotor speed','05'],
                                               ['Stationary hub Mx', '23']]}
                            post_path = os.path.join(path_bs, 'mb')
                            Bstats(path_run, fat_list, chan_dict, post_path, only_maxmin=False)

                        if self.var_result.hs_flag:
                            chan_dict = {'hs':[['Stationary hub Mx', '23']]}
                            post_path = os.path.join(path_bs, 'hs')
                            Bstats(path_run, fat_list, chan_dict, post_path, only_maxmin=False)

                        if self.var_result.br_flag:
                            chan_dict = {'br':[['Blade root 1 Mx', '22'],
                                               ['Blade root 1 My', '22'],
                                               ['Blade root 1 Mz', '22'],
                                               ['Blade root 1 Fx', '22'],
                                               ['Blade root 1 Fy', '22'],
                                               ['Blade root 1 Fz', '22']]}
                            post_path = os.path.join(path_bs, 'br')
                            Bstats(path_run, fat_list, chan_dict, post_path)

                        if self.var_result.yb_flag:
                            chan_dict = {'yb':[['Yaw bearing Mx', '24'],
                                               ['Yaw bearing My', '24'],
                                               ['Yaw bearing Mz', '24'],
                                               ['Yaw bearing Fx', '24'],
                                               ['Yaw bearing Fy', '24'],
                                               ['Yaw bearing Fz', '24']]}
                            post_path = os.path.join(path_bs, 'yb')
                            Bstats(path_run, fat_list, chan_dict, post_path)

                        if self.var_result.rs_flag:
                            chan_dict = {'rs':[['Rotor speed', '05']]}
                            post_path = os.path.join(path_bs, 'rs')
                            Bstats(path_run, fat_list, chan_dict, post_path, only_maxmin=False)
                            chan_dict = {'rs':[['Rotor speed', '05']]}
                            post_path = os.path.join(path_bs, 'rs_all')
                            dlc_list  = fat_list+ult_list
                            Bstats(path_run, dlc_list, chan_dict, post_path)
                        self.post_list.append(path_bs)

                    # 03_Clearance
                    if self.cbx13.isChecked():
                        print('begin to handle tower clearance')
                        path_tc = os.path.join(self.lin4.text(), '03_Clearance')
                        if self.var_result.tc_flag:
                            # dynamic
                            ult_chan = {'tc':self.var_result.tc_var_list}
                            path_dyn = os.path.join(path_tc, 'dynamic')
                            Ultimate(path_run, path_dyn, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)

                            # redirect dynamic tower clearance file
                            # for file in os.listdir(path_dyn):
                            #     file_path = os.path.join(path_dyn, file)
                            #     if os.path.isfile(file_path):
                            #         os.remove(file_path)

                            tc_path = os.path.join(path_dyn, 'tc').replace('/','\\')
                            pj_path = os.path.join(tc_path, 'tc.$PJ')
                            shutil.copy(pj_path, path_dyn)

                            in_path = os.path.join(tc_path,  'dtsignal.in')
                            shutil.copy(in_path, path_dyn)
                            shutil.rmtree(tc_path)

                            in_ori = os.path.join(path_dyn, 'dtsignal.in')

                            with open(in_ori) as f1:
                                content = f1.read()
                                content = content.replace('PATH %s\n'%tc_path, 'PATH %s\n'%path_dyn)
                            with open(in_ori, 'w+') as f2:
                                f2.write(content)

                        # steady
                        path_std = os.path.join(path_tc, 'steady')
                        Steadytc(self.lin1.text(), path_std)
                        self.post_list.append(path_tc)

                    # 04_Combination
                    if self.cbx12.isChecked():
                        path_comb = os.path.join(self.lin4.text(), '04_Combination')
                        if self.var_result.br_flag:
                            chan_dict = {'br': ['Blade root 1 Mx', 'Blade root 1 My']}
                            Combination(path_run, path_comb, fat_list, chan_dict, '22', '700')

                        if self.var_result.bs_flag:
                            chan_dict = {'brs1': ['Blade 1 Mx (Root axes)', 'Blade 1 My (Root axes)']}
                            Combination(path_run, path_comb, fat_list, chan_dict, '41', '800')

                            if len(self.var_result.brs_var_list)>8:
                                chan_dict = {'brs2': ['Blade 2 Mx (Root axes)', 'Blade 2 My (Root axes)']}
                                path_comb = os.path.join(self.lin4.text(), '04_Combination')
                                Combination(path_run, path_comb, fat_list, chan_dict, '42', '810')

                                chan_dict = {'brs3': ['Blade 3 Mx (Root axes)', 'Blade 3 My (Root axes)']}
                                path_comb = os.path.join(self.lin4.text(), '04_Combination')
                                Combination(path_run, path_comb, fat_list, chan_dict, '43', '820')
                        self.post_list.append(path_comb)

                    # 05_LDD
                    if self.cbx15.isChecked():
                        print('begin to handle ldd')
                        path_ldd = os.path.join(self.lin4.text(), '05_LDD')
                        if self.var_result.hs_flag:
                            ldd_chan = {'hs':[self.var_result.hs_fat_list]}
                            Probability(path_run, path_ldd, path_lct, fat_list, num_bins, des_life, eq_cycle, ldd_chan, '1')
                            Probability(path_run, path_ldd, path_lct, fat_list, '144', des_life, eq_cycle, ldd_chan, '1')

                        if self.var_result.br_flag:
                            ldd_chan = {'br':[self.var_result.br_fat_list[0]]}
                            Probability(path_run, path_ldd, path_lct, fat_list, num_bins, des_life, eq_cycle, ldd_chan, '1')

                        if self.var_result.yb_flag:
                            ldd_chan = {'yb':[self.var_result.yb_fat_list]}
                            Probability(path_run, path_ldd, path_lct, fat_list, num_bins, des_life, eq_cycle, ldd_chan, '1')

                        if self.var_result.hr_flag:
                            ldd_chan = {'hr':[self.var_result.hr_fat_list]}
                            Probability(path_run, path_ldd, path_lct, fat_list, num_bins, des_life, eq_cycle, ldd_chan, '1')
                            Probability(path_run, path_ldd, path_lct, fat_list, '144', des_life, eq_cycle, ldd_chan, '1')
                        self.post_list.append(path_ldd)

                    # 06_LRD
                    if self.cbx14.isChecked():
                        print('begin to handle lrd')
                        path_lrd = os.path.join(self.lin4.text(), '06_LRD')
                        if (self.var_result.br_flag and self.var_result.ws_flag):
                            lrd_chan = {'br':[self.var_result.br_fat_list[0]],
                                        'br_mxy':[['Blade root 1 Mxy']]}
                            lrd_angle = {'br': ['07', 'Mean pitch angle'],
                                         'br_mxy': ['07', 'Mean pitch angle']}
                            Probability(path_run, path_lrd, path_lct, fat_list, '64',  des_life, eq_cycle, lrd_chan, '2', lrd_angle)
                            Probability(path_run, path_lrd, path_lct, fat_list, '144', des_life, eq_cycle, lrd_chan, '2', lrd_angle)

                        if (self.var_result.hs_flag and self.var_result.rs_flag):
                            lrd_chan = {'hs':[self.var_result.hs_fat_list]}
                            lrd_angle = {'hs':['05','Rotor azimuth angle']}
                            Probability(path_run, path_lrd, path_lct, fat_list, '64',  des_life, eq_cycle, lrd_chan, '2', lrd_angle)
                            Probability(path_run, path_lrd, path_lct, fat_list, '144', des_life, eq_cycle, lrd_chan, '2', lrd_angle)
                        self.post_list.append(path_lrd)

                # component
                if self.cb_cmp.isChecked():

                    # ultimate
                    mean_ind  = self.lin5.text().strip()
                    half_ind  = self.lin6.text().strip()
                    # fatigue
                    num_bins = self.lin7.text()
                    des_life = self.lin8.text()
                    eq_cycle = self.lin9.text()+'E+07'

                    path_ult = os.path.join(self.lin4.text(), '07_Ultimate')
                    path_fat = os.path.join(self.lin4.text(), '08_Fatigue')

                    if self.cbx_etm.isChecked():
                        self.etm_opt = True
                    else:
                        self.etm_opt = False

                    print('begin to handle component')
                    # ************************ 06_Fatigue  *********************************************
                    # 01_BRS
                    if self.var_result.bs_flag and self.cbx16.isChecked():
                        print('begin to write 01_BRS')
                        fat_chan = dict({'brs1': [self.var_result.brs_fat_list[0], self.var_result.bld_radius]})
                        path     = os.path.join(path_fat, r'01_BRS\brs1')
                        Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        if len(self.var_result.brs_fat_list) == 3:
                            fat_chan = dict({'brs2': [self.var_result.brs_fat_list[1], self.var_result.bld_radius]})
                            path = os.path.join(path_fat, r'01_BRS\brs2')
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)

                            fat_chan = dict({'brs3': [self.var_result.brs_fat_list[2], self.var_result.bld_radius]})
                            path = os.path.join(path_fat, r'01_BRS\brs3')
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        self.post_list.append(os.path.join(path_fat, '01_BRS'))

                    # 02_BUS
                    if self.var_result.bu_flag and self.cbx16.isChecked():
                        print('begin to write 02_BUS')
                        fat_chan = dict({'bus1': [self.var_result.bus_fat_list[0], self.var_result.bld_radius]})
                        path     = os.path.join(path_fat, r'02_BUS\bus1')
                        Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)

                        if len(self.var_result.bus_fat_list)==3:
                            fat_chan = dict({'bus2': [self.var_result.bus_fat_list[1], self.var_result.bld_radius]})
                            path = os.path.join(path_fat, r'02_BUS\bus2')
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)

                            fat_chan = dict({'bus3': [self.var_result.bus_fat_list[2], self.var_result.bld_radius]})
                            path = os.path.join(path_fat, r'02_BUS\bus3')
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        self.post_list.append(os.path.join(path_fat, '02_BUS'))

                    # 03_BR_Mxy_Ang
                    if self.var_result.br_mxy_flag and self.cbx16.isChecked():
                        print('begin to write 03_BR_Mxy_Seg')
                        fat_chan = dict({'br_comb': [self.var_result.br_mxy_list]})
                        # print(self.var_result.br_mxy_list)
                        path     = os.path.join(path_fat, '03_BR_Mxy_Seg')
                        Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        self.post_list.append(path)

                    # 04_BRS_Mxy_Ang
                    if self.cbx16.isChecked():
                        print('begin to write 04_BRS_Mxy_Seg')
                        path = os.path.join(path_fat, '04_BRS_Mxy_Seg')
                        if self.var_result.brs1_mxy_flag:
                            fat_chan = dict({'brs1_comb': [self.var_result.brs1_mxy_list]})
                            # print(self.var_result.brs_mxy_list)
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)

                        if self.var_result.brs2_mxy_flag:
                            fat_chan = dict({'brs2_comb': [self.var_result.brs2_mxy_list]})
                            # print(self.var_result.brs_mxy_list)
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)

                        if self.var_result.brs3_mxy_flag:
                            fat_chan = dict({'brs3_comb': [self.var_result.brs3_mxy_list]})
                            # print(self.var_result.brs_mxy_list)
                            Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        self.post_list.append(path)

                    # 06_Tower
                    if self.var_result.tr_flag and self.cbx17.isChecked() and result[2]:
                        print('begin to write 06_Tower')
                        fat_chan = dict({'tr': [self.var_result.tr_fat_list, result[2]]})
                        path     = os.path.join(path_fat, '06_Tower')
                        Fatigue(path_run, path, path_lct, fat_list, num_bins, des_life, eq_cycle, fat_chan)
                        self.post_list.append(path)
                    elif self.cbx17.isChecked() and not result[2]:
                        raise Exception('Please choose tower section for Tower Fatigue option!')

                    # 07_Foundation
                    if self.var_result.tr_flag and self.cbx18.isChecked() and result[2]:
                        print('begin to write 07_Foundation')
                        fat_chan = dict({'tr': [self.var_result.tr_fat_list, result[2]]})
                        path     = os.path.join(path_fat, '07_Foundation')
                        Fatigue(path_run, path, path_lct, fat_list, 40, des_life, eq_cycle, fat_chan)
                        self.post_list.append(path)
                    elif self.cbx18.isChecked() and not result[2]:
                        raise Exception('Please choose tower section for Foundation Fatigue option!')

                    # ************************  07_Ultimate   ********************************************************
                    if self.var_result.bs_flag and self.cbx19.isChecked():
                        ult_chan = dict({'brs': [self.var_result.brs_var_list, self.var_result.bld_radius]})
                        # 01_BRS_Inclsf
                        path     = os.path.join(path_ult, '01_BRS_Inclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)
                        # 02_BRS_Exclsf
                        path     = os.path.join(path_ult, '02_BRS_Exclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)

                    if self.var_result.bu_flag and self.cbx19.isChecked():
                        ult_chan = dict({'bus': [self.var_result.bus_var_list, self.var_result.bld_radius]})
                        # 03_BUS_Inclsf
                        path     = os.path.join(path_ult, '03_BUS_Inclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)
                        # 04_BUS_Exclsf
                        path     = os.path.join(path_ult, '04_BUS_Exclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)

                    if self.var_result.hr_flag and self.cbx20.isChecked():
                        ult_chan = dict({'hr': self.var_result.hr_all_list})
                        # 05_HR_BR_Inclsf
                        path     = os.path.join(path_ult, '05_HR_BR_Inclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)

                    if self.var_result.hr_flag and self.cbx20.isChecked():
                        ult_chan = dict({'hr': self.var_result.hr_var_list,
                                         'hs': self.var_result.hs_var_list})
                        ult_dlc8 = [lc for lc in ult_list if 'DLC8' in lc]
                        # 06_HR_onlydlc8
                        path     = os.path.join(path_ult, '06_HR_Onlydlc8\Inclsf')
                        Ultimate(path_run, path, ult_dlc8, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)
                        path     = os.path.join(path_ult, '06_HR_Onlydlc8\Exclsf')
                        Ultimate(path_run, path, ult_dlc8, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)
                        # 07_HR_wodlc8_Exclsf
                        path     = os.path.join(path_ult, '07_HR_Wodlc8\Inclsf')
                        ult_wo8x = [dlc for dlc in ult_list if 'DLC8' not in dlc.upper()]
                        self.post_list.append(path)
                        Ultimate(path_run, path, ult_wo8x, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        path     = os.path.join(path_ult, '07_HR_Wodlc8\Exclsf')
                        Ultimate(path_run, path, ult_wo8x, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)

                    if self.var_result.tr_flag and self.cbx21.isChecked() and result[2]:
                        ult_chan = dict({'tr': [self.var_result.tr_var_list, result[2]]})
                        # 10_Tower_Inclsf
                        print('begin to write 10_Tower_Inclsf')
                        path     = os.path.join(path_ult, '10_Tower_Inclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)
                        # 11_Tower_Exclsf
                        print('begin to write 11_Tower_Exclsf')
                        path     = os.path.join(path_ult, '11_Tower_Exclsf')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)
                        # 12_Tower_dlc12
                        print('begin to write 10_Tower_dlc12')
                        path     = os.path.join(path_ult, '12_Tower_dlc12')
                        dlc_list = [lc for lc in fat_list if lc=='DLC12']
                        Ultimate(path_run, path, dlc_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, False)
                        self.post_list.append(path)

                    elif self.cbx21.isChecked() and not result[2]:
                        raise Exception('Please choose tower section for Tower Ultimate option!')

                    if self.var_result.nac_flag and self.cbx22.isChecked():
                        ult_chan = dict({'nac': self.var_result.nac_acc_list})
                        # 10_Tower_Inclsf
                        print('begin to write 13_Nacelle_Acc')
                        path = os.path.join(path_ult, '13_Nacelle_Acc')
                        Ultimate(path_run, path, ult_list, ult_chan, self.sf_flag, self.etm_opt, mean_ind, half_ind, True)
                        self.post_list.append(path)
                return True
            else:
                return False

        except Exception as e:
            QMessageBox.about(self, 'Window', 'Error occurs when generating post file\n%s'%e)
            print('Error occurs when generating post file\n%s'%e)
            return False

    def write_joblist(self):

        res = self.handle()

        if self.post_list and res:
            print('\nbegin to write joblist')
            try:
                joblist_dir = os.path.join(self.lin4.text(), '00_Joblist')
                if not os.path.isdir(joblist_dir):
                    os.makedirs(joblist_dir)
                file_list   = [file for file in os.listdir(joblist_dir) if 'Post' in file and 'joblist' in file]

                post_name = 'Post_%s' %(len(file_list)+1)
                jb_res = Create_Joblist(post_path=self.post_list,
                                        joblist_name=post_name,
                                        joblist_path=joblist_dir)

                joblist_path = os.path.join(joblist_dir, post_name+'.joblist')

                if os.path.isfile(joblist_path):
                    QMessageBox.about(self, 'Window', 'Joblist is done!')

            except Exception as e:
                QMessageBox.about(self, 'Window', 'Error occurs when generating joblist\n%s' % e)

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        # screen = QDesktopWidget().screenGeometry()
        # form   = self.geometry()
        # x_move_step = (screen.width()-form.width())/2
        # y_move_step = (screen.height()-form.height())/2
        # self.move(x_move_step, y_move_step)

if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = PostWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())