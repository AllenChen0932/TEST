# ！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2018/8/26 11:14
# @Author  : CE
# @File    : joblist.py
'''
1 本程序只适用于通过additional controller parameters对话框来设置控制器参数, 不适用于有多个控制器时的模型；
2 有多个控制器的模型需要在每个控制器中定义单独的参数文件；
'''

import sys, os
import datetime
import configparser
import multiprocessing as mp


from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QTextEdit,
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
                             QGroupBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore
# import pysnooper

# 修改joblist中输出路径
class modify_joblist(object):
    def __init__(self,
                 joblist,
                 new_list,
                 run_path,
                 dll_path,
                 xml_path,
                 newlist_path,
                 add_con_para,
                 other_con):
        '''
        定义输入输出
        :param joblist: 原始joblist的引用路径
        :param new_list: 新joblist的名称
        :param jobs_path: 将原来的jobs按照工况名生成新的jobs文件的路径
        :param run_path: 模型中仿真输出结果所在的run的路径
        :param dll_path: 新的dll文件的引用路径
        :param xml_path: 新的xml文件的引用路径

        '''

        self.job_list = joblist.replace('/','\\')
        self.new_list = new_list
        self.dll_path = dll_path.replace('/','\\') if dll_path else None
        self.xml_path = xml_path.replace('/','\\') if xml_path else None
        self.run_path = run_path.replace('/','\\')
        self.njl_path = newlist_path
        self.add_cont = add_con_para
        self.oth_cont = other_con

        self.old_path = None
        self.dlc_jobpath = {}  # dlc name：job old path
        self.rela_path = None

        self.get_jobpath()
        self.create_jobs()
        self.modify_joblist()

    def get_common_path(self, path1, path2):
        '''get common path for jobs defined in joblist'''

        list1 = path1.split(os.sep)
        list2 = path2.split(os.sep)
        for index, string in enumerate(list1):
            if string != list2[index]:
                if index > 1:
                    print(list1[:index])
                    return os.sep.join(list1[:index])
                else:
                    raise Exception('%s and %s are located in different path!'%(path1, path2))

    def get_jobpath(self):
        '''get jobname and job path defined in joblist'''

        print('Start to read joblist and get dlc name and path!')
        with open(self.job_list, 'r') as f:
            lines = f.readlines()

            # get relative job name for each dlc
            dlc_list = [line.split('>')[1].split('<')[0].split(os.sep)[-1]
                        for line in lines if line.endswith('</InputFileDirRelativeToBatch>\n')]
            # get job path for each dlc
            job_path = [line.split('>')[1].split('<')[0] for line in lines if line.endswith('</InputFileDir>\n')]
            # print(dlc_list)
            # print(job_path)
            self.dlc_jobpath = dict(zip(dlc_list, job_path))
            # print(self.dlc_jobpath)
            # get run path defined in joblist
            self.old_run_path = self.get_common_path(job_path[0], job_path[-1])
            print('Old run path: ', self.old_run_path)
        print('Read joblist is done!')

    def get_control_para(self, control_parameters):
        '''get control parameters'''

        def part_func(list, n):
            '''
            将list按照固定长度进行分割
            :param list:
            :param n:
            :return:
            '''
            new_list = []
            for i in range(0, len(list), n):
                temp = list[i:i + n]
                new_list.append(temp)

            return new_list

        para_name  = []
        para_value = []

        control_parameters = control_parameters.replace('<', '&lt;')
        control_parameters = control_parameters.replace('>', '&gt;')

        para_cont = control_parameters.split('\n')

        for i in range(len(para_cont)):
            para_cont[i] = para_cont[i].strip() + '\n'
        # print(para_cont)

        for i in range(len(para_cont)):
            if i%5 == 1:
                para_name.append(para_cont[i])
            if i%5 == 3:
                para_value.append(para_cont[i])

        return para_name, para_value, part_func(para_cont, 5)

    def get_content_para(self, other_parameters):
        '''get content to be replaced and new content'''
        
        cont_value = {}
        cont_para  = other_parameters.split(',;')

        for i in range(len(cont_para)):
            if cont_para[i]:
                temp = cont_para[i].split(',')
                old  = temp[0].strip()
                new  = temp[1].strip() if temp[1] else temp[1]

                cont_value[old] = new
        # print(cont_value)
        return cont_value

    def modify_prj(self, prj_path, job_path, dll_path, xml_path, control_parameters, other_parameters):
        '''read prj/PJ file'''

        num_con = 1  # 记录控制器的个数, 可能存在风速控制器, 默认第一个是风机控制器
        num_para = 1  # 记录第一个<AdditionalParameters>, 默认修改第一个参数
        para_index = dict()  # 记录additional parameter介绍的位置

        with open(prj_path, 'r', encoding='ISO-8859-1') as f:
            lines = f.readlines()

            for index, line in enumerate(lines):
                # 如果模型通过READ读入参数, 那么修改读入参数文件的路径
                if '<AdditionalParameters>' in line and xml_path and num_para < 2:
                    lines[index] = line.split('>')[0] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                    num_para += 1

                # 如果之前的模型没有通过READ读取参数, 那么可以通过以下命令增加参数文件的路径
                if '<AdditionalParameters />' in line and xml_path and num_para < 2:
                    lines[index] = line[:-4] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                    num_para += 1

                if '<Filepath>' in line and dll_path and num_con < 2:
                    lines[index] = line.split('>')[0] + '>' + dll_path + '</Filepath>' + '\n'

                    num_con += 1
                    # 修改xml在修改dll之前, 如果dll之前没有读取xml,dll之后的parameters也不用修改
                    num_para += 1

                # 记录所有控制参数的名称和位置
                if '&lt;Name&gt;P_' in line:
                    lines[index] = line.strip() + '\n'
                    para_index[line] = index

        if control_parameters:
            control = self.get_control_para(control_parameters)

            ori_para_names = para_index.keys()
            new_para_names = control[0]
            new_para_value = control[1]

            for i, para_name in enumerate(new_para_names):
                if  para_name in ori_para_names:
                    lines[para_index[para_name]+2] = new_para_value[i]
                else:
                    para_value = control[2][i]
                    for j, v in enumerate(para_value):
                        lines.insert(max(para_index.values())+4+j, para_value[j])

        # new content
        content = ''.join(lines)

        # replace content
        if other_parameters:
            cont_value = self.get_content_para(other_parameters)

            for key, value in cont_value.items():
                content = content.replace(key, value)

        file_path = os.path.join(job_path, os.path.split(prj_path)[1])
        with open(file_path, 'w+', encoding='ISO-8859-1') as f:
            f.write(content)

    def modify_in(self, in_path, job_path, dll_path, xml_path, control_parameters, other_parameters):
        '''modify in file'''

        num_con = 1  # 记录控制器的个数, 可能存在风速控制器, 默认第一个是风机控制器
        num_para = 1  # 记录第一个<AdditionalParameters>, 默认修改第一个参数
        para_index = dict()  # 记录additional parameter介绍的位置

        with open(in_path, 'r', encoding='ISO-8859-1') as f:
            lines = f.readlines()

            for index, line in enumerate(lines):
                # 如果模型通过READ读入参数, 那么修改读入参数文件的路径
                if '<AdditionalParameters>' in line and xml_path and num_para < 2:
                    lines[index] = line.split('>')[0] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                    num_para += 1

                # 如果之前的模型没有通过READ读取参数, 那么可以通过以下命令增加参数文件的路径
                if '<AdditionalParameters />' in line and xml_path and num_para < 2:
                    lines[index] = line[:-4] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                    num_para += 1

                if '<Filepath>' in line and dll_path and num_con < 2:
                    lines[index] = line.split('>')[0] + '>' + dll_path + '</Filepath>' + '\n'

                    num_con += 1
                    # 修改xml在修改dll之前, 如果dll之前没有读取xml,dll之后的parameters也不用修改
                    num_para += 1

                # 记录所有控制参数的名称和位置
                if '&lt;Name&gt;P_' in line:
                    para_index[line] = index

        if control_parameters:
            control = self.get_control_para(control_parameters)

            ori_para_names = para_index.keys()
            new_para_names = control[0]
            new_para_value = control[1]

            for index, para_name in enumerate(new_para_names):
                if  para_name in ori_para_names:
                    lines[para_index[para_name]+2] = new_para_value[index]
                else:
                    para_value = control[2][index]
                    for i, v in enumerate(para_value):
                        lines.insert(max(para_index.values())+4+i, para_value[i])

        # new content
        content = ''.join(lines)

        # replace content
        if other_parameters:
            cont_value = self.get_content_para(other_parameters)

            for key, value in cont_value.items():
                content = content.replace(key, value)

        file_path = os.path.join(job_path, 'dtbladed.in')
        with open(file_path, 'w+', encoding='ISO-8859-1') as f:
            f.write(content)

    def create_jobs(self):
        '''
        将原始的jobs按照工况进行命名, 然后保存到self.job_path下面
        :return:
        '''
        print('Start to create jobs...')
        pool = mp.Pool(processes=mp.cpu_count())
        for key, value in self.dlc_jobpath.items():
            # new path for jobs
            job_path = value.replace(self.old_run_path, self.run_path) if self.run_path else value
            if not os.path.exists(job_path):
                os.makedirs(job_path)

            # write pj
            prj_name = [file for file in os.listdir(value) if '.$PJ' in file.upper()][0]
            prj_path = os.path.join(value, prj_name)
            pool.apply_async(self.modify_prj, args=(prj_path, job_path, self.dll_path, self.xml_path, self.add_cont,
                                                    self.oth_cont))

            # write in
            in_name = [file for file in os.listdir(value) if file.lower().endswith('in')][0]
            in_path = os.path.join(value, in_name)
            pool.apply_async(self.modify_in, args=(in_path, job_path, self.dll_path, self.xml_path, self.add_cont,
                                                   self.oth_cont))
        
        pool.close()
        pool.join()

    def modify_joblist(self):
        '''generate new joblist through new result path and jobs path'''

        print('Start to cteate joblist...')
        content = ''

        with open(self.job_list, 'r') as f:
            for line in f.readlines():

                # 修改job的仿真结果路径
                if '<ResultDir>' in line and self.run_path:
                    temp = line.split('>')
                    old_path = temp[1].split('<')[0]
                    # dlc_path = '\\'.join(old_path.split('\\')[-2:])
                    # new_path = self.run_path.replace('/', '\\') + os.sep + dlc_path
                    new_path = old_path.replace(self.old_run_path, self.run_path)

                    line = temp[0] + '>' + new_path + '</ResultDir>' + '\n'

                if '<InputFileDir>' in line:
                    temp = line.split('>')
                    old_path = temp[1].split('<')[0]
                    self.rela_path = [k for k, v in self.dlc_jobpath.items() if old_path == v][0]

                    job_path = old_path.replace(self.old_run_path, self.run_path) if self.run_path else old_path
                    line = temp[0] + '>' + job_path + '</InputFileDir>' + '\n'

                # 相对路径似乎不起作用
                if '<InputFileDirRelativeToBatch>' in line:

                    temp = line.split('>')
                    line = temp[0] + '>' + self.rela_path + '</InputFileDirRelativeToBatch>' + '\n'
                content += line

        if not self.njl_path:
            job_list_path = os.path.split(self.job_list)[0]
        else:
            job_list_path = self.njl_path

        if self.new_list:
            new_name = self.new_list + '.joblist'
        else:
            date_time = str(datetime.date.today()).replace('-','')
            old_name = os.path.split(self.job_list)[1]
            new_name = '%s_%s.joblist' %(os.path.splitext(old_name)[0], date_time)

        jb_path = os.path.join(job_list_path, new_name)
        with open(jb_path, 'w+') as f:
            f.write(content)
        print('Joblist is done')

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

class JoblistWindow(QMainWindow):

    def __init__(self, parent=None):

        super(JoblistWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.setMinimumSize(800, 630)
        self.setWindowTitle('Modify Joblist')
        self.setWindowIcon(QIcon(".\icon\Text_Edit.ico"))
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

        # input and output
        self.label1 = QLabel("Origninal joblist")
        self.label1.setFont(self.cont_font)

        self.line1 = MyQLineEdit()
        self.line1.setFont(self.cont_font)
        # self.line1.setPlaceholderText("Pick joblist file")

        self.btn1 = QPushButton("...")
        self.btn1.setFont(self.cont_font)
        self.btn1.clicked.connect(self.load_job_list)

        self.label2 = QLabel("Result path")
        self.label2.setFont(self.cont_font)

        self.line2 = MyQLineEdit()
        self.line2.setFont(self.cont_font)
        self.line2.setPlaceholderText("If blank, original result path will be used")
        self.line2.setToolTip('result path for new simulation')

        self.btn2 = QPushButton("...")
        self.btn2.setFont(self.cont_font)
        self.btn2.clicked.connect(self.load_res_path)

        self.label3 = QLabel("New jobs Path")
        self.label3.setFont(self.cont_font)

        self.line3 = MyQLineEdit()
        self.line3.setFont(self.cont_font)
        self.line3.setToolTip('new jobs path for the new joblist')

        self.btn3 = QPushButton("...")
        self.btn3.setFont(self.cont_font)
        self.btn3.clicked.connect(self.load_job_path)

        # controller parameters
        self.label5 = QLabel("DLL Path")
        self.label5.setFont(self.cont_font)

        self.line5 = MyQLineEdit()
        self.line5.setFont(self.cont_font)
        # self.line5.setPlaceholderText("Pick DLL file")

        self.btn5 = QPushButton("...")
        self.btn5.setFont(self.cont_font)
        self.btn5.clicked.connect(self.load_dll_path)

        self.label6 = QLabel("XML Path")
        self.label6.setFont(self.cont_font)

        self.line6 = MyQLineEdit()
        self.line6.setFont(self.cont_font)
        # self.line6.setPlaceholderText("Pick XML file")

        self.btn6 = QPushButton("...")
        self.btn6.setFont(self.cont_font)
        self.btn6.clicked.connect(self.load_xml_path)

        # new joblist
        self.label4 = QLabel("Joblist name")
        self.label4.setFont(self.cont_font)

        self.line4 = MyQLineEdit()
        self.line4.setFont(self.cont_font)
        self.line4.setPlaceholderText("If blank, old joblist name with date suffix will be used")

        self.label7 = QLabel("Joblist path")
        self.label7.setFont(self.cont_font)

        self.line7 = MyQLineEdit()
        self.line7.setFont(self.cont_font)
        self.line7.setPlaceholderText("If blank, original joblist path will be used")


        self.btn4 = QPushButton("...")
        self.btn4.setFont(self.cont_font)
        self.btn4.clicked.connect(self.load_newpath)

        self.label8 = QLabel("Controller\nparameters")
        self.label8.setWordWrap(True)
        self.label8.setFont(self.cont_font)

        self.text1 = QTextEdit()
        self.text1.setFont(self.cont_font)
        self.text1.setPlaceholderText("Additional controller parameters")

        self.btn7 = QPushButton("Clear")
        self.btn7.setFont(self.cont_font)
        self.btn7.clicked.connect(self.clear_control)

        self.label8 = QLabel("Model\nParameters")
        self.label8.setWordWrap(True)
        self.label8.setFont(self.cont_font)

        self.text2 = QTextEdit()
        self.text2.setFont(self.cont_font)
        self.text2.setPlaceholderText("format: \n"
                                      "old, new,;\n"
                                      "e.g. WDIR	 0, WDIR	 1,;")

        self.btn8 = QPushButton("Clear")
        self.btn8.setFont(self.cont_font)
        self.btn8.clicked.connect(self.clear_content)

        self.btn8 = QPushButton("Create jobs and joblist")
        self.btn8.setFont(self.cont_font)
        self.btn8.clicked.connect(self.handle)

        self.sub1 = QGroupBox('Basic')
        self.sub1.setFont(self.title_font)
        self.grid1 = QGridLayout()

        self.grid1.addWidget(self.label1, 0, 0, 1, 1)
        self.grid1.addWidget(self.line1,  0, 1, 1, 5)
        self.grid1.addWidget(self.btn1,   0, 6, 1, 1)
        # self.grid1.addWidget(self.label3, 2, 0, 1, 1)
        # self.grid1.addWidget(self.line3,  2, 1, 1, 5)
        # self.grid1.addWidget(self.btn3,   2, 6, 1, 1)
        self.sub1.setLayout(self.grid1)

        # new joblist
        self.sub2 = QGroupBox('Result parameters')
        self.sub2.setFont(self.title_font)
        self.grid3 = QGridLayout(self.sub2)
        self.grid3.addWidget(self.label2, 0, 0, 1, 1)
        self.grid3.addWidget(self.line2,  0, 1, 1, 5)
        self.grid3.addWidget(self.btn2,   0, 6, 1, 1)
        self.grid3.addWidget(self.label4, 1, 0, 1, 1)
        self.grid3.addWidget(self.line4,  1, 1, 1, 5)
        self.grid3.addWidget(self.label7, 2, 0, 1, 1)
        self.grid3.addWidget(self.line7,  2, 1, 1, 5)
        self.grid3.addWidget(self.btn4,   2, 6, 1, 1)

        self.vbox1  = QVBoxLayout()
        self.vbox1.addWidget(self.sub1)
        self.vbox1.addWidget(self.sub2)

        self.group1 = QGroupBox('Joblist parameters')
        self.group1.setFont(self.title_font)
        self.group1.setLayout(self.vbox1)

        self.group2 = QGroupBox('Project parameters')
        self.group2.setFont(self.title_font)
        self.vbox2 = QVBoxLayout()

        self.sub3 = QGroupBox('Controller')
        self.group2.setFont(self.title_font)
        self.grid2 = QGridLayout()

        self.grid2.addWidget(self.label5, 0, 0, 1, 1)
        self.grid2.addWidget(self.line5,  0, 1, 1, 5)
        self.grid2.addWidget(self.btn5,   0, 6, 1, 1)
        self.grid2.addWidget(self.label6, 1, 0, 1, 1)
        self.grid2.addWidget(self.line6,  1, 1, 1, 5)
        self.grid2.addWidget(self.btn6,   1, 6, 1, 1)
        self.sub3.setLayout(self.grid2)

        # group: additional controller parameters
        self.sub4 = QGroupBox('Add/Modify controller parameters')
        self.sub4.setFont(self.title_font)
        self.grid4 = QGridLayout()

        self.grid4.addWidget(self.text1,  0, 1, 4, 5)
        self.sub4.setLayout(self.grid4)

        # group: replace model parameters
        self.sub5 = QGroupBox('Replacement/Remove')
        self.sub5.setFont(self.title_font)
        self.grid5 = QGridLayout()

        self.grid5.addWidget(self.text2,  0, 1, 4, 5)
        self.sub5.setLayout(self.grid5)
        self.hbox = QHBoxLayout()
        self.hbox.addWidget(self.sub4)
        self.hbox.addWidget(self.sub5)

        self.vbox2.addWidget(self.sub3)
        self.vbox2.addLayout(self.hbox)
        self.group2.setLayout(self.vbox2)

        self.main_layout = QVBoxLayout()
        self.main_layout.addWidget(self.group1)
        self.main_layout.addWidget(self.group2)
        self.main_layout.addWidget(self.btn8)
        self.main_layout.addStretch(1)

        self.mywidget.setLayout(self.main_layout)
        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def save_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if not config.has_section('Joblist'):
            config.add_section('Joblist')

        config['Joblist'] = {'Joblist path':self.line1.text(),
                             'New run path':self.line2.text(),
                             'New joblist name': self.line4.text(),
                             'New joblist path': self.line7.text(),
                             'New DLL path':self.line5.text(),
                             'New XML path':self.line6.text()}

        with open("config.ini",'w') as f:
            config.write(f)

    def load_setting(self):
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')

        if config.has_section('Joblist'):
            self.line1.setText(config.get('Joblist','Joblist path'))
            self.line2.setText(config.get('Joblist','New run path'))
            # self.line3.setText(config.get('Joblist','New Jobs path'))
            self.line5.setText(config.get('Joblist', 'New DLL path'))
            self.line6.setText(config.get('Joblist', 'New XML path'))
            self.line4.setText(config.get('Joblist', 'New joblist name'))
            self.line7.setText(config.get('Joblist', 'New joblist path'))

    def clear_setting(self):
        self.line1.setText('')
        self.line2.setText('')
        # self.line3.setText('')
        self.line5.setText('')
        self.line6.setText('')
        self.line4.setText('')
        self.line7.setText('')

    def user_manual(self):
        os.startfile(os.getcwd()+os.sep+'user manual\Joblist Modify.docx')

    def load_job_list(self):
        self.job_list, filetype1 = QFileDialog.getOpenFileName(self,
                                                               "Choose joblist dialog",
                                                               r".",
                                                               "joblist(*.joblist)")

        if self.job_list:
            self.line1.setText(self.job_list.replace('/', '\\'))
            self.line1.home(True)

    def load_res_path(self):
        self.res_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose result path",
                                                         r".")

        if self.res_path:
            self.line2.setText(self.res_path.replace('/', '\\'))
            self.line2.home(True)

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_job_path(self):
        self.job_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose job path",
                                                         r".")

        if self.job_path:
            self.line3.setText(self.job_path.replace('/', '\\'))
            self.line3.home(True)
            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_dll_path(self):
        self.dll_path, filetype1 = QFileDialog.getOpenFileName(self,
                                                               "Choose dll dialog",
                                                               r".",
                                                               "dll(*.dll)")

        if self.dll_path:
            self.line5.setText(self.dll_path.replace('/', '\\'))
            self.line5.home(True)
            # self.loadcase = os.path.basename(self.dll_path).split('.')[0]

            # QMessageBox.about(self, 'Window', 'Load File Successfully!')

    def load_xml_path(self):
        self.xml_path, filetype1 = QFileDialog.getOpenFileName(self,
                                                               "Choose parameter dialog",
                                                               r".",
                                                               "xml(*.xml)")

        if self.xml_path:
            self.line6.setText(self.xml_path.replace('/', '\\'))
            self.line6.home(True)

    def load_newpath(self):
        self.njl_path = QFileDialog.getExistingDirectory(self,
                                                         "Choose new joblist path",
                                                         r".")

        if self.njl_path:
            self.line7.setText(self.njl_path.replace('/', '\\'))

    def clear_control(self):
        self.text1.clear()

    def clear_content(self):
        self.text2.clear()

    def handle(self):

        self.job_list = self.line1.text().replace('\\', '/')
        self.res_path = self.line2.text().replace('\\', '/')
        self.job_path = self.line3.text().replace('\\', '/')
        self.lis_name = self.line4.text()
        self.dll_path = self.line5.text().replace('\\', '/')
        self.xml_path = self.line6.text().replace('\\', '/')
        self.njl_path = self.line7.text().replace('\\', '/')
        self.othe_con = self.text2.toPlainText()
        self.para_con = self.text1.toPlainText()

        if os.path.isfile(self.job_list) and self.job_list.endswith('joblist'):

            try:
                if self.dll_path or self.xml_path or self.othe_con or self.para_con:
                    if not os.path.isdir(self.res_path):
                        os.makedirs(self.res_path)

                    if not os.listdir(self.res_path):
                        modify_joblist(self.job_list, self.lis_name, self.res_path, self.dll_path,
                                       self.xml_path, self.njl_path, self.para_con, self.othe_con)
                        QMessageBox.about(self, 'Window', 'Modify joblist sucessfully!')
                    else:
                        reply = QMessageBox.information(self, 'Warnning', 'jobs under jobs path will be overwritten?',
                                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                        if reply == QMessageBox.Yes:
                            # print('Start to create...\n')
                            modify_joblist(self.job_list, self.lis_name, self.res_path, self.dll_path,
                                           self.xml_path, self.njl_path, self.para_con, self.othe_con)

                            QMessageBox.about(self, 'Window', 'Modify joblist sucessfully!')

                else:
                    QMessageBox.about(self, 'Window', 'Contents to modify are empty!\n'
                                                      'Please define dll/xml/controller parameters/replace contents first!')
            except Exception as e:
                QMessageBox.about(self, 'Warnning', 'Error occurs!\n%s' %e)

        elif not (os.path.isfile(self.job_list) or self.job_list.endswith('joblist')):
            QMessageBox.about(self, 'Warnning', 'Please make sure joblist are correct!')
        elif not (os.path.isdir(self.job_path)):
            QMessageBox.about(self, 'Warnning', 'Please make sure jobs path are correct!')

    def center(self):

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    app = QApplication(sys.argv)
    window = JoblistWindow()
    # window.setWindowOpacity(0.8)
    window.show()
    sys.exit(app.exec_())


    # 使用方法：
    # 1. 依次定义joblist/dll_path/jobs_path/result_dir/xml_path/new_joblist的路径；
    # 2. 其中joblist、jobs_path、result_dir和new_joblist是必须定义的；
    # 3. dll_path和xml_path可以根据需求进行定义, 如不用修改直接定义空即可；
    # 4. 定义路径时, 请注意引号前的r, 该字符不可或缺。

    # start = time.time()

    # 原joblist的引用路径
    # joblist    = r"E:\python/01_Joblist/batch/Job Lists/dlc12.joblist"
    #
    # # 生成的新的jobs的文件路径
    # jobs_path  = r"E:\python\01_Joblist\batch\Jobs test"
    #
    # # 模型中仿真输出的路径
    #
    # result_dir = r"E:\python\01_Joblist"
    #
    # # 新的joblist名称
    # new_joblist   = 'test'
    #
    # # 需要修改的新的dll文件引用路径
    # dll_path   = r"E:\python\01_Joblist\Controller_loop3.2_V0.0\Win32Discon_loop3.2_v0.0.dll"
    #
    # # 需要修改新的xml文件引用路径
    # xml_path   = r"E:\python\01_Joblist\Controller_loop3.2_V0.0\Parameters.xml"
    #
    # newlist_path = r'E:\python\01_Joblist\batch'


    # '''
    #     def __init__(self,
    #              joblist,
    #              new_list,
    #              jobs_path,
    #              run_path,
    #              dll_path,
    #              xml_path,
    #              new_list_path,
    #              add_con_para,
    #              other_con):
    # '''

    # modify_joblist(joblist, new_joblist, jobs_path, result_dir, dll_path, xml_path, None, None, None)

    # end = time.time()

    # print('Total time: ', end-start)

