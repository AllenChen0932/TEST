#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/27/2020 5:52 PM
# @Author  : CE
# @File    : delete_file.py

import os, sys
import shutil
import multiprocessing as mp

from PyQt5.QtWidgets import (QMainWindow,
                             QApplication,
                             QDesktopWidget,
                             QLabel,
                             QLineEdit,
                             QPushButton,
                             QComboBox,
                             QFileDialog,
                             QGridLayout,
                             QWidget,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

class Delete_File(object):
    
    def __init__(self, path, extension=None):
        
        self.run_path = path
        self.ext_list = extension

    # get children path under run
    def get_children_path(self,run_path):
    
        path_list = []
        dlc_list  = [dlc for dlc in os.listdir(run_path) if os.path.isdir(os.path.join(run_path, dlc))]
    
        for dlc in dlc_list:
    
            dlc_path = os.path.join(run_path, dlc)
            lc_list  = os.listdir(dlc_path)
    
            for lc in lc_list:
                lc_path = os.path.join(dlc_path, lc)
    
                if os.path.isdir(lc_path):
                    path_list.append(lc_path)
    
        return path_list

    def remove_except_pj_in(self, path):
        # extension in capital
        extension_list  = ['.$PJ', '.IN']
    
        file_list = os.listdir(path)
        # print(file_list)
        for file in file_list:
    
            file_path = os.path.join(path, file)
            extension = os.path.splitext(file)[1]
    
            if extension.upper() not in extension_list:
                os.remove(file_path)
                print(file_path)
    
    def remove_extension(self, path, ext_list):
        extension_list = [(os.extsep+ext).upper() if not ext.startswith(os.extsep) else ext.upper() for ext in ext_list]
        file_list = os.listdir(path)
        # print(file_list)
        for file in file_list:
    
            file_path = os.path.join(path, file)
            extension = os.path.splitext(file)[1]
    
            if extension.upper() in extension_list:
                os.remove(file_path)
                print(file_path)
    
    def remove_all(self, path):
    
        file_list = os.listdir(path)
        # print(file_list)
        for file in file_list:
            file_path = os.path.join(path, file)
    
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(file_path)
                
        os.rmdir(path)
    
    def delete_run_except_pj(self):
        
        if os.path.isdir(self.run_path):
            pool = mp.Pool(processes=mp.cpu_count())
    
            path_list = self.get_children_path(self.run_path)
            print(path_list)
    
            pool.map_async(self.remove_except_pj_in, path_list)
    
            pool.close()
            pool.join()
        else:
            print('Please choose a right path!')
    
    def delete_ext(self):
    
        if os.path.isdir(self.run_path):
    
            pool = mp.Pool(processes=mp.cpu_count())
    
            path_list = self.get_children_path(self.run_path)
            print(path_list)
    
            for path in path_list:
                pool.apply_async(self.remove_extension, args=(path, self.ext_list))
    
            pool.close()
            pool.join()

    def delete_run(self):
        
        pool = mp.Pool(processes=mp.cpu_count())
    
        path_list = self.get_children_path(self.run_path)
        path_list.append(self.run_path)
        print(path_list)
    
        pool.map_async(self.remove_all, path_list)
    
        pool.close()
        pool.join()
    
        if os.listdir(self.run_path):
            for file in os.listdir(self.run_path):
                file_path = os.path.join(self.run_path, file)
                os.remove(file_path)

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

class DeleteWindow(QMainWindow):
    def __init__(self, parent=None):

        super(DeleteWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 120)
        self.setWindowTitle('Delete')
        self.setWindowIcon(QIcon('./Icon/Text.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Run Path")
        self.lbl1.setFont(QFont("SimSun", 8))

        self.lbl2 = QLabel("Extension")
        self.lbl2.setFont(QFont("SimSun", 8))
        #
        self.lbl3 = QLabel('Delete Type')
        self.lbl3.setFont(QFont("SimSun", 8))

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(QFont("SimSun", 8))
        self.lin1.setPlaceholderText("Choose run path")

        self.lin2 = QLineEdit()
        self.lin2.setFont(QFont("SimSun", 8))
        self.lin2.setPlaceholderText('Seprated by ",“')

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("SimSun", 8))
        self.btn1.clicked.connect(self.choose_run_path)

        self.btn2 = QPushButton("Run delete")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.run_delete)

        self.cbx1 = QComboBox()
        self.cbx1.setFont(QFont("SimSun", 8))
        # self.cbx1.setDisabled(True)
        self.cbx1.currentTextChanged.connect(self.type_action)
        self.cbx1.addItem("Please select")
        self.cbx1.addItem("All below")
        self.cbx1.addItem("Except pj/in")
        self.cbx1.addItem("By extension")

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 0, 0, 1, 1)
        self.grid.addWidget(self.lin1, 0, 1, 1, 4)
        self.grid.addWidget(self.btn1, 0, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 1, 0, 1, 1)
        self.grid.addWidget(self.lin2, 1, 1, 1, 4)
        self.grid.addWidget(self.lbl3, 2, 0, 1, 1)
        self.grid.addWidget(self.cbx1, 2, 1, 1, 4)
        self.grid.addWidget(self.btn2, 3, 0, 1, 6)

        self.mywidget.setLayout(self.grid)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    # @pysnooper.snoop()     此处加这句之后程序直接退出
    def choose_run_path(self):
        run_path = QFileDialog.getExistingDirectory(self, "Choose path dialog", r".")

        if run_path:
            self.lin1.setText(run_path.replace('/', '\\'))
            self.lin1.home(True)

    def type_action(self):

        if self.cbx1.currentText() == 'By extension':
            self.lin2.setDisabled(False)
        else:
            self.lin2.setDisabled(True)

    def run_delete(self):

        run_path = self.lin1.text()
        ext_list = self.lin2.text().split(',')
        option   = self.cbx1.currentText()

        try:
            if os.path.isdir(run_path):
    
                if option == 'All below':
                    Delete_File(run_path).delete_run()

                elif option == 'Except pj/in':
                    Delete_File(run_path).delete_run_except_pj()

                elif option == 'By extension':
                    if ext_list:
                        Delete_File(run_path, ext_list).delete_ext()
                    else:
                        QMessageBox.about(self, 'Warning', 'Please define extension first!')
                else:
                    QMessageBox.about(self, 'Warning', 'Please choose delete type first!')

            else:
                QMessageBox.about(self, 'Warning', 'Please define run path first!')

        except Exception as e:
            QMessageBox.about(self, 'Warnning', '%s' %e)

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':

    mp.freeze_support()

    app = QApplication(sys.argv)
    window = DeleteWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())
