#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/9/2020 4:44 PM
# @Author  : CE
# @File    : copy_file.py

import os, sys
import shutil
import time

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
                             QRadioButton,
                             QMessageBox)
from PyQt5.QtGui import QIcon, QFont
from PyQt5 import QtCore

import multiprocessing as mp

class Copy_File(object):
    def __init__(self, source, result, extension, keep_path=True):

        self.sor_dir   = source
        self.res_dir   = result
        self.ext_list  = extension
        self.keep_path = keep_path

    def copy(self, root, file_list, res_dir, queue):

        if not os.path.exists(res_dir):
            os.makedirs(res_dir)

        for file in file_list:
            file_path = os.path.join(root, file)
            shutil.copy(file_path, res_dir)
            queue.put(file)

    def run(self):

        pool  = mp.Pool(processes=mp.cpu_count())
        queue = mp.Manager().Queue()

        for root, dirs, files in os.walk(self.sor_dir):

            # for ext in extension:
            file_list = [file for file in files for ext in self.ext_list if file.upper().endswith(ext.upper())]
            print(file_list)

            if self.keep_path:
                new_path = root.replace(self.sor_dir, self.res_dir).replace(' ', '_')
            else:
                new_path = self.res_dir

            pool.apply_async(self.copy, args=(root, file_list, new_path, queue))

        pool.close()
        pool.join()

        while True:

            if not queue.empty():
                file_name = queue.get()
                print('%s is done!' %file_name)
            else:
                break

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

class CopyWindow(QMainWindow):

    def __init__(self, parent=None):

        super(CopyWindow, self).__init__(parent)
        self.mywidget = QWidget(self)

        self.resize(600, 120)
        self.setWindowTitle('Delete')
        self.setWindowIcon(QIcon('Icon/copy.ico'))

        self.initUI()
        self.center()

    def initUI(self):

        self.lbl1 = QLabel("Source Path")
        self.lbl1.setFont(QFont("SimSun", 8))

        self.lbl2 = QLabel("Target Path")
        self.lbl2.setFont(QFont("SimSun", 8))
        #
        self.lbl3 = QLabel('Extension')
        self.lbl3.setFont(QFont("SimSun", 8))

        self.lin1 = MyQLineEdit()
        self.lin1.setFont(QFont("SimSun", 8))
        # self.lin1.setPlaceholderText()

        self.lin2 = MyQLineEdit()
        self.lin2.setFont(QFont("SimSun", 8))
        # self.lin1.setPlaceholderText()

        self.lin3 = MyQLineEdit()
        self.lin3.setFont(QFont("SimSun", 8))
        self.lin3.setPlaceholderText('Seprated by ",“')

        self.btn1 = QPushButton("...")
        self.btn1.setFont(QFont("SimSun", 8))
        self.btn1.clicked.connect(self.choose_source_path)

        self.btn2 = QPushButton("Run delete")
        self.btn2.setFont(QFont("SimSun", 8))
        self.btn2.clicked.connect(self.run_copy)

        self.rad1 = QRadioButton('Keep Path')
        self.rad1.setChecked(True)

        self.grid = QGridLayout()

        # 起始行，起始列，占用行，占用列
        self.grid.addWidget(self.lbl1, 0, 0, 1, 1)
        self.grid.addWidget(self.lin1, 0, 1, 1, 4)
        self.grid.addWidget(self.btn1, 0, 5, 1, 1)
        self.grid.addWidget(self.lbl2, 1, 0, 1, 1)
        self.grid.addWidget(self.lin2, 1, 1, 1, 4)
        self.grid.addWidget(self.lbl3, 2, 0, 1, 1)
        self.grid.addWidget(self.lin3, 2, 1, 1, 4)
        self.grid.addWidget(self.rad1, 2, 5, 1, 1)
        self.grid.addWidget(self.btn2, 3, 0, 1, 6)

        self.mywidget.setLayout(self.grid)

        self.setCentralWidget(self.mywidget)

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.close()

    def choose_source_path(self):
        source_path = QFileDialog.getExistingDirectory(self, "Choose path dialog", r".")

        if source_path:
            self.lin1.setText(source_path.replace('/', '\\'))
            self.lin1.home(True)

    def run_copy(self):

        sor_path = self.lin1.text()
        des_path = self.lbl2.text()
        ext_list = self.lin3.text().strip(',')
        option   = True if self.rad1.isChecked() else False

        try:
            if os.path.isdir(sor_path) and os.path.isdir(des_path) and ext_list:
                Copy_File(sor_path, des_path, ext_list, option).run()

            elif not os.path.isdir(sor_path):
                QMessageBox.about(self, 'Warning', 'Please define a right source path!')

            elif not os.path.isdir(des_path):
                QMessageBox.about(self, 'Warning', 'Please define a right target path!')

            elif not ext_list:
                QMessageBox.about(self, 'Warning', 'Please define extension first!')

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
    window = CopyWindow()
    # window.setWindowOpacity(0.9)
    window.show()
    sys.exit(app.exec_())
