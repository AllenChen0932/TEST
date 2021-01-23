# -*- coding: utf-8 -*-
# @Time    : 2020/4/8 21:08
# @Author  : CE
# @File    : Read_Bladed.py

import os
import numpy as np

class read_bladed:

    def __init__(self, pj_name, pj_path, ext):
        '''
        read bladed percent file and dollar file
        :param name: pj name
        :param path: dlc path
        :param ext: file extension
        '''

        self.pj_name = pj_name
        self.pj_path = pj_path
        self.ext     = str(ext)

        self.ext_format = None
        self.ext_type   = None
        self.data_len   = int()
        self.ext_dim    = None
        self.ext_var    = None
        self.ext_data   = None
        self.data       = None

        self.read_percent()
        self.read_dollar()

    def read_percent(self):

        file = ''.join([self.pj_name, '.%', self.ext])
        path = os.sep.join([self.pj_path, file])
        # print(path)

        with open(path, 'r') as f:

            for line in f.readlines():
                # print(line)

                if 'ACCESS' in line:
                    self.ext_format = line.strip("\n\t").split()[-1]
                    print(self.ext_format)
                    continue

                if 'RECL' in line:
                    self.ext_type = line.strip().split()[-1]
                    print(self.ext_type)
                    continue

                if line.startswith('DIMENS'):
                    self.ext_dim = line.strip().split()[1:]
                    print(self.ext_dim)
                    continue

                if 'VARIAB' in line:
                    temp = line.strip().split("'")

                    for i in temp:
                        if i in ' ' or '':
                            temp.remove(i)

                    self.ext_var = temp[1:]
                    print(self.ext_var)
                    break

    def read_dollar(self):

        path = self.pj_path + os.sep + self.pj_name + '.$' + self.ext
        # print(path)

        if self.ext_format == 'D':

            with open(path, 'rb') as f:

                if self.ext_type == '4':

                    self.data = np.fromfile(f, np.float32)
                    # print(self.data)

                elif self.ext_type == '8':

                    self.data = np.fromfile(f, np.float64)

                if len(self.ext_dim) == 2:

                    # print(int(self.ext_dim[ext][1]), int(self.ext_dim[ext][0]))
                    self.ext_data = self.data.reshape((int(self.ext_dim[1]), int(self.ext_dim[0])))
                    # print(self.ext_data)
                    self.data_len = int(self.ext_dim[1])

                elif len(self.ext_dim) == 3:

                    self.ext_data = self.data.reshape((int(self.ext_dim[2]),
                                                       int(self.ext_dim[1]), int(self.ext_dim[0])))

                    self.data_len = int(self.ext_dim[2])

        elif self.ext_format == 'S':

            with open(path, 'r') as f:

                if self.ext_type == '4':

                    self.data = np.loadtxt(f, np.float32)

                elif self.ext_type == '8':

                    self.data = np.loadtxt(f, np.float64)

                if len(self.ext_dim) == 2:

                    self.ext_data = self.data.reshape((int(self.ext_dim[1]), int(self.ext_dim[0])))

                    self.data_len = int(self.ext_dim[1])
                    # print(self.data_len)

                elif len(self.ext_dim) == 3:

                    self.ext_data = self.data.reshape((int(self.ext_dim[2]),
                                                       int(self.ext_dim[1]), int(self.ext_dim[0])))

                    self.data_len = int(self.ext_dim[2])

if __name__ == '__main__':

    pj_name = '12_aa-01'
    pj_path = r'E:\05 TASK\02_Tools\01_load summary append_v1\DLC12\12_aa-01'
    ext     = '09'
    res = read_bladed(pj_name,pj_path,ext).ext_data
    print(res.shape)
    print(res[:10,1,0])