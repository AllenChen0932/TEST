# -*- coding: utf-8 -*-
# @Time    : 2019/10/13 15:35
# @Author  : CE
# @File    : bladed_ldd.py
# 程序假设默认全风速下的ldd与各风速下的ldd变量的数量和顺序一致
# 默认统计的bin的个数为64, 文件格式为二进制

import os
import re
import numpy as np

class read_bladed(object):

    def __init__(self, name, path, ext):
        '''
        read bladed percent file and dollar file
        :param name: pj name
        :param path: dlc path
        :param ext: file extension
        '''

        self.pj_name = name
        self.pj_path = path
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

        path = self.pj_path + os.sep + self.pj_name + '.%' + self.ext
        # print(path)

        with open(path, 'r') as f:

            for line in f.readlines():

                if 'ACCESS' in line:
                    self.ext_format = line.strip("\n\t").split()[-1]
                    # print(self.ext_format)
                    continue

                if 'RECL' in line:
                    self.ext_type = line.strip().split()[-1]
                    continue

                regex = re.compile(r'\bDIMENS\b\s+(\d+)\s+(\d+)')
                match = regex.search(line)

                if match != None:
                    self.ext_dim = line.strip().split()[1:]
                    # print(self.ext_dim)
                    continue

                if 'VARIAB' in line:
                    temp = line.strip().split("'")

                    for i in temp:
                        if i in ' ' or '':
                            temp.remove(i)

                    self.ext_var = temp[1:]
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

                    self.ext_data = self.data.reshape((int(self.ext_dim[1]),
                                                       int(self.ext_dim[2]), int(self.ext_dim[0])))

                    self.data_len = int(self.ext_dim[2])

        elif self.ext_format == 'S':

            with open(path, 'r') as f:

                if self.ext_type == '4':

                    self.data = np.loadtxt(f, np.float32)

                elif self.ext_type == '8':

                    self.data = np.loadtxt(f, np.float64)

                if len(self.ext_dim) == 2:

                    self.ext_data = self.data

                    self.data_len = int(self.ext_dim[1])
                    # print(self.data_len)

                elif len(self.ext_dim) == 3:

                    self.ext_data = self.data.reshape((int(self.ext_dim[1]),
                                                       int(self.ext_dim[2]), int(self.ext_dim[0])))

                    self.data_len = int(self.ext_dim[2])

class BladedLDD(object):

    def __init__(self, ldd_path):

        self.ldd_path = ldd_path

        self.chan_pj  = {}       # channels: project list
        self.prj_ext  = {}       # project: extension list
        self.prj_var  = {}       # pj name: variable

        self.ch_pj_var_data = {} # channel:{pj name: {variable: data}}

        self.get_ldd_pj()
        self.get_ldd_result()

    def get_ldd_pj(self):
        '''
        get the ldd project file under path and eliminate other project
        :param path: ldd path
        :return: ldd project file
        '''

        file_list = os.listdir(self.ldd_path)
        # only hr_64
        channels  = [file for file in file_list if os.path.isdir(os.path.join(self.ldd_path, file)) and 'hr_64' in file]
        if not channels:
            raise Exception('No hr_64 ldd results under %s!' %self.ldd_path)
        # print(channels)

        for chan in channels:
            # get pj for each channel
            chan_path = os.path.join(self.ldd_path, chan)
            files     = os.listdir(chan_path)

            ldd_pj = [file.split('.')[0] for file in files if '.$PJ' in file.upper()]

            # eliminate revs and other project
            temp = []
            for pj in ldd_pj:

                pj_file = ''.join([pj, '.$PJ'])
                pj_path = os.path.join(chan_path, pj_file)

                with open(pj_path, 'r') as f:
                    content = f.read()
                    if 'MSTART PROBD' in content and 'AZNAME	"Revs"' not in content:
                        temp.append(pj)

            # eliminate channel with other result
            if temp:
                self.chan_pj[chan] = temp

        # all extensions for each LDD
        for chan, pj_list in self.chan_pj.items():

            chan_path = os.path.join(self.ldd_path, chan)
            res_list  = os.listdir(chan_path)

            for pj_name in pj_list:

                temp = [file for file in res_list if file.startswith(pj_name)]

                percent_list = [file.split('%')[1] for file in temp if (file.find('%') != -1)]
                self.prj_ext[pj_name] = percent_list

    def get_ldd_result(self):

        # get extension for each channel
        for chan, pj_list in self.chan_pj.items():
            # get project name
            pj_name = [pj.split('.')[0] for pj in pj_list]
            pj_name.sort()

            ext_var  = {}     # extension: variable
            var_data = {}     # variable : data
            prj_var  = {}     # pj_name  : variable

            for pj in pj_name:
                # ext: var
                for ext in self.prj_ext[pj]:
                    # percent file path
                    percent_path = os.path.join(self.ldd_path, chan, pj+'.%'+ext)

                    with open(percent_path, 'r') as f:
                        ext_var[ext] = [line.split("'")[-2] for line in f.readlines() if 'AXISLAB' in line][0]

                    res = read_bladed(pj, os.path.join(self.ldd_path,chan), ext)
                    # time at level per bin
                    var_data[ext_var[ext]] = res.ext_data

                prj_var[pj] = var_data

                self.prj_var[pj] = ext_var.values()

            self.ch_pj_var_data[chan] = prj_var

if __name__ == '__main__':

    path = r'E:\05 TASK\02_Tools\01_load summary append_v1\post\LDD'

    BladedLDD(path)