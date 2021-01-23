# -*- coding: utf-8 -*-
# @Time    : 2020/4/8 21:08
# @Author  : CE
# @File    : Read_Bladed.py

import os
import numpy as np

class read_bladed(object):

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
        # self.dlc12   = dlc12_path

        self.var_list  = []
        self.ori_units = []
        self.var_unit  = {}

        # self.read_percent()
        # self.read_dollar()
        # self.read_unit()

    def read_percent(self):
        '''
        read dollar file
        :return:
        '''
        file = ''.join([self.pj_name, '.%', self.ext])
        path = os.sep.join([self.pj_path, file])
        # print(path)

        try:
            with open(path, 'r') as f:

                for line in f.readlines():

                    if 'ACCESS' in line:
                        self.format = line.strip().split()[-1]
                        # print(self.format)
                        continue

                    if 'RECL' in line:
                        self.type = line.strip().split()[-1]
                        # print(self.type)
                        continue

                    if line.startswith('DIMENS'):
                        self.dim = line.strip().split()[1:]
                        # print(self.dim)
                        continue

                    if 'VARIAB' in line:
                        temp = line.strip().split("'")

                        self.var_list = temp[1::2]
                        # print(self.var_list)

                        continue

                    if 'VARUNIT' in line:
                        self.ori_units = line.strip().split()[1:]
                        # print(self.ori_units)
                        break

            self.read_unit()

            return self.var_list, self.var_unit

        except IOError:

            print('Open "%s" failed!' %path)

    def read_dollar(self):
        # read percent to get format info
        self.read_percent()

        dollar_path = os.path.join(self.pj_path, ''.join([self.pj_name,'.$',self.ext]))
        # print(path)
        data = None

        if self.format == 'D':
            # binary format
            try:
                with open(dollar_path, 'rb') as f:

                    if self.type == '4':
                        # float32
                        self.data = np.fromfile(f, np.float32)
                        # print(self.data)

                    elif self.type == '8':
                        # float64
                        self.data = np.fromfile(f, np.float64)

                    if len(self.dim) == 2:
                        # 2 dimension
                        data     = self.data.reshape((int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[1])

                    elif len(self.dim) == 3:
                        # 3 dimension
                        data     = self.data.reshape((int(self.dim[2]), int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[2])

            except IOError:
                print('open "%s" failed' %dollar_path)

        elif self.format == 'S':
            # txt format
            try:
                with open(dollar_path, 'r') as f:

                    if self.type == '4':
                        # float 32
                        self.data = np.loadtxt(f, np.float32)

                    elif self.type == '8':
                        # float 64
                        self.data = np.loadtxt(f, np.float64)

                    if len(self.dim) == 2:
                        # 2 dimension
                        data     = self.data.reshape((int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[1])
                        # print(self.data_len)

                    elif len(self.dim) == 3:
                        # 3 dimension
                        data     = self.data.reshape((int(self.dim[2]), int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[2])
            except IOError:
                print('open "%s" failed' % dollar_path)
        # print(self.data.shape)
        return self.data, data

    def read_unit(self):

        def rewrite_unit(unit):
            '''
            rewrite the unit, e.g. MLL->ML^2
            :param unit:
            :return:
            '''
            all_letter = [i for i in unit]
            letter_list = set(all_letter)
            letter_num = {}
            new_unit = ''

            for i in letter_list:
                letter_num[i] = unit.count(i)

            for i in unit:

                if i not in new_unit:
                    if letter_num[i] == 1:
                        new_unit += i
                    else:
                        new_unit += '%s^%s' % (i, letter_num[i])
            return new_unit

        def unit_transfer(unit):
            '''
            transfer unit for FL to Nm
            :param unit:
            :return:
            '''
            new_unit = ''

            if '-' in unit or 'N' in unit:
                new_unit = '-'
            elif '/' not in unit:
                new_unit = rewrite_unit(unit)
            elif '/' in unit:
                # in case of unit like FL/L
                pre_letter = unit.split('/')[0]
                las_letter = unit.split('/')[1]
                new_unit = '/'.join([rewrite_unit(pre_letter), rewrite_unit(las_letter)])

            # print(type(new_unit))
            new_unit = new_unit.replace("F", "N")
            new_unit = new_unit.replace('L', 'm')
            new_unit = new_unit.replace('A', 'rad')
            new_unit = new_unit.replace('M', 'Kg')
            new_unit = new_unit.replace('T', 's')
            new_unit = new_unit.replace('P', 'W')
            new_unit = new_unit.replace('Q', 'VA')
            new_unit = new_unit.replace('I', 'A')
            new_unit = new_unit.replace('V', 'V')
            # print(new_unit)
            return new_unit

        if self.var_list and self.ori_units:

            for ind, unit in enumerate(self.ori_units):

                self.var_unit[self.var_list[ind]] = unit_transfer(unit)
        # print(self.var_unit.items())

if __name__ == '__main__':

    pj_name = '12_aa-01'
    pj_path = r'\\172.20.0.4\fs02\CE\V3\loop05\run_0615\DLC12\12_aa-01'
    ext     = '08'
    res = read_bladed(pj_name, pj_path, ext)
    res.read_percent()
    print(res.var_list)
    # res = read_bladed(pj_name, pj_path, ext).read_dollar()
    # print(res.shape)
    # print(res[:][1])