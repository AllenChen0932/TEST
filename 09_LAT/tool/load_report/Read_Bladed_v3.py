# -*- coding: utf-8 -*-
# @Time    : 2020/7/17 11:35
# @Author  : CE,CJG
# @File    : Read_Bladed_v3.py

# v3: bug fixed in data reshape when dim=3.
#     bug fixed in getting var_list.
#     add reading GENLAB,AXISLAB and axis info. in percent file.
#     call read_percent(), read_dollar() in init()

import os
import re
import numpy as np


class read_bladed(object):

    def __init__(self, pj_name, pj_path, ext):
        '''
        read a single bladed percent-dollar file,
        output variable list and data in narray format.
        :param name: pj name
        :param path: pj path
        :param ext: file extension ID
        '''

        self.pj_name = pj_name
        self.pj_path = pj_path
        self.ext = str(ext)
        self.format = ''
        self.type = ''
        self.dim = []
        self.channel = ''
        self.var_list = []
        self.axis = []
        self.ori_units = []
        self.var_unit = {}
        self.data = None
        self.data_len = None

        self.read_percent()
        self.read_dollar()
        # self.read_unit()  # called in read_percent()

    def read_percent(self):

        file = ''.join([self.pj_name, '.%', self.ext])
        path = os.sep.join([self.pj_path, file])
        # print(path)

        try:
            with open(path, 'r') as f:

                indx = []
                lines = f.readlines()
                num = len(lines)

                for i in range(len(lines)):
                    line = lines[i]
                    if line.startswith('ACCESS'):
                        self.format = line.strip("\n\t").split()[-1]

                    if line.startswith('RECL'):
                        self.type = line.strip().split()[-1]

                    if line.startswith('DIMENS'):
                        self.dim = line.strip().split()[1:]

                    if line.startswith('GENLAB'):
                        self.channel = line.strip().split("'")[1]

                    if line.startswith('VARIAB'):
                        self.var_list = line.strip().split("'")[1::2]

                    if 'VARUNIT' in line:
                        self.ori_units = line.strip().split()[1:]

                    if line.startswith('AXISLAB'):
                        indx.append(i)

                    if line.startswith('NVARS'):
                        num = i
                        break

                # read axis info. of dim[1], dim[2], ...
                indx.append(min(num, len(lines)))
                for i in range(len(self.dim)-1):
                    # get axis label
                    axislines = lines[indx[i]:indx[i+1]]
                    # print(axislines)
                    self.axis.append({})
                    self.axis[i]['label'] = axislines[0].strip().split("'")[1]
                    self.axis[i]['unit']  = axislines[1].strip().split()[1]
                    aximeth = axislines[2].strip().split()[1]

                    # get axis tick (overlook time tick because it is achievable in dollar file)
                    if not self.axis[i]['label'] == 'Time':

                        if aximeth == '1':
                            dlclist = []
                            pattern = "AXITICK(.*)"
                            dlcstr = re.findall(pattern, ''.join([l for l in axislines]), re.S)[0]  # re.S allows "." accounts for \n
                            dlclistori = dlcstr.strip().split()
                            for dlc in dlclistori:
                                dlclist.append(dlc.split('\\')[-1])
                            self.axis[i]['tick'] = dlclist

                        if aximeth == '2':
                            start = float(axislines[3].strip().split()[1])
                            step  = float(axislines[4].strip().split()[1])
                            end   = start+step*int(self.dim[i+1])
                            self.axis[i]['tick'] = np.arange(start, end, step)
                            # print(self.axis[i]['tick'].shape)

                        elif aximeth == '3':
                            axisval = axislines[3].strip().split()[1:]
                            self.axis[i]['tick'] = np.array(axisval)
                            # print(self.axis[i]['tick'].shape)
                # print(self.axis)

            self.read_unit()
            return self.var_list, self.var_unit

        except IOError:

            print('Open "%s" failed!' %path)

    def read_dollar(self):
        # read percent to get format info
        # self.read_percent()

        dollar_path = os.path.join(self.pj_path, ''.join([self.pj_name,'.$',self.ext]))

        # read dollar file to get stored data in original format
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
                        self.data     = self.data.reshape((int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[1])
                    elif len(self.dim) == 3:
                        # 3 dimension
                        self.data     = self.data.reshape((int(self.dim[2]), int(self.dim[1]), int(self.dim[0])))
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
                        self.data     = self.data.reshape((int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[1])
                        # print(self.data_len)
                    elif len(self.dim) == 3:
                        # 3 dimension
                        self.data     = self.data.reshape((int(self.dim[2]), int(self.dim[1]), int(self.dim[0])))
                        self.data_len = int(self.dim[2])

            except IOError:
                print('open "%s" failed' % dollar_path)
        # print(self.data.shape)
        return self.data

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

    # pj_name = 'hs'
    # pj_path = r'\\172.20.0.4\fs02\LSY\V3\loop04_backup20200623\BackUp\post\rainflow\main_25y_1E7\Hub_stationary'
    # ext     = '008'
    # pj_name = 'br'
    # pj_path = r'\\172.20.0.4\fs02\LSY\V3\loop04_backup20200623\BackUp\post\rainflow\main_25y_1E7\Blade_root'
    # ext     = '001'
    # pj_name = '12_ac-02'
    # pj_path = r'\\172.20.0.4\fs03\CJG\V2B\loop2.2\run\DLC12\12_ac-02'
    # ext = '58'
    pj_name = 'bstats'
    pj_path = r'\\172.20.0.4\fs01\LSY\W4800-146-90_S\W4800-146-90_loop3.2\BackUp\post\basic_statistics\generator_speed'
    ext = '002'
    # res = read_bladed(pj_name,pj_path,ext).read_dollar()
    # print(res.shape)
    # print(res[:][1])
    res = read_bladed(pj_name, pj_path, ext)#.read_percent()
    # print(res.var_list)
    # print(res.data)
