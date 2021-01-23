# -*- coding: utf-8 -*-
# @Time    : 2019/11/24 11:31
# @Author  : CE
# @File    : bladed_ultimate.py

import os
import numpy as np
import operator

__version__ = "2.0.2"
'''
2020.4.21_v2.0.2: modify the variable output sequence, from M to F
'''
# read % file
def read_percent(path, name, ext):
    '''
    读取channel文件夹下的百分号文件
    :param path: channel路径
    :param name: PJ名称
    :param ext: extension list
    :return:
    '''

    access    = None
    recl      = None
    variab    = None
    dimension = None

    axis_index = 0
    dlc_list   = ''

    file      = ''.join([name,'.%',ext])
    file_path = os.sep.join([path, file])
    # print(file_path)

    if os.path.isfile(file_path):

        with open(file_path, 'r') as f:

            lines = f.readlines()

            for line in lines:

                if 'ACCESS' in line:
                    access = line.split()[1]
                    # print(access)

                if 'RECL' in line:
                    recl = line.split()[1]
                    # print(recl)

                if 'VARIAB' in line:
                    variab = line.split("'")[1]
                    var    = [v for v in variab.split() if 'F' in v or 'M' in v]

                    if 'Root axes' in variab:
                        # root axex
                        variab = var[0]+'BRS'

                    elif 'User axes' in variab:
                        # user  axes
                        variab = var[0]+'BUS'

                    elif 'Principal axes' in variab:
                        # user  axes
                        variab = var[0]+'BPS'

                    elif 'Aerodynamic axes' in variab:
                        # user  axes
                        variab = var[0]+'BAS'

                    elif 'Blade root' in variab:
                        # blade root
                        variab = var[0]+'BR'

                    elif 'Stationary' in variab:
                        # hub stationary
                        variab = var[0]+'HS'

                    elif 'Rotating' in variab:
                        # rotating hub
                        variab = var[0]+'HR'

                    elif 'Yaw bearing' in variab:
                        # yaw bearing
                        variab = var[0]+'YB'

                    elif ',' in variab and 'Tower' in variab:
                        # tower
                        temp = variab.split(',')
                        if 'height' in temp[1]:
                            # onshore
                            height = temp[1].split('=')[1].strip().split('m')[0]
                            variab = temp[0] + ',' + '%.2f' % float(height)

                        elif 'Location' in temp[1]:
                            # offshore
                            height = temp[1].split('=')[1]
                            variab = temp[0] + ',' + height

                        else:
                            # offshore
                            variab = temp[0] + ',' + temp[1].strip()

                        if variab[-1] == '.':
                            variab = variab[:-1]

                if line.startswith('DIMENS'):
                    # DIMENS
                    dimension = line.split()[1:]
                    # print(dimension)

                if 'AXITICK' in line:
                    # AXITICK
                    axis_index = lines.index(line)

            for i in range(axis_index, len(lines)):
                dlc_list += lines[i]

            axitick = dlc_list.split()[1:]

        return access, recl, dimension, variab, axitick

    else:
        raise Exception('%s not exist!' %file_path)
# read $ file
def read_dollar(path, name, ext, access, recl, dimension, axitick):

    dlc_value = {}
    ult_value = {}

    # data = np.array
    file     = ''.join([name, '.$', ext])
    file_dir = os.sep.join([path, file])


    if os.path.isfile(file_dir):

        if access == 'D':

            with open(file_dir, 'rb') as f:

                if recl == '4':

                    data = np.fromfile(f, np.float32)
                    # print(data)

                    data = data.reshape(int(dimension[1]), int(dimension[0]))
                    # print(data)

                    ult_value[axitick[0]] = data[0, :]

                    for var in axitick:

                        axitick.index(var)

                        dlc_value[var] = data[axitick.index(var), :]

                elif recl == '8':

                    data = np.fromfile(f, np.float64)

                    data = data.reshape(int(dimension[1]), int(dimension[0]))

                    ult_value[axitick[0]] = data[0, :]

                    for var in axitick:

                        axitick.index(var)

                        dlc_value[var] = data[axitick.index(var), :]

        elif access == 'S':

            with open(file_dir, 'r') as f:

                data = np.loadtxt(f)

                data = data.reshape(int(dimension[1]), int(dimension[0]))

                ult_value[axitick[0]] = data[0, int(dimension[0])]

                for var in axitick:

                    axitick.index(var)

                    dlc_value[var] = data[axitick.index(var), int(dimension[0])]

        return dlc_value, ult_value

    else:

        raise Exception('%s not exist!' %file_dir)

class BladedUltimate(object):

    def __init__(self, ultimate_path, main=True):

        self.ultimate = ultimate_path
        self.main_all = main        #True: main, False: all

        self.ch_path  = {}          #{channel：channel路径}
        self.chan_pj  = {}          #{channel：pj名称}
        # self.chan_ext = {}          #{channel: extensions}
        self.chan_tow = []          #{channel-tower top, channel-tower bottom}
        self.tow_mres = {}          #{channel: Mres}
        self.chan_ext = ['001', '002', '003', '004', '005', '006', '007', '008']

        # output
        self.channels = []          #输出通道
        self.main_cha = []          #主要输出通道
        self.dlc_path = {}          #工况名：路径
        self.var_ult  = {}          #{var: {dlc:val}, {变量：{工况（所有）：最大值}}
        self.ult_val  = {}          #{var: {dlc:val}, {变量：{工况（单个）：极值}
        self.chan_var = {}          #{chan:var}
        self.dlcs_DLC = {}          #{12_aa-01: DLC12}
        self.chan_br  = None
        self.br_mx    = {}


        self.get_channel()
        # self.get_extension()
        self.get_dlc_path()
        self.get_result()
        self.read_BRMxMy()
        if self.tow_mres:
            self.sort_channel()
            # self.change_var()

    # 读取ultimate路径下的输出通道名称及PJ名称
    def get_channel(self):
        '''
        Read all channels from the PJ file in ultimate result path
        Channel names are based on the variables, not on the file names,
        such as Blade root, Blade root axes, Hub rotating, Hub stationary, Tower station height=0m(or Tower Mbr)
        :return:
        '''

        for root, dirs, files in os.walk(self.ultimate, False):
            pj_num = 0

            for file in files:
                if '.$PJ' in file.upper():
                    pj_num += 1
                    pj_path = os.path.join(root, file)
                    pj_name = file.split('.$')[0]

                    with open(pj_path, 'r') as f:
                        lines = f.readlines()

                        for line in lines:
                            # channel name according to variable name
                            if 'DESCRIPTION' in line:
                                # tower channel
                                if 'Tower' in line:
                                    # Tower Mbr 10 End 1
                                    # Tower Location=Mbr 13 End 1
                                    if 'height' not in line:
                                        var_list  = ['Tower', line.split('"')[-2].split(',')[-1]]
                                        chan_name = ''.join(var_list)
                                    # Tower station height= 0m
                                    else:
                                        chan_name = line.split('"')[-2].split(',')[-1].strip()
                                    # print(chan_name)
                                # Blade root axes
                                elif 'Root axes' in line:
                                    chan_name = 'Blade root axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes
                                elif 'User axes' in line:
                                    chan_name = 'Blade user axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes
                                elif 'Principal axes' in line:
                                    chan_name = 'Blade principal axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes
                                elif 'Aerodynamic axes' in line:
                                    chan_name = 'Blade aerodynamic axes ' + line.split('"')[-2].split()[1]

                                # Blade root
                                # Stationary hub
                                # Rotating hub
                                # Yaw bearing
                                else:
                                    var_name  = line.split('"')[-2].split()
                                    chan_name = ' '.join(var_name[:2])
                                    if 'Blade root'  in chan_name:
                                        self.chan_br = chan_name

                                if chan_name:
                                    self.ch_path[chan_name] = root
                                    self.chan_pj[chan_name] = pj_name
                                    self.channels.append(chan_name)
                                break

            if pj_num > 1:
                raise Exception('%s exists %s projects!' %(root, pj_num))

        print('All channels:', self.channels)

    def read_BRMxMy(self):

        if self.chan_br:
            file_path = os.path.join(self.ch_path[self.chan_br], self.chan_pj[self.chan_br]+'.$MX')

            if not os.path.isfile(file_path):
                raise Exception('%s is not a file!' %file_path)

            with open(file_path, 'r') as f:
                lines = f.readlines()[2:6]

                mxbr_max = lines[0].strip().split('\t')
                # print(mxbr_max)
                self.br_mx['MxBR_Max'] = [mxbr_max[2].strip(), float(mxbr_max[3].strip())]
                mxbr_min = lines[1].strip().split('\t')
                self.br_mx['MxBR_Min'] = [mxbr_min[2].strip(), float(mxbr_min[3].strip())]

                mybr_max = lines[2].strip().split('\t')
                self.br_mx['MyBR_Max'] = [mybr_max[2].strip(), float(mybr_max[4].strip())]
                mybr_min = lines[3].strip().split('\t')
                self.br_mx['MyBR_Min'] = [mybr_min[2].strip(), float(mybr_min[4].strip())]

    def get_dlc_path(self):
        '''get dlc name and directory defined in PJ file'''

        for chan in self.channels:
            pj_path = os.sep.join([self.ch_path[chan], self.chan_pj[chan]+'.$PJ'])

            # No need to check file path and raise error
            with open(pj_path, 'r') as f:
                for line in f.readlines():

                    if 'DIRECTORY' in line:
                        dlc_dir  = line.split()[1]
                        dlc_name = dlc_dir.split('\\')[-1]
                        dlc      = dlc_dir.split('\\')[-2]

                        self.dlc_path[dlc_name] = dlc_dir
                        self.dlcs_DLC[dlc_name] = dlc
        # print(self.dlc_path.keys())

    def get_result(self):
        '''get result for each channel'''

        for chan in self.channels:
            var_name  = []    #record variables in each channel
            chan_path = self.ch_path[chan]

            for ext in self.chan_ext:
                percent = read_percent(chan_path, self.chan_pj[chan], ext)
                var_name.append(percent[3])

                dlc_val = read_dollar(chan_path, self.chan_pj[chan], ext,
                                      percent[0], percent[1], percent[2], percent[4])

                self.var_ult[percent[3]] = dlc_val[0]
                self.ult_val[percent[3]] = dlc_val[1]

                # record tower m_res
                # onshore
                if ('height' in chan) and (ext == '003'):
                    self.tow_mres[chan] = [val for key,val in dlc_val[1].items()][0]
                # offshore
                elif ('Mbr' in chan) and (ext == '004'):
                    self.tow_mres[chan] = [val for key,val in dlc_val[1].items()][0]

            self.chan_var[chan] = var_name

        # record main channel
        self.main_cha = [chan for chan in self.channels if 'Tower' not in chan]
        # sort main channel, e.g.: br, brs, hr, hs, yb
        self.main_cha.sort()

    def sort_channel(self):
        '''sort channel in sequence'''

        # sort tower channel according to Mres sequence
        self.tow_mres = sorted(self.tow_mres.items(), key=operator.itemgetter(1))

        # get tower top channel
        tower_top = self.tow_mres[0][0]
        self.main_cha.append(tower_top)
        self.chan_tow.append(tower_top)

        # get tower bottom channel
        tower_bot = self.tow_mres[-1][0]
        self.main_cha.append(tower_bot)
        self.chan_tow.append(tower_bot)
        # print(self.main_cha)

    def change_var(self):
        '''change variable in tower top and tower bottom'''

        # tower top
        tower_top    = self.chan_tow[0]
        top_var_list = self.chan_var[tower_top]
        new_var_list = []   #variables in tower
        # print(top_var_list)

        # change variables in self.chan_var
        for var in top_var_list:
            new_var_list.append(var.split(',')[0].split()[-1]+'TT')
        self.chan_var[tower_top] = new_var_list
        # print(new_var_list)

        # change variables in self.var_ult
        for key, val in self.var_ult.items():
            if key in top_var_list:
                self.var_ult[new_var_list[top_var_list.index(key)]] = self.var_ult.pop(key)

        # change variables in self.ult_val
        for key, val in self.ult_val.items():
            if key in top_var_list:
                self.ult_val[new_var_list[top_var_list.index(key)]] = self.ult_val.pop(key)
        # print(self.chan_var[tower_top])

        # tower bottom
        tower_bot    = self.chan_tow[1]
        bot_var_list = self.chan_var[tower_bot]
        new_var_list = []   #variables in tower

        # change variables in self.chan_var
        for var in bot_var_list:
            new_var_list.append(var.split(',')[0].split()[-1]+'TB')
        self.chan_var[tower_bot] = new_var_list

        # change variables in self.ult_val
        for key, val in self.var_ult.items():
            if key in bot_var_list:
                self.var_ult[new_var_list[bot_var_list.index(key)]] = self.var_ult.pop(key)

        # change variables in self.ult_val
        for key, val in self.ult_val.items():
            if key in bot_var_list:
                self.ult_val[new_var_list[bot_var_list.index(key)]] = self.ult_val.pop(key)

if __name__ == '__main__':

    path = r'\\172.20.4.133\fs03\CE\W6250-172-18m\post\0428\ultimate'
    BladedUltimate(path, False)
