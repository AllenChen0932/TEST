# -*- coding: utf-8 -*-
# @Time    : 2019/12/3 15:49
# @Author  : CJG
# @File    : bladed_rainflow.py
"""
修复了shortname()取叶片、塔架高度的bug;
修复了tower变量名字不统一的bug；
新增功能：提取三叶片中的疲劳最大值；
"""
import os
import numpy as np

class BladedRainflow(object):
    """读取一个loop对应的疲劳数据"""
    def __init__(self, rainflow_path, main=True):

        self.rainflow = rainflow_path
        self.main_all = main

        self.chan_pj      = {}       # {通道：pj名称}
        self.ch_path      = {}       # {channel: channel path}
        self.tower_height = {}       # {channel: tower height}

        self.channel     = []      # rainflow路径下的各输出通道，即blade, hub,...
        self.main_chan   = []      # main channels, including tower
        self.chan_tower  = []      # tower bottom/tower top channel
        self.tower_del   = {}      # tower channel: my del
        self.dlc_path    = {}      # {工况名：路径}
        self.chan_var    = {}      # {通道：[力变量列表]}
        self.fatigue_val = {}      # {variable name: {m: {dlc: value}}}

        self.get_channel()
        # self.sort_channel()
        # self.get_pj()
        self.get_result()
        self.get_tower_channel()
        self.get_main_channel()
        self.get_max_blade()
        # self.change_tower_var()
        # self.get_tower_del()

    # get channel from pj
    def get_channel(self):

        for root, dirs, files in os.walk(self.rainflow, False):
            pj_num = 0

            for file in files:

                if '.$PJ' in file.upper():

                    pj_num += 1
                    # get pj from file path
                    pj_path = os.path.join(root, file)
                    pj_name = file.split('.$')[0]

                    with open(pj_path, 'r') as f:

                        lines = f.readlines()
                        for line in lines:

                            if 'DESCRIPTION' in line:
                                # get channel name through variable
                                if 'Tower' in line:
                                    # tower channel
                                    if 'height' not in line:
                                        # onshore
                                        # Tower Mbr 10 End 1
                                        # Tower Location=Mbr 13 End 1
                                        var_list = ['Tower', line.split('"')[-2].split(',')[-1]]
                                        chan_name = ''.join(var_list)

                                    else:
                                        # offshore
                                        # Tower station height= 0m
                                        chan_name = line.split('"')[-2].split(',')[-1].strip()

                                # Blade root axes 1/2/3
                                elif 'Root axes' in line:
                                    chan_name = 'Blade root axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes 1/2/3
                                elif 'User axes' in line:
                                    chan_name = 'Blade user axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes 1/2/3
                                elif 'Principal axes' in line:
                                    chan_name = 'Blade Principal axes ' + line.split('"')[-2].split()[1]

                                # Blade user axes 1/2/3
                                elif 'Aerodynamic axes' in line:
                                    chan_name = 'Blade aerodynamic axes ' + line.split('"')[-2].split()[1]

                                elif 'Blade root' in line:
                                    # Blade root 1/2/3
                                    var_name = line.split('"')[-2].split()
                                    chan_name = ' '.join(var_name[:3])

                                else:
                                    # Stationary hub
                                    # Rotating hub
                                    var_name = line.split('"')[-2].split()
                                    chan_name = ' '.join(var_name[:2])

                                if chan_name:
                                    # print(chan_name)
                                    self.ch_path[chan_name] = root
                                    self.chan_pj[chan_name] = pj_name
                                    self.channel.append(chan_name)

                                break

            if pj_num > 1:
                raise Exception('%s exists %s projects!' %(root, pj_num))
        print(self.channel)

    # check pj file under each channel post directory
    def get_pj(self):

        for key, val in self.ch_path.items():

            files = os.listdir(val)
            # print(files)
            for file in files:

                pj_num = 0
                if '.$PJ' in file.upper():
                    # channel: pj name
                    self.chan_pj[key] = file.split('.$')[0]

                    pj_num += 1
                    if pj_num > 1:
                        print('Warning: more than 1 .$PJ file found in ' + val)

         # print(self.chan_pj)

    # get rainflow result
    def get_result(self):

        percent = []

        for chan, chan_path in self.ch_path.items():

            pj_name = self.chan_pj[chan]

            percent = read_percent(chan_path, pj_name, '026')

            fatigue = read_dollar(chan_path, pj_name, '026',
                                  percent[0], percent[1], percent[2], percent[3], percent[5], percent[6])
            # print(fatigue[0])

            self.fatigue_val.update(fatigue[0])

            self.chan_var[chan] = percent[3]

        # it is assumed that all channels share the same dlc_path and dlc_list (run path):
        dlc_path = percent[4]
        dlc_list = percent[5]

        self.dlc_path = dict(zip(dlc_list, dlc_path))

        # print(self.fatigue_val.keys())

    def get_tower_channel(self):
        '''
        get my del for each tower channel
        '''
        for chan in self.channel:

            if 'Tower' in chan and 'Mbr' in chan:
                # offshore
                mbr_num = chan.split()[-3]

                variable = [key for key in self.fatigue_val.keys() if 'Tower My' in key and mbr_num in key][0]
                m_list = [float(key) for key in self.fatigue_val[variable].keys()]
                m_list.sort()

                my_del = self.fatigue_val[variable][str(m_list[0])]['Total']
                self.tower_del[chan] = my_del
                # print(chan, variable, my_del)

            elif 'Tower' in chan and 'height' in chan:
                # onshore
                height = chan.split()[-1][:-1]

                variable = [key for key in self.fatigue_val.keys() if ('Tower My' in key) and (height in key)][0]
                m_list = [float(key) for key in self.fatigue_val[variable].keys()]
                m_list.sort()

                my_del = self.fatigue_val[variable][str(m_list[0])]['Total']
                self.tower_del[chan] = my_del
                # print(chan, variable, my_del)

    def get_main_channel(self):

        for chan in self.channel:
            if 'Tower' not in chan:
                self.main_chan.append(chan)
        self.main_chan.sort()

        # get main channel and tower channel
        my_del = [val for key, val in self.tower_del.items()]
        # my_del.sort()

        min_del   = min(my_del)
        tower_top = [key for key, val in self.tower_del.items() if val == min_del][0]
        self.main_chan.append(tower_top)
        self.chan_tower.append(tower_top)

        max_del   = max(my_del)
        tower_bot = [key for key, val in self.tower_del.items() if val == max_del][0]
        self.main_chan.append(tower_bot)
        self.chan_tower.append(tower_bot)

    def get_max_blade(self):
        """
        新建一个channel，存取3叶片中的最大疲劳(取m=4,10)，并删除原channel
        需要更新以下变量：self.channel, self.fatigue_val, self.chan_var
        """

        if self.main_all:
            all_var = ' '.join(list(self.fatigue_val.keys()))

            # Blade root:
            if 'BR1' in all_var and 'BR2' in all_var and 'BR3' in all_var:
                max_var  = [0 for i in range(6)]

                # 去除原channel, chan_var中的Blade root相关量：
                # 注意：self.fatigue_val应保持原样，供后续索引调用（写入excel时根据channel，chan_var进行索引）
                chan_del = []
                for chan, var_list in self.chan_var.items():
                    if 'BR1' in ''.join(var_list) or 'BR2' in ''.join(var_list) or 'BR3' in ''.join(var_list):
                        chan_del.append(chan)
                # print(chan_del)
                for chan in chan_del:
                    del self.chan_var[chan]
                    self.main_chan.remove(chan)

                # 查找最大值：
                tmp_MF1 = [0 for i in range(6)]

                for var in self.fatigue_val.keys():
                    if 'BR' in var and 'Mx' in var and 'BRS' not in var:
                        Mx2 = self.fatigue_val[var]['4.0']['Total']

                        if Mx2 >= tmp_MF1[0]:
                            max_var[0] = var
                            tmp_MF1[0] = Mx2

                    if 'BR' in var and 'My' in var and 'BRS' not in var:
                        My2 = self.fatigue_val[var]['4.0']['Total']

                        if My2 >= tmp_MF1[1]:
                            max_var[1] = var
                            tmp_MF1[1] = My2

                    if 'BR' in var and 'Mz' in var and 'BRS' not in var:
                        Mz2 = self.fatigue_val[var]['4.0']['Total']

                        if Mz2 >= tmp_MF1[2]:
                            max_var[2] = var
                            tmp_MF1[2] = Mz2

                    if 'BR' in var and 'Fx' in var and 'BRS' not in var:
                        Fx2 = self.fatigue_val[var]['4.0']['Total']

                        if Fx2 >= tmp_MF1[3]:
                            max_var[3] = var
                            tmp_MF1[3] = Fx2

                    if 'BR' in var and 'Fy' in var and 'BRS' not in var:
                        Fy2 = self.fatigue_val[var]['4.0']['Total']

                        if Fy2 >= tmp_MF1[4]:
                            max_var[4] = var
                            tmp_MF1[4] = Fy2

                    if 'BR' in var and 'Fz' in var and 'BRS' not in var:
                        Fz2 = self.fatigue_val[var]['4.0']['Total']

                        if Fz2 >= tmp_MF1[5]:
                            max_var[5] = var
                            tmp_MF1[5] = Fz2

                # 添加新channel, variable
                self.main_chan.insert(0, 'Blade root')
                self.chan_var.update({'Blade root': max_var})
                # print(self.main_chan)
                # print([[var,self.fatigue_val[var]['4.0']['Total']] for var in self.chan_var['Blade root']])

            elif 'BR1' in all_var or 'BR2' in all_var or 'BR3' in all_var:

                blade_root = [[key,val] for key, val in enumerate(self.main_chan) if 'Blade root' in val][0]
                del self.main_chan[blade_root[0]]
                self.main_chan.insert(0, 'Blade root')

                self.chan_var['Blade root'] = self.chan_var.pop(blade_root[1])

            # Blade Root axes:
            if 'BRS1' in all_var and 'BRS2' in all_var and 'BRS3' in all_var:

                max_var = [0 for i in range(6)]

                # 去除原channel, chan_var中的Blade root相关量：
                # 注意：self.fatigue_val应保持原样，供后续索引调用（写入excel时根据channel，chan_var进行索引）
                chan_del = []

                for chan, var_list in self.chan_var.items():
                    all_var = ' '.join(var_list)
                    # print(all_var)
                    if 'BRS1' in all_var or 'BRS2' in all_var or 'BRS3' in all_var:
                        chan_del.append(chan)
                # print('del', chan_del)

                for chan in chan_del:
                    del self.chan_var[chan]
                    self.main_chan.remove(chan)

                # 查找最大值：
                tmp_MF1 = [0 for i in range(6)]

                for var in self.fatigue_val.keys():

                    if 'BRS' in var and 'Mx' in var:
                        Mx2 = self.fatigue_val[var]['10.0']['Total']

                        if Mx2 >= tmp_MF1[0]:
                            max_var[0] = var
                            tmp_MF1[0] = Mx2

                    if 'BRS' in var and 'My' in var:
                        My2 = self.fatigue_val[var]['10.0']['Total']

                        if My2 >= tmp_MF1[1]:
                            max_var[1] = var
                            tmp_MF1[1] = My2

                    if 'BRS' in var and 'Mz' in var:
                        Mz2 = self.fatigue_val[var]['10.0']['Total']

                        if Mz2 >= tmp_MF1[2]:
                            max_var[2] = var
                            tmp_MF1[2] = Mz2

                    if 'BRS' in var and 'Fx' in var:
                        Fx2 = self.fatigue_val[var]['10.0']['Total']

                        if Fx2 >= tmp_MF1[3]:
                            max_var[3] = var
                            tmp_MF1[3] = Fx2

                    if 'BRS' in var and 'Fy' in var:
                        Fy2 = self.fatigue_val[var]['10.0']['Total']

                        if Fy2 >= tmp_MF1[4]:
                            max_var[4] = var
                            tmp_MF1[4] = Fy2

                    if 'BRS' in var and 'Fz' in var:
                        Fz2 = self.fatigue_val[var]['10.0']['Total']

                        if Fz2 >= tmp_MF1[5]:
                            max_var[5] = var
                            tmp_MF1[5] = Fz2

                # 添加新channel, variable
                self.main_chan.insert(1, 'Blade root axes')
                self.chan_var.update({'Blade root axes': max_var})
                # print(self.main_chan)
                # print(self.chan_var['Blade root axes'])

            elif 'BRS1' in all_var or 'BRS2' in all_var or 'BRS3' in all_var:

                blade_root = [[key, val] for key, val in enumerate(self.main_chan) if 'axes' in val][0]
                del self.main_chan[blade_root[0]]
                self.main_chan.insert(0, 'Blade root axes')

                self.chan_var['Blade root axes'] = self.chan_var.pop(blade_root[1])

    def change_tower_var(self):

        # tower top
        tower_top = self.chan_tower[0]
        top_list  = self.chan_var[tower_top]   # original tower top variable list
        var_list  = []                         # new tower top variable list
        for var in top_list:
            var_list.append(var.split(',')[0].split()[-1]+'TT')
        self.chan_var[tower_top] = var_list

        if 'height' in tower_top:
            # onshore
            for var in top_list:
                var_name = var.split(',')[0].split()[-1]+'TT'
                # print(var_name)
                self.fatigue_val[var_name] = self.fatigue_val.pop(var)

        elif 'Mbr' in tower_top:
            # offshore
            for var in top_list:
                var_name = var.split(',')[0].split()[-1] +'TT'
                # print(var_name)
                self.fatigue_val[var_name] = self.fatigue_val.pop(var)

        # tower bottom
        tower_bot = self.chan_tower[1]
        bot_list  = self.chan_var[tower_bot] # original tower bottom variable list
        # print(bot_list)
        var_list  = []                       # new tower bottom variable list
        for var in bot_list:
            var_list.append(var.split(',')[0].split()[-1] +'TB')
        self.chan_var[tower_bot] = var_list

        if 'height' in tower_bot:
            # onshore
            for val in  bot_list:
                var_name = val.split(',')[0].split()[-1] +'TB'
                self.fatigue_val[var_name] = self.fatigue_val.pop(val)

        elif 'Mbr' in tower_bot:
            # offshore
            for val in  bot_list:
                var_name = val.split(',')[0].split()[-1] +'TB'
                # print(var_name)
                self.fatigue_val[var_name] = self.fatigue_val.pop(val)

        # print(self.chan_var.items())
        print(self.fatigue_val.keys())

# ------------------------------------sub functions-------------------------------------
# read DELs on a specified component
def read_percent(path, name, ext):

    access    = None
    recl      = None
    dimension = None
    load_name = None
    dlc_file  = None
    dlc_path  = []
    dlc_list  = []
    m_SN      = None

    file = ''.join([name, '.%', ext])
    file_path = os.sep.join([path, file])
    # print(file_path)

    if os.path.isfile(file_path):
        with open(file_path, 'r') as f:
            flag = 0

            for line in f.readlines():

                if 'ACCESS' in line:
                    access = line.split()[1]
                    # print(access)

                if 'RECL' in line:
                    recl = line.split()[1]
                    # print(line.split()[-1])
                    # print(recl)

                if line.startswith('DIMENS'):
                    dimension = line.split()[1:]
                    # print(dimension)

                if 'VARIAB' in line:
                    load_name = line.split("'")[1::2]
                    load_name = shortname(load_name)
                    # print(load_name)

                if 'AXITICK' in line:   # the last one is 'Total'
                    dlc_file = line.split()[1:]
                    flag = 1

                if flag == 1 and 'Inverse SN Slope' not in line and 'AXITICK'not in line:
                    dlc_file.extend(line.split())

                if 'Inverse SN Slope' in line:
                    flag = 0

                if 'AXIVAL' in line:
                    m_SN = line.split()[1:]
                    # print(m_SN)

            # 记录时序工况路径、工况名：
            # if ext == '026':
            for i in range(len(dlc_file)-1):

                dlc_path.append(os.sep.join(dlc_file[i].split(os.sep)[:-1]))

                dlc_list.append(dlc_file[i].split(os.sep)[-1])

            dlc_path.append(dlc_file[-1])  # 'Total'

            dlc_list.append(dlc_file[-1])  # 'Total'
            # print(dlc_list)

            # 若run文件夹路径中出现空格，程序不能使用：
            if len(dlc_path) != int(dimension[1]):

                print('\nPlease do not include blank in run path:' + '\n' + dlc_path[0] + '\n')

        return access, recl, dimension, load_name, dlc_path, dlc_list, m_SN

    else:

        raise Exception('%s not exist!' %file_path)

def read_dollar(path, name, ext, access, recl, dimension, load_name, dlc_list, m_SN):

    # 数据存放方式：
    # fatigue{ load_name1:{m1:{dlc1:value1, ..., total:value}}, {m2:{...}, ...}, load_name2:{...} ... }
    # dlc{ dlc_name1:dlc_path, dlc_name2:dlc_path, ... }

    fatigue = {}
    data = np.array

    # 读取$026结果文件，数据写入data：
    file = ''.join([name, '.$', ext])
    file_dir = os.sep.join([path, file])

    if os.path.isfile(file_dir):

        if access == 'D':

            with open(file_dir, 'rb') as f:

                if recl == '4':

                    data = np.fromfile(f, np.float32)

                elif recl == '8':

                    data = np.fromfile(f, np.float64)

        elif access == 'S':

            with open(file_dir, 'r') as f:

                data = np.loadtxt(f)

        # print(data.shape)
        # print(data[:61])

        data = data.reshape(int(dimension[2]), int(dimension[1]), int(dimension[0]))
        # print(data.shape)
        # print(data[0, :, 0])

        # 先按不同载荷变量，再按不同m值，接着按不同工况，将DEL写入fatigue嵌套字典中：
        for load in load_name:

            m_dict = {}

            for m in m_SN:

                dlc_dict = {}

                for dlc in dlc_list:

                    dlc_dict.update({dlc: data[m_SN.index(m), dlc_list.index(dlc), load_name.index(load)]})

                m_dict.update({m: dlc_dict})

                fatigue.update({load: m_dict})

        # print(fatigue['Blade root 1 Mx']['4.0'])

        return fatigue, data

    else:

        raise Exception('%s not exist!' % file_dir)

def shortname(load_name):

    load_name_new = []

    for load in load_name:

        if 'Root axes' in load:
            temp = load.split(',')[0]
            load = ''.join([temp.split()[2],'BRS',temp.split()[1]])
            # print(load)

        elif 'User axes' in load:
            temp = load.split(',')[0]
            load = ''.join([temp.split()[2],'BUS',temp.split()[1]])

        elif 'Principal axes' in load:
            temp = load.split(',')[0]
            load = ''.join([temp.split()[2],'BPS',temp.split()[1]])

        elif 'Aerodynamic axes' in load:
            temp = load.split(',')[0]
            load = ''.join([temp.split()[2],'BAS',temp.split()[1]])

        elif 'Tower' in load:
            temp = load.split(',')

            if 'height' in temp[1]:
                height = temp[1].split('=')[1].strip().split('m')[0]
                load = temp[0] + ',' + '%s' % height

            elif 'Location' in temp[1]:
                height = temp[1].split('=')[1]
                load = temp[0] + ',' + height

            else:
                load = temp[0] + ',' + temp[1].strip()

        elif 'Rotating' in load:
            load = ''.join([load.split()[-1], 'HR'])

        elif 'Stationary' in load:
            load = ''.join([load.split()[-1], 'HS'])

        elif 'Yaw' in load:
            load = ''.join([load.split()[-1], 'YB'])

        elif 'Blade root' in load:
            load = ''.join([load.split()[-1], 'BR', load.split()[-2]])
            # print(load)

        load_name_new.append(load)
        # print(load)

    return load_name_new

if __name__ == '__main__':
    '''
    path = r'E:\00_Tools_Dev\02_LoadSummaryTable\Test_Data\Onshore_all\rainflow\Tower_station_0'
    pj_name = 'ts_0'

    percent = read_percent026(path, pj_name)
    result  = read_dollar026(path,
                             pj_name,
                             percent[0],
                             percent[1],
                             percent[2],
                             percent[3],
                             percent[5],
                             percent[6])
    print(result.values())
    '''
    path = r'E:\00_Tools_Dev\02_LoadSummaryTable\Test_Data\Onshore_all\rainflow'
    res = BladedRainflow(path, True)
    # print(res.channel)

