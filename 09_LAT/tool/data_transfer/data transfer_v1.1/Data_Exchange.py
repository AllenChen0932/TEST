#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/6/2020 12:32 PM
# @Author  : CE
# @File    : data_exchange_tower.py

import datetime
import multiprocessing as mp
import os
import struct
import time

import numpy as np

try:
    from tool.data_transfer.Read_Bladed_v2 import read_bladed as rb
except:
    from Read_Bladed_v2 import read_bladed as rb

def unit_exchange(var_list, var_unit):

    transfer_factor = []

    for var in var_list:

        unit = var_unit[var]

        # time
        if unit.lower() == 's':
            factor = 1.0

        # angle
        elif unit.lower() == 'rad':
            factor = 1.0
        elif unit.lower() == 'deg':
            factor = 180.0/np.pi
        elif unit.lower() == 'rev':
            factor = 1.0/(2*np.pi)

        # force
        elif unit.lower() == 'n':
            factor = 1.0
        elif unit.lower() == 'kn':
            factor = 1.0/1000
        elif unit.lower() == 'mn':
            factor = 1.0/1000000

        # moment
        elif unit.lower() == 'nm':
            factor = 1.0
        elif unit.lower() == 'knm':
            factor = 1.0/1000
        elif unit.lower() == 'mnm':
            factor = 1.0/1000000

        # speed
        elif unit.lower() == 'm/s':
            factor = 1.0
        elif unit.lower() == 'rad/s':
            factor = 1.0
        elif unit.lower() == 'deg/s':
            factor = 180.0/np.pi
        elif unit.lower() == 'rev/s':
            factor = 1.0/(2*np.pi)
        elif unit.lower() == 'rpm':
            factor = 60/(2*np.pi)

        # acceleration
        elif unit.lower() == 'm/s^2':
            factor = 1.0
        elif unit.lower() == 'deg/s^2':
            factor = 180.0/np.pi
        elif unit.lower() == 'rad/s^2':
            factor = 1.0

        # power
        elif unit.lower() == 'w':
            factor = 1.0
        elif unit.lower() == 'kw':
            factor = 1.0/1000
        elif unit.lower() == 'mw':
            factor = 1.0/1000000

        # others
        else:
            factor = 1.0

        transfer_factor.append(factor)

    return np.array(transfer_factor)

def sort_extension_index(vars_list, exts_list):

    ext_list       = []
    ext_vars_index = {}

    # element the same extension
    for ext in exts_list:
        if ext not in ext_list:
            ext_list.append(ext)

    # to get the variable list and index list for each extension
    for ext in ext_list:
        var_list = []
        ind_list = []
        for ind, var in enumerate(vars_list):
            if exts_list[ind] == ext:
                var_list.append(var)
                ind_list.append(ind)

        ext_vars_index[ext] = [var_list,ind_list]

    return ext_list, ext_vars_index

class Get_Data(object):

    def __init__(self,
                 step_num,
                 transfer_type,
                 data_format,
                 run_path,
                 res_path,
                 dlc_list,
                 unit_list,
                 index_list,
                 variable_list,
                 extension_list,
                 node_list=None,
                 section_list=None,
                 section_index=None,
                 matrix_type=None):

        self.step_num  = step_num
        self.tran_type = transfer_type
        self.data_form = data_format
        self.run_path  = run_path
        self.res_path  = os.path.join(res_path, '_'.join((transfer_type.upper(), data_format)))
        self.dlcs_list = dlc_list
        self.vars_list = variable_list
        self.unit_list = unit_list         # {var: unit}
        self.inde_list = index_list        # variable index listed in each %
        self.node_list = node_list
        self.exts_list = extension_list
        self.secs_list = section_list
        self.sec_index = section_index
        # self.b2b_flag  = bottom2top
        self.matr_type = matrix_type

        self.var_index = dict(zip(variable_list, index_list))

        self.lc_list   = []
        self.lc_dlc    = {}

        self.transfer_matrix = {}
        self.transfer_matrix['b2t_offshore'] = np.array([[ 0, 0, 0,-1, 0, 0],
                                                         [ 0, 0, 0, 0, 0, 1],
                                                         [ 0, 0, 0, 0, 1, 0],
                                                         [-1, 0, 0, 0, 0, 0],
                                                         [ 0, 0, 1, 0, 0, 0],
                                                         [ 0, 1, 0, 0, 0, 0]])

        self.transfer_matrix['t2b_offshore'] = np.array([[ 0, 0, 0,-1, 0, 0],
                                                         [ 0, 0, 0, 0, 0,-1],
                                                         [ 0, 0, 0, 0, 1, 0],
                                                         [-1, 0, 0, 0, 0, 0],
                                                         [ 0, 0,-1, 0, 0, 0],
                                                         [ 0, 1, 0, 0, 0, 0]])

        self.transfer_matrix['b2t_onshore'] = np.array([[ 0, 0, 0, 0, 0, 1],
                                                        [ 0, 0, 0, 0, 1, 0],
                                                        [ 0, 0, 0,-1, 0, 0],
                                                        [ 0, 0, 1, 0, 0, 0],
                                                        [ 0, 1, 0, 0, 0, 0],
                                                        [-1, 0, 0, 0, 0, 0]])

        self.transfer_matrix['blade'] = np.array([[ 0, 0, 0, 0, 0, 1],
                                                  [ 0, 0, 0, 0,-1, 0],
                                                  [ 0, 0, 0, 1, 0, 0],
                                                  [ 0, 0, 1, 0, 0, 0],
                                                  [ 0,-1, 0, 0, 0, 0],
                                                  [ 1, 0, 0, 0, 0, 0]])

        self.transfer_matrix['ones'] = np.identity(6)

        self.prj_path    = {}
        self.dlc_lc_list = {}

    def write_flex_sensor(self, variable_list, node_list, section_list, res_path):
        '''
        write flex sensor
        :param variable_list: variable list to output
        :param node_list:     node list for each section
        :param section_list:  height list for each section
        :param res_path:      result path
        :return:
        '''
        # reorder variable list from Fx to Mz
        # variable_list.sort()

        content  = 'Sensor list : aeroFLEX V3.2, xxxx\n'
        content += ' No   scale  offset  korr.    Volt    Unit   Name    Description------------\n'

        if 'onshore' in self.matr_type:
            for ind, node_num in enumerate(node_list):
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+1, variable_list[4], node_num, variable_list[4], float(section_list[ind]))
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+2, variable_list[5], node_num, variable_list[5], float(section_list[ind]))
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+3, variable_list[6], node_num, variable_list[6], float(section_list[ind]))
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+4, variable_list[1], node_num, variable_list[1], float(section_list[ind]))
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+5, variable_list[2], node_num, variable_list[2], float(section_list[ind]))
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,h=%.3fm\n' \
                           %(ind*6+6, variable_list[3], node_num, variable_list[3], float(section_list[ind]))
        elif 'blade' in self.matr_type:
            for ind, node_num in enumerate(node_list):
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+1, variable_list[1], node_num, variable_list[1], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+2, variable_list[2], node_num, variable_list[2], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+3, variable_list[3], node_num, variable_list[3], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+4, variable_list[4], node_num, variable_list[4], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+5, variable_list[5], node_num, variable_list[5], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%.3f\n' \
                           %(ind*6+6, variable_list[6], node_num, variable_list[6], section_list[ind])
        elif 'offshore' in self.matr_type:
            for ind, node_num in enumerate(node_list):
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%s\n' \
                           %(ind*6+1, variable_list[4], node_num, variable_list[4], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%s\n' \
                           %(ind*6+2, variable_list[5], node_num, variable_list[5], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kN)  %s,N%s  %s,S%s\n' \
                           %(ind*6+3, variable_list[6], node_num, variable_list[6], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%s\n' \
                           %(ind*6+4, variable_list[1], node_num, variable_list[1], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%s\n' \
                           %(ind*6+5, variable_list[2], node_num, variable_list[2], section_list[ind])
                content += '%s    1.0    0.0    0.00    1.0    (kNm)  %s,N%s  %s,S%s\n' \
                           %(ind*6+6, variable_list[3], node_num, variable_list[3], section_list[ind])

        if not os.path.exists(res_path):
            os.makedirs(res_path)

        sensor_path = os.path.join(res_path, 'Sensor')
        with open(sensor_path, 'w+') as f:
            f.write(content)

    def get_lc_list(self):
        '''
        get load case list for each DLC directory
        :return:
        '''
        for dlc in self.dlcs_list:
            dlc_path = os.path.join(self.run_path, dlc)

            for root, dirs, files in os.walk(dlc_path):

                for file in files:
                    if file.upper().endswith('$PJ'):
                        lc = os.path.splitext(file)[0]

                        self.prj_path[lc] = root
                        self.lc_list.append(lc)

                        self.lc_dlc[lc] = dlc

    def data_exchange(self, dlc, lc):

        print('%s start time: %s' %(lc, time.strftime('%H:%M:%S', time.localtime(time.time()))))
        # print(os.getpid())
        # write header
        if self.tran_type == 'flex':

            header  = '%s%s' %(lc, '\n')
            header += '%s%s' %((time.strftime('%m.%d.%Y', time.localtime(time.time()))),'\n')
            header += '%s%s' %((time.strftime('%H:%M:%S', time.localtime(time.time()))),'\n')
            header += '%s' % ((len(self.vars_list)-1) if not self.node_list else (len(self.vars_list)-1)*len(self.node_list))
            # header.replace('\n', '\n')
        else:
            header  = '%s%s' %(lc, '\n')
            for var in self.vars_list:
                header += '\t%s' %(var if 'Time' not in var else 'Time')
            header += '\n'
            for unit in self.unit_list:
                header += '\t(%s)' %unit

        # get dimensions for load case
        dimensions = {}
        for ext in self.exts_list:
            percent_path = os.path.join(self.prj_path[lc], lc + '.%' + ext)

            with open(percent_path, 'r') as f:
                lines = f.readlines()
                for line in lines:
                    if line.startswith('DIMENS'):
                        dimensions[ext] = line.strip().split()[1:]
                        break

        # create a empty result array
        data_length = int(dimensions[self.exts_list[0]][-1])
        if not self.node_list:
            if int(self.step_num) == 1:
                result = np.zeros((data_length, len(self.vars_list)))
            else:
                result = np.zeros((int(data_length/int(self.step_num)), len(self.vars_list)))
        else:
            if int(self.step_num) == 1:
                result = np.zeros((data_length, (len(self.vars_list)-1)*len(self.node_list)+1))
            else:
                result = np.zeros((int(data_length/int(self.step_num)), len(self.vars_list)*len(self.node_list)))
        # print(result.shape)

        # to sort extension for getting the variable result in the same extension
        sort_result = sort_extension_index(self.vars_list, self.exts_list)
        # print(sort_result)

        # get result for all variable result in each extension
        # other than get the variable result one by one
        for ext in sort_result[0]:

            data = rb(lc, self.prj_path[lc], ext).read_dollar()

            # variables listed in ext file
            vars_list = sort_result[1][ext][0]
            # print(vars_list)
            # variable and index listed in variable file
            var_index = dict(zip(sort_result[1][ext][0], sort_result[1][ext][1]))
            # unit transfer array
            vars_unit = dict(zip(self.vars_list, self.unit_list))
            unit_tran = unit_exchange(vars_list, vars_unit)

            if len(dimensions[ext]) == 2:
                # variable index listed in dollar file
                index_ori = [self.var_index[var] for var in vars_list]
                # variable index for each output file
                index_out = [var_index[var] for var in vars_list]

                if int(self.step_num) == 1:
                    result[:, index_out] = (data[1][:, index_ori])*unit_tran
                else:
                    result[:, index_out] = (data[1][::int(self.step_num), index_ori])*unit_tran

            # tower/blade
            elif len(dimensions[ext]) == 3:
                # variable index listed in dollar file
                index_ori = [i*int(dimensions[ext][0])+j for i in self.sec_index for j in self.inde_list[1:]]
                # print(index_ori)

                # variable index for each output file
                index_var = [var_index[var] for var in vars_list]
                # print(index_var)
                index_out = [i*len(vars_list)+j for i in range(len(self.node_list)) for j in index_var]
                unit_tran = [j for i in self.node_list for j in unit_tran]
                # print(index_out)

                # transfer matrix
                trans = np.zeros((len(index_ori), len(index_ori)), dtype=int)
                num   = len(vars_list)
                for i in range(len(self.secs_list)):
                    trans[num*i:num*(i+1),num*i:num*(i+1)] = self.transfer_matrix[self.matr_type]

                if int(self.step_num) == 1:
                    res = data[0].reshape((int(dimensions[ext][2]),int(dimensions[ext][0])*int(dimensions[ext][1])))
                    res = res[:, index_ori]*unit_tran
                else:
                    res = data[0].reshape((int(dimensions[ext][2]),int(dimensions[ext][0])*int(dimensions[ext][1])))
                    res = res[::int(self.step_num), index_ori]*unit_tran

                result[:,index_out] = np.dot(res, trans)
                # print(result)

        # time is not zero for variable 'Time from start of output' in Bladed
        result[:, 0] = result[:, 0] - result[0, 0]

        # eliminate the value below 1e-45
        result = np.where(abs(result)>1E-45, result, 0)

        result_path = os.sep.join([self.res_path, dlc])
        if not os.path.exists(result_path):
            os.makedirs(result_path)

        if (self.tran_type == 'bladed') and (self.data_form == 'ascii'):
            result_file = os.path.join(result_path, lc+'.txt')
            with open(result_file, 'w') as f:
                # the following format is right for GUI!!!
                # the results from script and GUI are different
                np.savetxt(f, result, fmt='%8.6E', delimiter='\t', header=header, comments='')

        elif (self.tran_type == 'flex') and (self.data_form == 'ascii'):
            result_file = os.path.join(result_path, lc+'.txt')
            with open(result_file, 'w') as f:
                # the following format is right for GUI!!!
                # the results from script and GUI are different
                np.savetxt(f, result, fmt='%15.6E', delimiter='', header=header, comments='')

        elif (self.tran_type == 'flex') and (self.data_form == 'binary'):
            result_file = os.path.join(result_path, lc+'.bin')
            with open(result_file, 'wb') as f:
                f.write(struct.pack('i', 72))
                f.write(struct.pack('i', datetime.datetime.now().year))
                f.write(struct.pack('i', datetime.datetime.now().month))
                f.write(struct.pack('i', datetime.datetime.now().day))
                f.write(struct.pack('i', datetime.datetime.now().hour))
                f.write(struct.pack('i', datetime.datetime.now().minute))
                f.write(struct.pack('i', datetime.datetime.now().second))
                f.write(struct.pack('160s', (lc+(160-len(lc))*'#').encode()))
                f.write(struct.pack('i', 72))
                f.write(struct.pack('i', 4480))
                f.write(struct.pack('i', (len(self.vars_list)-1)*len(self.node_list)))
                f.write(struct.pack('i', 4480))
                data_num = result.size

                # for i in range(result.shape[0]):
                #     for j in range(result.shape[1]):
                #         f.write(struct.pack('f', result[i,j]))
                f.write(struct.pack(data_num*'f', *tuple(result.flatten().tolist())))
            print('%s is done!' %lc)

    def run(self):

        self.get_lc_list()

        # write sensor
        if (self.tran_type=='flex') and self.vars_list and self.node_list and self.secs_list and self.res_path:
            self.write_flex_sensor(self.vars_list, self.node_list, self.secs_list, self.res_path)

        pool = mp.Pool(processes=mp.cpu_count())

        parameter_set = [[self.lc_dlc[lc], lc] for lc in self.lc_list]
        # print(len(parameter_set))
        # print(parameter_set)

        # pool.map_async(self.run_data_exchange, parameter_set)
        pool.starmap_async(self.data_exchange, parameter_set)

        pool.close()
        pool.join()

        print('%s is done!' %self.res_path.split(os.sep)[-2])

if __name__ == '__main__':

    mp.freeze_support()

    run_path = r'\\172.20.0.4\fs02\CE\V3\loop05\run_0615'
    res_dirs = r'\\172.20.0.4\fs02\CE\V3\result_test1\component\Pitch_Bearing'
    dlc_list = ['DLC12', 'DLC24', 'DLC31', 'DLC41', 'DLC64', 'DLC72']
    ind_list = [2, 0, 12, 13, 15, 8, 9]

    ext_list = ['07', '08', '22', '22', '22', '22', '22']
    var_list = ['Time from start of output', 'Blade 1 pitch angle', 'Blade root 1 Fx', 'Blade root 1 Fy', 'Blade root 1 Fz', 'Blade root 1 Mx', 'Blade root 1 My']
    unt_list = ['s', 'rad', 'kN', 'kN', 'kN', 'kNm', 'kNm']
    Get_Data(step_num=1,
             transfer_type='flex',
             run_path=run_path,
             data_format='ascii',
             res_path=res_dirs,
             dlc_list=dlc_list,
             unit_list=unt_list,
             index_list=ind_list,
             variable_list=var_list,
             extension_list=ext_list).run()




