# -*- coding: utf-8 -*-
# @Time    : 2019/10/18 23:07
# @Author  : CE
# @File    : Equivalent_torque.py

import os
import numpy as np
import configparser

def get_ext(path, pj_name):
    '''
    读取结果文件的后缀及每个后缀所包含的信息
    :param path:
    :param pj_name:
    :return:
    '''

    ext_var = { }
    ext_dim = { }
    file_format = None

    # get all extension under path
    file_list = os.listdir(path)
    temp      = [file for file in file_list if file.startswith(pj_name)]
    ext_list  = [file.split('%')[1] for file in temp if (file.find('%') != -1)]
    if not ext_list:
        raise Exception('No result under %s!' %path)

    for ext in ext_list:
        temp_path = os.path.join(path, '.%'.join([pj_name, ext]))

        with open(temp_path, 'r') as f:
            for line in f.readlines():

                if 'ACCESS'in line:
                    file_format = line.split()[-1]
                    continue

                if line.startswith('DIMENS'):
                    ext_dim[ext] = line.strip().split()[1:]
                    continue

                if 'AXISLAB' in line:
                    variable = line.split("'")[-2]
                    ext_var[ext] = variable
                    continue

    return ext_var, ext_dim, file_format

def get_pj(path):

    file_list = os.listdir(path)
    pj_list   = [file for file in file_list if os.path.splitext(file)[1].upper() == '.$PJ']

    for pj in pj_list[::-1]:
        pj_path = os.path.join(path, pj)

        with open(pj_path, 'r') as f:
            # eliminate the project which is not ldd
            if 'MSTART PROBD' not in f.read():
                pj_list.remove(pj)
                continue

    return pj_list

def get_data(ldd_path):

    hub_mx = None
    time   = None

    pj = get_pj(ldd_path)
    # print(pj)

    # 提示路径是否有错
    if len(pj) > 1:
        raise Exception('More than 2 projects under %s!' %ldd_path)
    elif len(pj) == 0:
        raise Exception('No project under %s!' % ldd_path)

    pj_name = os.path.splitext(pj[0])[0]

    percent_result = get_ext(ldd_path, pj_name)
    ext_var        = percent_result[0]
    ext_dim        = percent_result[1]
    file_format    = percent_result[2]

    ext = [key for key, value in ext_var.items()
           if value == 'Stationary hub Mx' or value == 'Rotating hub Mx']
    # print(ext)

    percent_path = os.path.join(ldd_path, '.$'.join([pj_name, ext[0]]))
    # print(percent_path)

    if file_format == 'D':
        with open(percent_path, 'r') as f:
            ori_data = np.fromfile(f, np.float32)
            data     = ori_data.reshape((int(ext_dim[ext[0]][1]), int(ext_dim[ext[0]][0])))

            hub_mx   = data[:, 0]
            time     = data[:, 1]

    elif file_format == 'S':
        with open(percent_path, 'r') as f:
            ori_data =  np.loadtxt(f, np.float32)
            data     = ori_data.reshape((int(ext_dim[ext[0]][1]), int(ext_dim[ext[0]][0])))

            hub_mx   = data[:, 0]
            time     = data[:, 1]

    return hub_mx, time

def handle(ldd_path, table, loop_name, loop_index):

    col_start     = 3*loop_index
    sum_eq_torque = 0

    sheet_names = table.sheetnames
    new_sheet   = 'Load Summary'

    if new_sheet in sheet_names:
        sheet = table.get_sheet_by_name(new_sheet)
    else:
        sheet = table.create_sheet(new_sheet)

    if not sheet.cell(row=1, column=col_start+1).value:
        sheet.cell(row=1, column=col_start+1, value=loop_name)

    sheet.cell(row=2, column=col_start+1, value='Gearbox')
    sheet.cell(row=3, column=col_start+1, value='Path:')
    sheet.cell(row=3, column=col_start+2, value=ldd_path)
    sheet.cell(row=3, column=col_start+3, value=' ')

    result = get_data(ldd_path)

    hub_mx   = result[0]
    time     = result[1]
    sum_time = sum(time)

    # get pitch bearing parameters
    m = None
    try:
        config = configparser.ConfigParser()
        config.read('config_m.dat', encoding='utf-8')
        if config.has_section('Gearbox Equivalent Torque'):
            m = config.get('Gearbox Equivalent Torque', 'm')
    except Exception:
        raise Exception('Error occurs in reading config_m.dat!')

    for i in range(len(hub_mx)):
        sum_eq_torque += time[i]*(hub_mx[i]**2)**(float(m)/2)

    Teq = (sum_eq_torque/sum_time)**(1/float(m))/1000

    sheet.cell(row=4, column=col_start+1, value='T_eqv(kNm)')
    sheet.cell(row=4, column=col_start+2, value=Teq)
