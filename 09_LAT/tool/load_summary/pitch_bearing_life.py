# -*- coding: utf-8 -*-
# @Time    : 2019/10/14 14:22
# @Author  : CE
# @File    : pitch_bearing_life.py

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
    ext_var = {}
    ext_dim = {}

    file_list = os.listdir(path)
    # print(file_list)

    temp     = [file for file in file_list if file.startswith(pj_name)]
    ext_list = [file.split('%')[1] for file in temp if (file.find('%') != -1)]
    if not ext_list:
        raise Exception('No result under %s!' %path)

    for ext in ext_list:
        temp_path = os.path.join(path, '.%'.join([pj_name, ext]))

        with open(temp_path, 'r') as f:

            for line in f.readlines():
                if line.startswith('DIMENS'):
                    ext_dim[ext] = line.strip().split()[1:]
                    continue

                if 'AXISLAB' in line:
                    variable = line.split("'")[-2]
                    ext_var[ext] = variable
                    break
    # print(ext_var)
    return ext_var, ext_dim

def get_pj(path):
    '''get ldd project under lrd'''
    file_list = os.listdir(path)
    pj_list   = [file for file in file_list if os.path.splitext(file)[1].upper() == '.$PJ']

    for pj in pj_list[::-1]:
        pj_path = os.path.join(path, pj)

        with open(pj_path, 'r') as f:
            content = f.read()
            # ProbDensity but not Revs
            if 'MSTART PROBD' in content and 'AZNAME	"Revs"' not in content:
                pj_list.remove(pj)

    return pj_list

def calculate(mxy, fa, fr, rev):
    '''calculate pitch bearing life'''

    # get pitch bearing parameters
    dia, Ca, m, life = None, None, None, None
    try:
        config = configparser.ConfigParser()
        config.read('config_m.dat', encoding='utf-8')
        if config.has_section('Pitch Bearing Life'):
            dia = config.get('Pitch Bearing Life', 'Nominal diameter')
            Ca = config.get('Pitch Bearing Life', 'Ca')
            m = config.get('Pitch Bearing Life', 'm')
            life = config.get('Pitch Bearing Life', 'life')
    except Exception:
        raise Exception('Error occurs in reading config_m.dat!')

    if (not fa) or (not fr):
        Pea = 2*mxy/float(dia)
    else:
        Pea = (2*mxy/float(dia)+fa+0.75*fr)

    # rev
    leqv = np.power(float(Ca)/Pea,float(m))*1000000
    damage = rev/leqv
    life_est = float(life)/sum(damage)

    # pitch equivalent torque
    T_eqv = np.power(sum(abs(mxy)**float(m)*rev/sum(rev)),1/float(m))/1000

    return life_est, T_eqv

def handle(ldd_path, table, loop_name, loop_index):

    pj = get_pj(ldd_path)
    if len(pj) > 1:
        raise Exception('More than 2 projects under %s!' %ldd_path)
    elif len(pj) == 0:
        raise Exception('No project under %s!' %ldd_path)

    pj_name = os.path.splitext(pj[0])[0]

    ext_var = get_ext(ldd_path, pj_name)[0]
    ext_dim = get_ext(ldd_path, pj_name)[1]

    mxy_ext = [key for key, value in ext_var.items() if 'Mxy' in value]
    fz_ext  = [key for key, value in ext_var.items() if 'Fz' in value]
    fxy_ext = [key for key, value in ext_var.items() if 'Fxy' in value]

    ext_list = [mxy_ext, fz_ext, fxy_ext]

    if not mxy_ext:
        raise Exception('No LRD results under path %s!' %ldd_path)

    dollar_path = ldd_path + os.sep + pj_name + '.$' + ext_list[0][0]
    with open(dollar_path, 'r') as f:
        ori_data = np.fromfile(f, np.float32)

        data = ori_data.reshape((int(ext_dim[ext_list[0][0]][1]), int(ext_dim[ext_list[0][0]][0])))
        mxy  = data[:,0]
        revs = data[:,1]/(2*np.pi)

    fxy, fz = None, None
    if ext_list[1] and ext_list[2]:
        # read fz
        dollar_path = ldd_path + os.sep + pj_name + '.$' + ext_list[1]
        with open(dollar_path, 'r') as f:
            ori_data = np.fromfile(f, np.float32)

            data = ori_data.reshape((int(ext_dim[ext_list[1][0]][1]), int(ext_dim[ext_list[1][0]][0])))
            fz   = data[:, 0]

        # read fxy
        dollar_path = ldd_path + os.sep + pj_name + '.$' + ext_list[2]
        with open(dollar_path, 'r') as f:
            ori_data = np.fromfile(f, np.float32)
            data = ori_data.reshape((int(ext_dim[ext_list[2][0]][1]), int(ext_dim[ext_list[2][0]][0])))
            fxy   = data[:, 0]

    # write result
    sheet_names = table.sheetnames
    new_sheet   = 'Load Summary'

    if new_sheet in sheet_names:
        sheet = table.get_sheet_by_name(new_sheet)
    else:
        sheet = table.create_sheet(new_sheet)

    col_start = 3*loop_index

    if not sheet.cell(row=1, column=col_start+1).value:
        sheet.cell(row=1, column=col_start+1, value=loop_name)

    sheet.cell(row=6, column=col_start+1, value='Pitch Bearing')
    sheet.cell(row=7, column=col_start+1, value='Path:')
    sheet.cell(row=7, column=col_start+2, value=ldd_path)
    sheet.cell(row=7, column=col_start+3, value=' ')

    life = calculate(mxy, fz, fxy, revs)
    sheet.cell(row=8, column=col_start+1, value='T_eqv(kNm)')
    sheet.cell(row=8, column=col_start+2, value=life[1])
    sheet.cell(row=9, column=col_start+1, value='Life(year)')
    sheet.cell(row=9, column=col_start+2, value=life[0])