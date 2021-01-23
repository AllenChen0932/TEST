#!/usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019.09.24
# @Author  : cjg
# @File    : mainBearingLife.py

'''
calculate the main bearing life
based on the collected LDD loads from Bladed result files
'''

import os
import re
import numpy as np

def read_header(filename):
    """
    :param filename  : .% file path_name (load probability distribution result file)
    :return          : var name, data dimension, format (Binary or ASCII)
    """
    with open(filename, 'r') as f:
        tmpf = f.read()

        pattern1 = re.compile(r'\bACCESS\b\s+([D.S])')  # D=Binary or S=Ascii
        # fFormat = pattern1.findall(tmpf)
        fFormat = pattern1.search(tmpf).group(1)
        # print('Format of this .$ file is (D=Binary, S=Ascii): '+fFormat)

        pattern2 = re.compile(r'\bDIMENS\b\s+(\d+)\s+(\d+)')  # in this case, NDIMENS = 2
        # dim = pattern2.findall(tmpf)
        varNum = int(pattern2.search(tmpf).group(1))
        binNum = int(pattern2.search(tmpf).group(2))
        # print('Dimension of data in this .$ file is: [%d, %d].' % (varNum, binNum))

        pattern3 = re.compile(r'\bAXISLAB\b\s+\'(.*)\'')
        # varName = pattern3.findall(tmpf)
        varName = pattern3.search(tmpf).group(1)
        # print('Variable stored in this .$ file is: \'%s\'' % varName)

    return varName, fFormat, varNum, binNum

def read_dollar(filename, fFormat, varNum, binNum, varInd):
    """
    read .$ data file based on data format: D=Binary, S=ASCII
    :param filename :  .$ file path+name
    :param fFormat  :  data format of the .$ file
    :param varNum   :  number of variables in the .$ file
    :param binNum   :  bin number of each variable (probability distribution result)
    :param varInd   :  index of a specified variable in the .$ file --- [1,2,...,varNum]
    :return         :  the specified variable array in the .$ file  --- ndarray type
    """
    if fFormat == 'D':
        with open(filename, 'rb') as f:
            data = np.fromfile(f, np.float32)
            # print(data.shape)
            # print(type(data))
            # print(data)
            dataLength = varNum * binNum  # dataLength = len(data)
            y = data[list(range(int(varInd - 1), dataLength - 1, int(varNum)))]
            # print(y.shape)
            # print(y)
        return y

    elif fFormat == 'S':
        data = np.loadtxt(filename)
        # print(data.shape)
        # print(type(data))
        # print(data)
        y = data[:, varInd - 1]
        # print(y.shape)
        # print(y)
        return y

def read_all_hub_loads(folderPath):
    """
    read all the hub load result files and
    collect the [Fx Fy Fz Mx My Mz] LDD data.
    :param filePath  : folder path containing .%, .$ files
    :return          : LDD = [Mx, dT, My, dT, ..., Fz, dT]
    """
    LDDlist = ['Stationary hub Mx', 'Stationary hub My', 'Stationary hub Mz',
               'Stationary hub Fx', 'Stationary hub Fy', 'Stationary hub Fz']
    tmp = 0
    for maindir, subdir, file_name_list in os.walk(folderPath):
        for filename in file_name_list:
            if '.%' in filename:
                # read .% header file
                headerpath = os.path.join(maindir, filename)
                varName, fFormat, varNum, binNum = read_header(headerpath)
                # print(headerpath)

                # access the related .$ data file
                header = filename.split('.%')
                dollarname = header[0] + '.$' + header[1]
                dollarpath = os.path.join(maindir, dollarname)
                # print(dollarpath)

                # read .$ data file and collect LDD
                if tmp == 0:
                    LDD = np.zeros([binNum, 2 * len(LDDlist)])
                tmp += 1
                pos = LDDlist.index(varName)
                LDD[:, 2*pos]   = read_dollar(dollarpath, fFormat, varNum, binNum, 1)  # load
                LDD[:, 2*pos+1] = read_dollar(dollarpath, fFormat, varNum, binNum, 2)  # load duration
    # print(LDD.shape)
    return LDD

def get_equivalent_fatigue(load, time):
    """
    translate the LDD loads into equivalent negative and positive loads
    :param load : load bins
    :param time : load duration
    :return     : equivalent negtive and positive loads and time ratios
    """
    load_bin = np.multiply(np.power(np.abs(load), 4.05), time)
    # print(load_bin)
    sum_load_neg = np.sum(load_bin[load < 0])
    sum_load_pos = np.sum(load_bin[load >= 0])
    sum_time_neg = np.sum(time[load < 0])
    sum_time_pos = np.sum(time[load >= 0])

    if sum_time_neg != 0:
        loadNeg = -1 * np.power(sum_load_neg/sum_time_neg, 1/4.05)
    else:
        loadNeg = np.nan
    if sum_time_pos != 0:
        loadPos = np.power(sum_load_pos/sum_time_pos, 1/4.05)
    else:
        loadPos = np.nan
    timeRatioNeg = sum_time_neg/(sum_time_neg+sum_time_pos)
    timeRatioPos = sum_time_pos/(sum_time_neg+sum_time_pos)

    out = np.array([[loadNeg, timeRatioNeg], [loadPos, timeRatioPos]])
    return out

def get_load_case(case_result, add_load):
    """
    put add_load into case_result in some way
    :param case_result:
    :param add_load:
    :return:
    """
    # print(type(add_load[1,0]))
    if np.isnan(add_load[0, 0]):
        # no negative loads
        add_case_rep = np.tile(add_load[1, 0], (case_result.shape[0], 1))
        load_case    = np.hstack((add_case_rep, case_result))

    elif np.isnan(add_load[1, 0]):
        # no positive loads
        add_case_rep = np.tile(add_load[0, 0], (case_result.shape[0], 1))
        load_case    = np.hstack((add_case_rep, case_result))

    else:
        add_case_rep    = np.repeat(add_load, case_result.shape[0], axis=0)  # 按行复制
        case_result_rep = np.tile(case_result, (2, 1))  # 按块复制
        load_case       = np.hstack((add_case_rep[:, :-1], case_result_rep))
        time_ratio_new  = np.multiply(case_result_rep[:, -1], add_case_rep[:, -1])
        load_case[:,-1] = time_ratio_new

        np.savetxt('load_case.txt', load_case)
        # print(case_table)
    return load_case

def cal_main_bearing_life(case_table):
    """
    calculate the main bearing life based on the given case table
    :param case_table:
    :return: bearing life
    """
    # rotor parameters:
    rotor_speed = 10.66  # [rpm] 风轮额定转速
    tilt = 6             # [deg] 倾角主轴
    L1 = 1.78            # [m] 轮毂中心与主轴承中心距离
    L2 = 0.7225          # [m] 主轴承中心与齿轮箱支撑臂中心距离
    L3 = 3               # [m] 轮毂重心与齿轮箱的重心的距离
    L4 = 0.461           # [m] 齿轮箱两支撑臂的距离
    L5 = 3.688           # [m] 风轮重心与轮毂中心的距离
    Gs = 12362*10/1000   # [kN] 主轴重量
    Gg = 24500*10/1000   # [kN] 齿轮箱的重量

    # bearing parameters:
    C0 = 25000  # [kN] 轴承基本额定静负荷
    C  = 16215  # [kN] 轴承基本额定动负荷，相对寿命与此变量没有关系
    Pu = 1460   # [kN] 轴承疲劳负荷极限
    e  = 0.26   # 更改成0.35没用影响
    X0 = 1
    X1 = 0.67  # 径向负荷系数
    X2 = 0.67
    Y0 = 2.5
    Y1 = 2.6  # 轴向负荷系数
    Y2 = 3.9

    life_loadcase = np.zeros(case_table.shape[0])
    for ii in range(0, case_table.shape[0]):
        # axial force
        Fxb = case_table[ii, 0]+np.sin(tilt/180*np.pi)*(Gs+Gg)
        # horizontal radial force
        Fyb = (case_table[ii, 4]-case_table[ii, 1]*(L1+L2+L3))/L3
        # vertical radial force
        Fzb = (Gs*np.cos(tilt/180*np.pi)*(L3-L4)-Gg*np.cos(tilt/180*np.pi)*(L5-L3) -
               case_table[ii, 2]*(L1+L2+L3)-case_table[ii, 3])/L3
        # axial resultant force
        Fxyb = np.sqrt(Fyb**2+Fzb**2)
        # equivalent load
        if np.abs(Fxb/Fxyb) < np.e:
            P0b = Fxyb+Y1*np.abs(Fxb)
        else:
            P0b = 0.67*Fxyb+Y2*np.abs(Fxb)
        L10h = 1e6*np.power(C/P0b, 3.333)/(60*rotor_speed)
        life_loadcase[ii] = case_table[ii, 5]/L10h

    life_bearing = 1/np.sum(life_loadcase)

    return life_bearing

def handle(ldd, excel, loop_name, loop_index):

    ldd_path   = ldd
    table      = excel

    LDD = read_all_hub_loads(ldd_path)

    My_eq = get_equivalent_fatigue(LDD[:, 2],  LDD[:, 3])
    Mz_eq = get_equivalent_fatigue(LDD[:, 4],  LDD[:, 5])
    Fx_eq = get_equivalent_fatigue(LDD[:, 6],  LDD[:, 7])
    Fy_eq = get_equivalent_fatigue(LDD[:, 8],  LDD[:, 9])
    Fz_eq = get_equivalent_fatigue(LDD[:, 10], LDD[:, 11])

    case_table = Mz_eq
    case_table = get_load_case(case_table, My_eq)
    case_table = get_load_case(case_table, Fz_eq)
    case_table = get_load_case(case_table, Fy_eq)
    case_table = get_load_case(case_table, Fx_eq)

    life_bearing = cal_main_bearing_life(case_table)

    # 输出到excel
    # table = load_workbook(excel_path)
    sheet_names = table.sheetnames
    new_sheet   = 'Load Summary'

    if new_sheet in sheet_names:
        sheet = table.get_sheet_by_name(new_sheet)
    else:
        sheet = table.create_sheet(new_sheet)

    col_start = 3*loop_index

    if not sheet.cell(row=1, column=col_start+1).value:
        sheet.cell(row=1, column=col_start+1, value=loop_name)

    sheet.cell(row=11, column=col_start+1, value='Main Bearing')
    sheet.cell(row=12, column=col_start+1, value='Path:')
    sheet.cell(row=12, column=col_start+2, value=ldd_path)
    sheet.cell(row=12, column=col_start+3, value=' ')
    sheet.cell(row=13, column=col_start+1, value='Life(year)')
    sheet.cell(row=13, column=col_start+2, value=life_bearing)

if __name__ == '__main__':
    # # filePath = r'D:\03 Post Processing Tool Dev\mainBearingLife_python\LDD\Hub_stationary_ASCII\'
    # # filename_h = r'D:\03 Post Processing Tool Dev\mainBearingLife_python\LDD\Hub_stationary_ASCII\probdist.%001'
    # filename_h = r'D:\03 Post Processing Tool Dev\mainBearingLife_python\LDD\Hub_stationary_Binary\probdist.%001'
    # varName, fFormat, varNum, binNum = read_header(filename_h)
    # # tmp = read_header(filename)
    # # print(tmp[0])
    #
    # filename_d = r'D:\03 Post Processing Tool Dev\mainBearingLife_python\LDD\Hub_stationary_Binary\probdist.$001'
    # # filename_d = r'D:\03 Post Processing Tool Dev\mainBearingLife_python\LDD\Hub_stationary_ASCII\probdist.$001'
    # read_dollar(filename_d, fFormat, varNum, binNum, 1)

    handle(r'E:\05 TASK\02_Tools\load summary append\post\LDD_Hub',
           r"E:\05 TASK\02_Tools\load summary append\LoadComparisonTemplate.xlsm")






