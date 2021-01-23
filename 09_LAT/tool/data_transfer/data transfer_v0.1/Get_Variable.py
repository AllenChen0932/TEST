# -*- coding: utf-8 -*-
# @Time    : 2020/4/9 14:14
# @Author  : CE
# @File    : Get_Variable.py

import linecache
import os
import time

from Read_Bladed_v2 import read_bladed as rb


class Get_Variable(object):

    def __init__(self, run_path, variable_file):

        self.run_path = run_path
        self.var_path = variable_file

        # variable to output
        self.sign_numb_list = []      # list to record start line and end line for each component loads with #
        self.component_list = []
        self.component_dlcs = {}      # component: dlcs
        self.component_vars = {}      # component: variables
        self.component_secs = {}      # component: sections(for blade and tower)
        self.component_exts = {}
        self.component_type = {}      # component: [flex, binary, 2]
        self.component_unit = {}      # component: unit_list
        self.comp_varwounit = {}      # variable without unit

        # from % file
        self.dlc_list       = []      # dlc list
        self.prj_list       = []      # project list
        self.prj_path       = {}      # project  : path
        self.var_ext_index  = {}      # variable : [extension, index]
        self.ext_variables  = {}      # extension: variables
        self.variable_unit  = {}      # variable : unit
        # self.variables_ext  = {}      # variable : [extension, section]
        self.comp_section   = {}      # blade: section

        self.parse_variable()
        self.get_loadcase()
        self.get_var_unit()
        self.check_var_unit()
        self.get_var_ext()

    def parse_variable(self):

        lineNumber = 1
        start_flag = True
        end_flag   = False

        try:
            with open(self.var_path, 'r') as f:

                lines = f.readlines()
                for line in lines:
                    # print(line)
                    if line.startswith('#') and len(line) > 1 and start_flag:

                        start_flag = False
                        end_flag   = True
                        self.sign_numb_list.append(lineNumber)

                    if (line == '#\n') and end_flag:

                        start_flag = True
                        end_flag   = False
                        self.sign_numb_list.append(lineNumber)

                    lineNumber += 1
                    # print(start_flag, end_flag)
        except:
            raise Exception('Please choose a right variable input file!')

        if len(self.sign_numb_list) % 2 == 0:

            for j in range(len(self.sign_numb_list)):

                if j%2 == 0:

                    start = self.sign_numb_list[j]
                    end   = self.sign_numb_list[j+1]

                    file_name = linecache.getline(self.var_path, start).strip()[1:]
                    comp_name = file_name.split(',')[0]
                    dlc_list  = linecache.getline(self.var_path, start+1).strip().split(',')

                    self.component_list.append(comp_name)
                    self.component_dlcs[comp_name] = dlc_list
                    self.dlc_list += dlc_list
                    self.component_type[comp_name] = [_.strip() for _ in file_name.split(',')[1:]]

                    if 'section' not in comp_name.lower():

                        var_list = linecache.getlines(self.var_path)[start+1:end-1]
                        # print(var_list)
                        vars_list   = []
                        unit_list   = []
                        var_no_unit = []

                        for var in var_list:

                            if ("#" not in var) and (',' in var):
                                temp = var.strip().split(',')
                                vars_list.append(temp[0].strip())
                                unit_list.append(temp[1].strip())
                                # print(temp[0].strip(),temp[1].strip())

                            elif ("#" not in var) and (',' not in var):
                                var_no_unit.append(var.strip())

                        vars_list.insert(0, 'Time from start of output')
                        unit_list.insert(0, 's')
                        # print(unit_list)

                        self.component_vars[comp_name] = vars_list
                        self.component_unit[comp_name] = unit_list

                        if var_no_unit:
                            self.comp_varwounit[comp_name] = var_no_unit
                            raise Exception('Please define unit for the variables:\n%s'
                                            % (','.join([var for k, v in self.comp_varwounit.items() for var in v])))

                    else:

                        var_list = linecache.getlines(self.var_path)[start+1:end-2]
                        vars_list = []
                        unit_list = []

                        sec_line= linecache.getlines(self.var_path)[end-2:end-1][0]

                        if 'tower' in comp_name.lower():
                            sec_mbr  = sec_line.strip().split(',')[0].strip()
                            tr_type  = sec_line.strip().split(',')[1].strip()
                            sec_list = [int(sec.strip()) for sec in sec_line.strip().split(',')[2:]]
                            self.component_secs[comp_name] = [sec_mbr, tr_type, sec_list]
                        else:
                            sec_list = sec_line.strip().split(',')[1:]
                            if sec_list[0] != 'a':
                                self.component_secs[comp_name] = [sec.strip() for sec in sec_list]

                        var_no_unit = []

                        for var in var_list:

                            if ("#" not in var) and (',' in var):
                                temp = var.strip().split(',')
                                vars_list.append(temp[0].strip())
                                unit_list.append(temp[1].strip())
                                # print(temp[0].strip(), temp[1].strip())

                            elif ("#" not in var) and (',' not in var):
                                var_no_unit.append(var.strip())

                        vars_list.insert(0, 'Time from start of output')
                        unit_list.insert(0, 's')

                        self.component_vars[comp_name] = vars_list
                        self.component_unit[comp_name] = unit_list

                        if var_no_unit:

                            self.comp_varwounit[comp_name] = var_no_unit

                            raise Exception('Please define unit for the variables:\n%s'
                                            %(','.join([var for k,v in self.comp_varwounit.items() for var in v])))

        else:
            raise Exception('Please make sure the variable defined between #!')

        self.dlc_list = set(self.dlc_list)
        # print(self.component_vars)
        # print(self.component_type)

    def get_loadcase(self):
        '''
        get all pj under self.dir_path
        :return:
        '''

        for dlc in self.dlc_list:

            dlc_path = os.path.join(self.run_path, dlc)

            if os.path.exists(dlc_path):

                for root, dirs, files in os.walk(dlc_path):

                    for file in files:
                        # print(file)

                        if '.$PJ' in file:

                            dlc_name = file.split('.')[0]

                            self.prj_list.append(dlc_name)
                            self.prj_path[dlc_name] = root

            else:
                raise Exception('%s not exist under %s' %(dlc, self.run_path))

        self.prj_list.sort()
        # print(self.prj_path)

    def get_var_unit(self):
        '''
        get all variables and extension number for each load case,
        supposing all load case run under the same setting
        :return:
        '''

        pj_name = self.prj_list[0]
        pj_path = self.prj_path[pj_name]

        tower_mult   = False      # flag for mult-member
        tower_height = []

        prj_path = os.path.join(pj_path, pj_name+'.$PJ')
        with open(prj_path) as f:

            lines = f.readlines()
            for line in lines:

                if line.startswith('TJ'):
                    self.comp_section['tower_section'] = [float(i) for i in line.strip().split()[-1].split(',')]

                if line.startswith('RJ'):
                    self.comp_section['blade'] = [float(i.split(',')[0]) for i in line.split()[1:]][::2]

                if line.startswith('TCLOCAL'):
                    tower_mult=True
                    tower_height.append(float(line.strip().split()[-1]))
                elif line.startswith('TZORNT'):
                    tower_mult=False

                if tower_mult:
                    tower_height.append(float(line.strip().split()[-1]))

                if line.startswith('STPM'):
                    tower_bottom = [float(line.split()[1]), float(line.split()[-1])]
                    self.comp_section['tower_node'] = [tower_bottom,tower_height]

        # blade_comp = [comp for comp in self.component_list if ('blade' in comp) and ('section' in comp)]
        # if blade_comp:
        #     comp_name = blade_comp[0]
        #     sec_list = self.component_secs[comp_name]
        #     if sec_list[0].startswith('a'):
        #         self.component_secs[comp_name] = list(range(1, len(self.comp_section['blade'])))
        #     else:
        #         self.component_secs[comp_name] = [sec.strip() for sec in sec_list]
        # print(pj_name, pj_path)

        file_list = os.listdir(pj_path)
        file_list.sort()

        for file in file_list:

            if '%' in file:

                # file_path = os.path.join(pj_path, file)
                extension = file.split('%')[-1]
                variables = rb(pj_name, pj_path, extension).read_percent()

                self.ext_variables[extension] = variables[0]
                self.variable_unit.update(variables[1])

                for index, var in enumerate(variables[0]):

                    self.var_ext_index[var] = [extension, index]
        # print(self.variable_ext.keys())
    def check_var_unit(self):

        vars_unit = ['s','rad', 'deg', 'rev',
                     'n', 'kn', 'mn',
                     'nm', 'knm', 'mnm',
                     'm/s', 'rad/s', 'res/s', 'rpm',
                     'm/s^2', 'deg/s^2', 'rad/s^2',
                     'w','kw','mw']

        # check variable
        var_list  = [var for k,v in self.component_vars.items() for var in v]
        var_error = []

        for var in var_list:
            if var not in self.variable_unit.keys():
                var_error.append(var)

        if var_error:
            raise Exception('Variables not in result:\n%s' %(','.join(var_error)))

        # check unit
        unit_list  = [u for units in self.component_unit.values() for u in units]

        unit_error = []

        for unit in unit_list:

            if unit.lower() not in vars_unit:
                # print(unit.lower() not in vars_unit)
                unit_error.append(unit)

        if unit_error:
            raise Exception('Variables not in result:\n%s' %(','.join(unit_error)))

    def get_var_ext(self):
        '''
        get extension list for each component
        :return:
        '''
        for key, values in self.component_vars.items():
            # print(var, len(var))
            ext_list = [self.var_ext_index[var][0] for var in values]
            self.component_exts[key] = ext_list

if __name__ == '__main__':

    time_start = time.time()

    run_path = r'\\172.20.0.4\fs02\CE\V3\loop05\run_0615'
    var_path = r"C:\Users\10700700\Desktop\tool\py\17_data_exchange\config2.0.dat"
    res      = Get_Variable(run_path, var_path)
    # print(res.component_vars['Pitch_Bearing'])
    # print(res.component_exts['Pitch_Bearing'])
    # print(res.component_dlcs['Pitch_Bearing'])
    # print(res.component_unit['Pitch_Bearing'])

    time_end = time.time()
    print('Total time: ',time_end-time_start)