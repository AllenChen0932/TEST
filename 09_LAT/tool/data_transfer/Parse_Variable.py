#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/8/2020 4:04 PM
# @Author  : CE
# @File    : Parse_Variable.py

import linecache

class Get_Variable(object):
    
    def __init__(self, var_path):
        
        self.var_path = var_path
        
        self.number_list  = []
        self.file_list    = []
        self.vars_list    = []
        self.file_dlc_var = {}
        self.vars_no_unit = {}

        self.parse_variable()
        # print(self.vars_no_unit)
    
    def parse_variable(self):
        '''parse variable setting'''
    
        lineNumber  = 1

        # get variable setting between ##
        with open(self.var_path, 'r') as f:
            lines = f.readlines()

            for line in lines:
                if line.startswith('#') and len(line)>1:
                    self.number_list.append(lineNumber)
                
                if (('#' in line) and (len(line)==1)) or not len(line):
                    self.number_list.append(lineNumber)
                lineNumber += 1

        if not len(self.number_list)%2 == 0:
            raise Exception('Please check the variable definition file!')

        for j in range(len(self.number_list)):
            if j%2 == 0:
                start = self.number_list[j]
                end   = self.number_list[j+1]

                file_name = linecache.getline(self.var_path,start).strip()[1:]
                dlc_list  = linecache.getline(self.var_path,start+1).strip().split(',')

                if 'section'.upper() not in file_name.upper():
                    var_list  = linecache.getlines(self.var_path)[start+1:end-1]

                    vars_list   = []
                    unit_list   = []
                    var_no_unit = []

                    for var in var_list:

                        if "#" not in var:

                            if ',' in var:
                                temp = var.strip().split(',')

                                vars_list.append(temp[0])
                                unit_list.append(temp[1])
                                self.vars_list.append(temp[0])
                            else:
                                temp = var.strip()
                                var_no_unit.append(temp)

                    self.file_list.append(file_name)

                    if var_no_unit:
                        self.vars_no_unit[file_name] = var_no_unit

                    self.file_dlc_var[file_name] = [dlc_list, vars_list, unit_list]
                else:
                    var_list = linecache.getlines(self.var_path)[start+1:end-2]
                    # print(var_list)
                    vars_list = []
                    unit_list = []

                    section = linecache.getlines(self.var_path)[end-2:end-1][0]
                    section = section.strip().split(',')[1:]

                    var_no_unit = []

                    for var in var_list:
                        # eliminate comments
                        if "#" in var:
                            pass
                        if ',' in var:
                            temp = var.strip().split(',')

                            vars_list.append(temp[0])
                            unit_list.append(temp[1])

                            self.vars_list.append(temp[0])
                        else:
                            temp = var.strip()
                            var_no_unit.append(temp)

                    self.file_list.append(file_name)

                    if var_no_unit:
                        self.vars_no_unit[file_name] = var_no_unit

                    self.file_dlc_var[file_name] = [dlc_list, vars_list, unit_list, section]

if __name__ == '__main__':

    config = r"C:\Users\10700700\Desktop\tool\py\17_data_exchange\config2.0.dat"

    Get_Variable(config)