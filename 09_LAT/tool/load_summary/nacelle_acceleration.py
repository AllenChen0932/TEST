# -*- coding: utf-8 -*-
# @Time    : 2020/07/31 11:10
# @Author  : CE
# @File    : nacelle_acceleration.py


import os

class WriteNacelleAcc(object):

    def __init__(self, pj_path, table, loop_name, loop_index):

        self.ult_path = pj_path
        self.ls_table = table
        self.lp_name  = loop_name
        self.lp_index = loop_index

        self.read_dollarMX()
        self.write_result()

    def read_dollarMX(self):
        
        pj_list = [(file,root) for root,dirs,files in os.walk(self.ult_path) for file in files if '.$MX' in
                   file.upper()]
        if not pj_list:
            raise Exception('No $PJ in %s' %self.ult_path)

        pj_file = pj_list[0]
        mx_file = os.path.join(pj_file[1], os.path.splitext(pj_file[0])[0]+'.$MX')
        if not os.path.isfile(mx_file):
            raise Exception('No $MX in %s' %self.ult_path)
       
        with open(mx_file, 'r') as f:
            lines = f.readlines()[2:6]

            fa_acc_max = lines[0].strip().split('\t')
            fa_acc_min = lines[1].strip().split('\t')
            self.fore_aft = [float(fa_acc_max[3].strip()), float(fa_acc_min[3].strip())]
            self.fore_aft.sort(key=lambda x: abs(x))
            # print(self.fore_aft)

            ss_acc_max = lines[2].strip().split('\t')
            ss_acc_min = lines[3].strip().split('\t')
            self.side_side = [float(ss_acc_max[4].strip()), float(ss_acc_min[4].strip())]
            self.side_side.sort(key=lambda x: abs(x))
            # print(self.side_side)

    def write_result(self):
        '''write result to load summary sheet'''
        # write result
        sheet_names = self.ls_table.sheetnames
        new_sheet = 'Load Summary'

        if new_sheet in sheet_names:
            sheet = self.ls_table.get_sheet_by_name(new_sheet)
        else:
            sheet = self.ls_table.create_sheet(new_sheet)

        col_start = 3*self.lp_index

        sheet.cell(row=15, column=col_start+1, value='Nacelle Acceleration')
        sheet.cell(row=16, column=col_start+1, value='Path:')
        sheet.cell(row=16, column=col_start+2, value=self.ult_path)
        sheet.cell(row=16, column=col_start+3, value=' ')

        sheet.cell(row=17, column=col_start+1, value='FA(m/s^2)')
        sheet.cell(row=17, column=col_start+2, value=self.fore_aft[-1])
        sheet.cell(row=18, column=col_start+1, value='SS(m/s^2)')
        sheet.cell(row=18, column=col_start+2, value=self.side_side[-1])

        print('Nacelle acceleration is done!')


