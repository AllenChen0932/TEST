#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/1/2020 8:04 PM
# @Author  : CE
# @File    : get_var_dlc.py

import os

class Get_Variable_DLC(object):

    def __init__(self, project, run_path):

        self.prj_path = project
        self.run_path = run_path

        self.fatigue_list  = []
        self.ultimate_list = []

        self.twr_height    = []
        self.bld_radius    = []
        self.twr_output    = []
        self.bra_output    = []    # blade root axis
        self.baa_output    = []    # blade aerodynamic axis
        self.brp_output    = []    # blade principle axis
        self.bua_output    = []    # blade user axis

        self.twr_model     = None

        self.br_var_list   = []
        self.yb_var_list   = []
        self.tr_var_list   = []
        self.hr_var_list   = []
        self.hs_var_list   = []
        self.bp_var_list   = []
        self.br_var_list   = []
        self.ba_var_list   = []
        self.bu_var_list   = []

        # self.get_loadcase()
        # self.get_variable()

    def get_loadcase(self):

        files   = os.listdir(self.run_path)
        lc_list = [file for file in files if os.path.isdir(os.sep.join([self.run_path, file]))]

        # sort load case
        fat_case = ' '.join(['DLC12', 'DLC24', 'DLC31', 'DLC41', 'DLC64', 'DLC72'])
        self.ultimate_list = [lc for lc in lc_list if lc not in fat_case]
        self.fatigue_list  = [lc for lc in lc_list if lc in fat_case]
        # print(self.ultimate_list)
        # print(self.ultimate_list)
        #
        # print(lc_list)
        return self.ultimate_list, self.fatigue_list

    def get_variable(self):

        pj_path, pj_file = os.path.split(self.prj_path)
        print(pj_path, pj_file)

        with open(self.prj_path) as f:

            lines = f.readlines()
            for line in lines:

                if line.startswith('RJ'):
                    self.bld_radius = line.split()[1].split(',')

                if line.startswith('Tj'):
                    self.twr_height = line.split()[1].split(',')

                if line.startswith('TMODEL'):
                    self.twr_model = line.split()[1]

        pj_name = pj_file.split('.$')[0]
        br_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))

        bs1_file = os.path.join(pj_path, '.%'.join([pj_name,'41']))
        bs2_file = os.path.join(pj_path, '.%'.join([pj_name,'42']))
        bs3_file = os.path.join(pj_path, '.%'.join([pj_name,'43']))
        bp1_file = os.path.join(pj_path, '.%'.join([pj_name,'15']))
        bp2_file = os.path.join(pj_path, '.%'.join([pj_name,'16']))
        bp3_file = os.path.join(pj_path, '.%'.join([pj_name,'17']))
        ba1_file = os.path.join(pj_path, '.%'.join([pj_name,'59']))
        ba2_file = os.path.join(pj_path, '.%'.join([pj_name,'60']))
        ba3_file = os.path.join(pj_path, '.%'.join([pj_name,'61']))
        bu1_file = os.path.join(pj_path, '.%'.join([pj_name,'62']))
        bu2_file = os.path.join(pj_path, '.%'.join([pj_name,'63']))
        bu3_file = os.path.join(pj_path, '.%'.join([pj_name,'64']))
        hr_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))
        hs_file = os.path.join(pj_path, '.%'.join([pj_name,'23']))
        yb_file = os.path.join(pj_path, '.%'.join([pj_name,'24']))
        tr_file = os.path.join(pj_path, '.%'.join([pj_name,'25']))






if __name__ == '__main__':

    pj_path  = r"E:\05 TASK\02_Tools\01_Load Summary\DLC12\12_aa-01\12_aa-01.$PJ"
    run_path = r'\\172.20.0.4\fs03\CE\W6250-172-14m\run_0429'
    Get_Variable_DLC(pj_path, run_path=None)