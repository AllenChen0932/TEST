#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/1/2020 8:04 PM
# @Author  : CE
# @File    : Get_LoadCase.py

import os

class Get_DLC(object):

    def __init__(self, run_path):

        self.run_path = run_path

        self.fatigue_list  = []
        self.ultimate_list = []
        self.loadcase_list = []
        self.ul_dlc12_list = []

        self.get_loadcase()

    def get_loadcase(self):

        files   = os.listdir(self.run_path)
        self.loadcase_list = [file for file in files if os.path.isdir(os.sep.join([self.run_path, file]))]

        # sort load case
        fat_case = ' '.join(['DLC12', 'DLC24', 'DLC31', 'DLC41', 'DLC64', 'DLC72'])
        self.ultimate_list = [lc for lc in self.loadcase_list if (lc[:5] not in fat_case) and lc.startswith('DLC')]
        self.fatigue_list  = [lc for lc in self.loadcase_list if lc[:5] in fat_case]

        if 'DLC12' in self.loadcase_list:
            self.ul_dlc12_list.append('DLC12')

        for lc in self.ultimate_list:
            if 'DLC11' in lc:
                self.ultimate_list.remove(lc)

        for lc in self.loadcase_list:
            if 'DLC16' in lc or 'DLC65' in lc:
                # self.ultimate_list.append(lc)
                self.fatigue_list.append(lc)
        self.fatigue_list.sort()

        for lc in self.ultimate_list:
            self.ul_dlc12_list.append(lc)

if __name__ == '__main__':

    # pj_path  = r"E:\05 TASK\02_Tools\01_Load Summary\DLC12\12_aa-01\12_aa-01.$PJ"
    run_path = r'\\172.20.0.4\fs03\CE\W6250-172-14m\run_0429'
    Get_DLC(run_path)