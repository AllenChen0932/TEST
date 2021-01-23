# -*- coding: utf-8 -*-
# @Time    : 2020/07/17 15:49
# @Author  : CJG
# @File    : readRainflow_v1_0.py

import os
import numpy as np

try:
    from tool.load_report.Read_Bladed_v3 import read_bladed as ReadBladed
except:
    from Read_Bladed_v3 import read_bladed as ReadBladed

class readRainflow(object):

    def __init__(self, rfpath, content=('DEL','RFC','Markov')):
        self.rfpath = rfpath
        self.content = content
        self.varlist = []
        self.markov = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'var_range': {}, 'var_mean': {}, 'var_cycles': {}}
        self.RFC    = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'var_rfc': {}}
        self.DEL    = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'm': None, 'var_del': {}}
        # self.caseDEL = {'perfilelist':[],'varlist':[],'per_var':{},'m':None,'caselist':[],'var_del':{}} #unused now
        # self.caseTime = {'caselist': [], 'casetimelist': []}  # unused now

        self.read_rainflow()

    def read_rainflow(self):

        path = self.rfpath
        for file in os.listdir(path):
            if os.path.splitext(file)[1][0:2] == '.%':
                [pername, perext] = os.path.splitext(file)
                perext = perext[2:]

                # read percent and dollar file USING 'ReadBladed'
                # For now, ALL variables and contents ('DEL','RFC','Markov') in the given rfpath are extracted,
                # though for example only 'DEL' are required to ouput.
                res = ReadBladed(pername, path, perext)
                # print(res.channel)
                # print(res.var_list)

                if 'Rainflow cycle distribution' in res.channel and 'Markov' in self.content:
                    self.markov['perfilelist'].append(file)
                    var = res.channel.split('[')[1].split(']')[0]
                    self.markov['varlist'].append(var)
                    self.markov['per_var'][perext] = var
                    self.markov['var_range'][var]  = res.axis[0]['tick']
                    self.markov['var_mean'][var]   = res.axis[1]['tick']
                    self.markov['var_cycles'][var] = res.data[:,:,0]
                    self.markov['datalen'] = [int(res.dim[2]),int(res.dim[1])]
                    self.varlist = self.markov['varlist']

                elif 'Equivalent load' in res.channel and 'DEL' in self.content:
                    self.DEL['perfilelist'].append(file)
                    var = res.channel.split('[')[1].split(']')[0]
                    self.DEL['varlist'].append(var)
                    self.DEL['per_var'][perext] = var
                    self.DEL['m'] = res.axis[0]['tick']  # array containing strings
                    self.DEL['var_del'][var] = res.data
                    self.DEL['datalen'] = res.data_len
                    self.varlist = self.DEL['varlist']

                elif 'Rainflow cycles by range' in res.channel and 'RFC' in self.content:
                    self.RFC['perfilelist'].append(file)
                    var = res.channel.split('[')[1].split(']')[0]
                    self.RFC['varlist'].append(var)
                    self.RFC['per_var'][perext] = var
                    self.RFC['var_rfc'][var] = res.data
                    self.RFC['datalen'] = res.data_len
                    self.varlist = self.RFC['varlist']

                elif 'Lifetime Weighted Equivalent Loads' in res.channel:
                    pass  # to develop in the future

                elif 'Time Per Load Case' in res.channel:
                    pass  # to develop in the future

        # print(self.markov['var_range']['Blade root 1 My'])
        # print(self.varlist)
        # print(self.RFC['per_var'])
        # print(self.DEL)
if __name__ == '__main__':

    rf_path = r'\\172.20.0.4\fs02\LSY\V3\loop04_backup20200623\BackUp\post\rainflow\main_25y_1E7\Blade_root'
    readRainflow(rf_path, content=('DEL','RFC','Markov'))


