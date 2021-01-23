# -*- coding: utf-8 -*-
# @Time    : 2020/07/17 15:49
# @Author  : CJG
# @File    : readRainflow_v1_0.py

import os
# import numpy as np
try:
    from tool.post_export.Read_Bladed_v3 import read_bladed as ReadBladed
except:
    from Read_Bladed_v3 import read_bladed as ReadBladed

class readRainflow(object):

    def __init__(self, rfpath, content=('DEL','RFC','Markov')):
        self.rfpath  = rfpath
        self.content = content
        self.varlist = []

        self.markov = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'var_range': {}, 'var_mean': {}, 'var_cycles': {}}
        self.RFC    = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'var_rfc': {}}
        self.DEL    = {'perfilelist': [], 'varlist': [], 'per_var': {}, 'm': None, 'var_del': {}}
        # self.caseDEL = {'perfilelist':[],'varlist':[],'per_var':{},'m':None,'caselist':[],'var_del':{}} #unused now
        # self.caseTime = {'caselist': [], 'casetimelist': []}  # unused now

        self.read_rainflow()

    def read_rainflow(self):

        file_list = os.listdir(self.rfpath)
        if not file_list or len(file_list)<7:
            raise Exception('No results under %s' %self.rfpath)

        pj_list = [file for file in file_list if file.upper().endswith('.$PJ')]
        if not pj_list:
            raise Exception('No rainflow project file under path %s!' %self.rfpath)
        else:
            pj_file = pj_list[0]

        pj_path = os.path.join(self.rfpath, pj_file)
        me_path = os.path.join(self.rfpath, os.path.splitext(pj_file)[0]+'.$ME')

        with open(pj_path) as f:
            lines = f.readlines()
            for line in lines:
                if line.startswith('CALCULATION'):
                    if line.split()[-1] != '25':
                        raise Exception('%s is not a rainflow project!' %pj_path)

        if not os.path.isfile(me_path):
            raise Exception('No results under %s' %self.rfpath)
        else:
            with open(me_path) as f:
                content = f.read()
                if 'ERROR' in content.upper():
                    raise Exception('Error occurs in %s' %me_path)

        for file in os.listdir(self.rfpath):

            if os.path.splitext(file)[1][0:2] == '.%':
                [pername, perext] = os.path.splitext(file)
                perext = perext[2:]

                # read percent and dollar file USING 'ReadBladed'
                # For now, ALL variables and contents ('DEL','RFC','Markov') in the given rfpath are extracted,
                # though for example only 'DEL' are required to ouput.
                res = ReadBladed(pername, self.rfpath, perext)
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

if __name__ == '__main__':

    rf_path = r'\\172.20.4.132\fs02\CE\V3\loop06\post_1121-1\08_Fatigue\05_Main\brs1_0'
    readRainflow(rf_path, content=('DEL',))


