#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/21/2020 8:53 AM
# @Author  : CE
# @File    : Write_Bstats.py

import os

class Bstats(object):

    def __init__(self, run_path, dlc_list, chan_dict, post_path, only_maxmin=True):

        self.run_path  = run_path
        self.dlc_list  = dlc_list
        self.chan_dict = chan_dict     #{channel:[[variable, extension, section_list]]}
        self.post_path = post_path
        self.mami_flag = only_maxmin

        self.lc_path   = {}
        self.lc_list   = []

        self.get_loadcase()
        self.write_pj()
        self.write_in()

    def get_loadcase(self):
        '''
        get subgroup sorted by mean and half
        :return:
        '''
        self.dlc_list.sort()

        for dlc in self.dlc_list:

            dlc_path = os.path.join(self.run_path, dlc)
            lc_list  = os.listdir(dlc_path)

            for lc in lc_list:

                lc_path = os.path.join(dlc_path, lc)

                if os.path.isdir(lc_path):

                    self.lc_path[lc] = lc_path
                    self.lc_list.append(lc)

    def write_pj(self):

        for key, value in self.chan_dict.items():

            # header
            content = '<?xml version="1.0" encoding="ISO-8859-1" ?>\n' \
                      '<BladedProject ApplicationVersion="4.8.0.41">\n' \
                      '	<DataModelVersion>0.8</DataModelVersion>\n' \
                      '	<BladedData dataFormat="project">\n' \
                      '<![CDATA[\n' \
                      'VERSION	4.8.0.34\n' \
                      'MULTIBODY	 1\n' \
                      'CALCULATION	16\n' \
                      'OPTIONS	0\n' \
                      'PROJNAME	\n' \
                      'DATE	\n' \
                      'ENGINEER	\n' \
                      'NOTES	""\n' \
                      'PASSWORD	\n' \
                      'MSTART BSTAT\n' \
                      'TYPE	2\n' \
                      "ATTRIBF	''\n" \
                      'VARIAB	""\n' \
                      'NDIMENS	0\n'
            content += '%s' %('FULLSTATS	0\n' if self.mami_flag else '')
            content += 'MEND\n\n'

            # variable
            content += self.write_pj_variable(value)

            # load case
            content += 'MSTART MULTISIM\n' \
                       'NCASES	%s\n' %len(self.lc_list)

            for lc in self.lc_list:

                content += 'DIRECTORY	%s\n' %self.lc_path[lc]
                content += 'RUNNAME	%s\n' %lc
            content += 'MEND\n\n'

            content += '0BSTAT\n'
            content += '0MULTISIM\n'
            content += '0MULTIVARIABLE\n'

            content += '		]]>\n' \
                       '	</BladedData>\n' \
                       '</BladedProject>\n'

            pj_name = key+'.$PJ'
            pj_file = self.post_path.replace(' ', '_')
            if not os.path.isdir(pj_file):
                os.makedirs(pj_file)

            pj_path = os.path.join(pj_file, pj_name)
            with open(pj_path, 'w+') as f:
                f.write(content)
            print('%s pj is done!' %key)

    def write_in(self):

        for key, value in self.chan_dict.items():
            # print(key, value)

            # variable
            pj_path = self.post_path.replace(' ', '_')

            # header
            content  = 'PTYPE	1\n' \
                       'SDSTAT	1\n' \
                       'DONGLOG	0\n' \
                       'OUTSTYLE	B\n'
            content += 'PATH	%s\n' %pj_path
            content += 'RUNNAME	%s\n' %key
            content += 'MSTART BSTAT\n' \
                       'TYPE	2\n'
            content += '%s' %('FULLSTATS	0\n' if self.mami_flag else '')
            content += 'MEND\n\n'

            # variable
            content += self.write_in_variable(value)

            # load case
            content += 'MSTART MULTIRUN\n' \
                       'NCASES	%s\n' %len(self.lc_list)
            content += 'OUTCHOICE	2\n'

            for lc in self.lc_list:

                content += 'DIRECTORY	%s\n' %(self.lc_path[lc]+'\\')
                content += 'RUNNAME	%s\n' %lc
                content += 'OUTEXTEN	 0\n'

            content += 'MEND\n\n\n'

            if not os.path.isdir(pj_path):
                os.makedirs(pj_path)

            with open(os.sep.join([pj_path, 'dtsignal.in']), 'w+') as f:
                f.write(content)
            print('%s in is done!' %key)

    def write_pj_variable(self, var_list):

        content  = 'MSTART MULTIVARIABLE\n'
        content += 'NVARS	%s\n' %len(var_list)
        for var in var_list:
            content += 'ATTRIBF	%%%s\n'  %var[1]
            content += 'VARIAB	"%s"\n'  %(var[0] if len(var)==2 else var[0]+', Distance along blade= %sm' %var[2])
            content += 'DESCRIPTION	"%s"\n' % var[0]
            content += 'NDIMENS	%s\n' %('2' if len(var)==2 else '3')
            content += 'DIMFLAG	-1\n'
            content += '%s' %('' if len(var)==2 else ('DIM2	%s\n' %var[2]))

        content += 'MEND\n\n'
        return content

    def write_in_variable(self, var_list):

        content  = 'MSTART MULTIVAR\n'
        content += 'NVARS	%s\n' % len(var_list)
        for var in var_list:
            content += 'ATTRIBF	%%%s\n'  %var[1]
            content += "VARIAB	'%s'\n"  %var[0]
            content += '%s' %('' if len(var)==2 else ('DIM2	 %s\n' %var[2]))
            content += "FULL_NAME	'%s'\n" %(var[0] if len(var)==2 else var[0]+', Distance along blade= %sm' %var[2])

        content += 'MEND\n\n'
        return content

if __name__ == '__main__':

    post_path = r'\\172.20.0.4\fs02\CE\V3\post_test'
    run_path  = r'\\172.20.0.4\fs02\CE\V3\loop05\run_0615'
    dlc_list  = ['DLC12']

    chan_dict = {'br':[['Rotor speed','05'],['Hub wind speed magnitude','07'],['Stationary hub Mx','23']]}

    Bstats(run_path, dlc_list, chan_dict, post_path)
