#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 3/10/2020 2:37 PM
# @Author  : CE
# @File    : tower_chan_comb.py

import os
import numpy as np
# import pysnooper

class Combination(object):

    def __init__(self, run_path, res_path, fat_list, chan_dict, file_extension, extension, angle_list):

        # self.prj_name = prj_name
        self.run_path = run_path
        self.res_path = res_path
        self.fat_list = fat_list
        self.chan_dict = chan_dict
        self.file_ext  = file_extension
        self.extension = extension
        self.ang_list = angle_list

        # self.ang_list = list(range(0,360,15))

        if not os.path.exists(self.res_path):
            os.makedirs(self.res_path)

        self.lc_list = []

        self.get_loadcase()
        self.write_content()

    def get_loadcase(self):

        for lc in self.fat_list:
            # eg: run/DLC12
            lc_path = os.path.join(self.run_path, lc)
            files = os.listdir(lc_path)

            for file in files:
                if os.path.isdir(os.path.join(lc_path, file)):
                    self.lc_list.append(os.path.join(lc_path, file))
    # @pysnooper.snoop()
    def pj_content(self, lc_list, channel, variable_list):

        # write header
        content  = '<?xml version="1.0" encoding="ISO-8859-1" ?>\n'
        content += '<BladedProject ApplicationVersion="4.8.0.41">\n'
        content += '	<DataModelVersion>0.8</DataModelVersion>\n'
        content += '	<BladedData dataFormat="project">\n'

        # write data
        content += '<![CDATA[\n'
        content += 'VERSION	4.8.0.34\n'
        content += 'MULTIBODY	 1\n'
        content += 'CALCULATION	28\n'
        content += 'OPTIONS	0\n'
        content += 'PROJNAME	\n'
        content += 'DATE	\n'
        content += 'ENGINEER	\n'
        content += 'NOTES	""\n'
        content += 'PASSWORD	\n'

        # write content
        content += 'MSTART CHAN\n'
        content += 'TYPE	3\n'
        content += 'NTABLES	0\n'
        content += 'MATFLAG	0\n'
        content += 'NMATRIX	0\n'
        content += 'VECTFLAG	0\n'
        content += 'NVECT	0\n'
        content += 'NVECTVARS	 0\n'

        # write_exptession
        content += 'NCHANS	%s\n' %(len(self.ang_list))

        for ind,val in enumerate(self.ang_list):

            cosx = np.cos(np.deg2rad(float(val)))
            sinx = np.sin(np.deg2rad(float(val)))

            content += "EXPRESSION	'( $1 * %.6f ) + ( $2 * %.6f )'\n" %(cosx, sinx)
            content += "DISPNAME	''\n"
            content += "VARIAB	'%s_mxy at %s deg'\n" %(channel, val)
            content += "UNITS	FL\n"

        content += "GENLAB	'%s_mxy_angle'\n" %channel
        content += "OUTEXTEN	 %s\n" %self.extension
        content += "NVARS	 2\n"

        # variable
        if len(variable_list[0])==1:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += 'VARIAB	"%s"\n' %variable_list[0]
            content += 'DESCRIPTION	"%s"\n' %variable_list[0]
            content += 'DISPNAME	""\n'
            content += 'NDIMENS	2\n'
            content += 'DIMFLAG	-1\n'
            content += 'UNITS	"Nm"\n'
        else:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += 'VARIAB	"%s"\n' %variable_list[0][0]
            content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %tuple(variable_list[0][:2])
            content += 'DISPNAME	""\n'
            content += 'NDIMENS	3\n'
            content += 'DIMFLAG	-1\n'
            content += 'DIM2	 %s\n' %variable_list[0][2]
            content += 'UNITS	"Nm"\n'
        if len(variable_list[1])==1:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += 'VARIAB	"%s"\n' %variable_list[1]
            content += 'DESCRIPTION	"%s"\n' %variable_list[1]
            content += 'DISPNAME	""\n'
            content += 'NDIMENS	2\n'
            content += 'DIMFLAG	-1\n'
            content += 'UNITS	"Nm"\n'
        else:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += 'VARIAB	"%s"\n' %variable_list[1][0]
            content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %tuple(variable_list[1][:2])
            content += 'DISPNAME	""\n'
            content += 'NDIMENS	3\n'
            content += 'DIMFLAG	-1\n'
            content += 'DIM2	 %s\n' %variable_list[1][2]
            content += 'UNITS	"Nm"\n'
        content += 'MEND\n'
        content += '\n'

        # write load case
        content += 'MSTART MULTISIM\n'
        content += 'NCASES	%s\n' %(len(lc_list))

        for lc in lc_list:
            content += ('DIRECTORY	%s'+'\\'+'\n') %lc
            lc_name = lc.split('\\')[-1]
            content += 'RUNNAME	%s\n' %lc_name
        content += 'MEND\n'

        # write end
        content += '\n'
        content += '0CHAN\n'
        content += '0MULTISIM\n'
        content += '		]]>\n'
        content += '	</BladedData>\n'
        content += '</BladedProject>'

        return content

    def in_content(self, lc_list, post_path, channel, variable_list):

        # write header
        content  = 'PTYPE	13\n' \
                   'SDSTAT	1\n' \
                   'DONGLOG	0\n' \
                   'OUTSTYLE	B\n' \
                   'PATH	%s\n' \
                   'RUNNAME	%s\n' \
                   'MSTART CHAN\n' \
                   'TYPE	3\n' %(post_path, channel)

        # write_exptession
        content += 'NCHANS	%s\n' %(len(self.ang_list))

        for ind,val in enumerate(self.ang_list):

            cosx = np.cos(np.deg2rad(float(val)))
            sinx = np.sin(np.deg2rad(float(val)))

            content += "EXPRESSION	'( $1 * %.6f ) + ( $2 * %.6f )'\n" %(cosx, sinx)
            content += "OUTPUT   1\n"
            content += "VARIAB	'%s_mxy at %s deg'\n" %(channel,val)
            content += "UNITS	FL\n"

        content += "GENLAB	'%s_mxy_angle'\n" %channel
        content += 'MEND\n'

        # write load case
        content += 'MSTART MULTIRUN\n'
        content += 'NCASES	%s\n' %(len(lc_list))

        for lc in lc_list:
            content += ('DIRECTORY	%s'+'\\'+'\n') %lc

            lc_name = lc.split('\\')[-1]
            content += 'RUNNAME	%s\n' %lc_name
            content += "OUTEXTEN	 %s\n" %self.extension
        content += 'MEND\n'

        # variable
        content += "MSTART MULTIVAR\n"
        content += 'NVARS	2\n'
        if len(variable_list[0])==1:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += "VARIAB	'%s'\n" %variable_list[0]
            content += "FULL_NAME	'%s'\n" %variable_list[0]
        else:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += "VARIAB	'%s'\n" %variable_list[0][0]
            content += 'DIM2	 %s\n' %variable_list[0][2]
            content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %tuple(variable_list[0][:2])
        if len(variable_list[1])==1:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += "VARIAB	'%s'\n" %variable_list[1]
            content += "FULL_NAME	'%s'\n" %variable_list[1]
        else:
            content += "ATTRIBF	%%%s\n" %self.file_ext
            content += "VARIAB	'%s'\n" %variable_list[1][0]
            content += 'DIM2	 %s\n' %variable_list[1][2]
            content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %tuple(variable_list[1][:2])
        content += 'MEND\n'

        return content

    def write_content(self):

        for key, values in self.chan_dict.items():
            chan_path = os.path.join(self.res_path, key)
            lc_len = len(self.lc_list)
            jb_num = int(np.ceil(lc_len/100))

            for i in range(jb_num):
                temp = self.lc_list[i*100:(i+1)*100]

                file_name = 'combine'+str(i+1)
                file_path = os.path.join(chan_path, file_name)

                if not os.path.exists(file_path):
                    os.makedirs(file_path)

                prj_path  = os.path.join(file_path, '%s.$PJ' %('combine'+'_'+str(i+1)))
                in_path   = os.path.join(file_path, 'dtsignal.in')

                content_pj = self.pj_content(temp, key, values)
                content_in = self.in_content(temp, file_path, key, values)

                with open(prj_path, 'w+') as f1, open(in_path, 'w+') as f2:

                    f1.write(content_pj)
                    f2.write(content_in)

if __name__ == '__main__':

    run_path = r'\\172.20.4.132\fs02\CE\WE3600NB-167\run_0119'
    res_path = r'\\172.20.4.132\fs02\CE\WE3600NB-167\post\0119\04_Combination'
    fat_list = ['DLC12', 'DLC24a', 'DLC24b', 'DLC24c', 'DLC24d', 'DLC31', 'DLC41', 'DLC64']
    chan_dict = {'brs1': [('Blade 1 Mx (Root axes)', '0', '0'), ('Blade 1 My (Root axes)', '0', '0')]}

    Combination(run_path, res_path, fat_list, chan_dict, '41', '800', list(range(0, 360, 15)))
