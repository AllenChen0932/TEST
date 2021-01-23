# -*- coding: utf-8 -*-
# @Time    : 2020/5/3 9:10
# @Author  : CE
# @File    : ultimate.py

import os

class Ultimate(object):

    def __init__(self,
                 run_path,
                 ult_path,
                 dlc_list,
                 channel_dict,
                 iec_stand,
                 etm_option,
                 mean_index='#',
                 half_index='+',
                 include_sf=True):

        self.run_path  = run_path.replace('/', '\\')
        self.post_path = ult_path.replace('/', '\\')
        if not os.path.exists(self.post_path):
            os.makedirs(self.post_path)
        self.dlc_list  = dlc_list
        self.iec_stand = iec_stand
        self.etm_option = etm_option
        self.cha_dict  = channel_dict    #{channel: variable list}, e.g. yb:[Yaw bearing Fx, Yaw bearing Fy]

        self.mean_index = mean_index
        self.half_index = half_index
        self.include_sf = include_sf

        self.dlc_number   = 0
        self.all_loadcase = []
        self.allcase_dlc  = {}
        self.all_lc_path  = {}

        self.allgroup_list  = []     # all groups
        self.subgroup_list  = []     # sub groups

        self.lc_group_name  = {}     # load case name: group name

        self.allgroup_index = {}     # 13_aa#: 0/1/2/3/4/...
        self.allgroup_types = {}     # 13_aa#: 0/1/2

        # group for determine safety factor
        self.allgroup_dlc   = {}     # 13_aa#: DLC13
        self.dlc_sf         = {}

        # self.finish_ult    = False
        # self.error_ult     = False

        if self.iec_stand == '1':
            self.dlc_sf = {'DLC11':1.25,
                           'DLC12':1.0,
                           'DLC13':1.35,
                           'DLC14':1.35,
                           'DLC15':1.35,
                           'DLC16':1.35,
                           'DLC21':1.35,
                           'DLC22':1.1,
                           'DLC23':1.1,
                           'DLC24':1.0,
                           'DLC31':1.0,
                           'DLC32':1.35,
                           'DLC33':1.35,
                           'DLC41':1.0,
                           'DLC42':1.35,
                           'DLC51':1.35,
                           'DLC61':1.35,
                           'DLC62':1.1,
                           'DLC63':1.35,
                           'DLC64':1.0,
                           'DLC71':1.1,
                           'DLC72':1.0,
                           'DLC81':1.5,
                           'DLC82':1.1,
                           'DLC83':1.0}
        elif self.iec_stand == '2':
            self.dlc_sf = {'DLC11': 1.25,
                           'DLC12': 1.0,
                           'DLC13': 1.35,
                           'DLC14': 1.35,
                           'DLC15': 1.35,
                           'DLC16': 1.35,
                           'DLC21': 1.35,
                           'DLC22': 1.1,
                           'DLC23': 1.1,
                           'DLC24': 1.0,
                           'DLC25': 1.2,
                           'DLC31': 1.0,
                           'DLC32': 1.35,
                           'DLC33': 1.35,
                           'DLC41': 1.0,
                           'DLC42': 1.35,
                           'DLC51': 1.35,
                           'DLC61': 1.35,
                           'DLC62': 1.1,
                           'DLC63': 1.35,
                           'DLC64': 1.0,
                           'DLC71': 1.1,
                           'DLC72': 1.0,
                           'DLC81': 1.35,
                           'DLC82': 1.1,
                           'DLC83': 1.0,
                           'DLC84': 1.0}

            # safety factor for DLC23 if ETM used
            if self.etm_option:
                self.dlc_sf['DLC23'] = 1.35
        elif self.iec_stand == '3':
            self.read_sf()

        self.file_name ={'br':'Blade_root',
                         'hs':'Hub_stationary',
                         'hr':'Hub_rotating',
                         'tr':'Tower',
                         'yb':'Yaw_bearing',
                         'bps':'Blade_principal',
                         'brs':'Blade_root_axes',
                         'bas':'Blade_aerodynamic',
                         'bus':'Blade_user_axes',
                         'ep':'Extrapolation',
                         'tc':'Tower_clearance'}

        try:
            self.get_subgroup()
            self.write_pj()
            self.write_in()
            self.finish_ult = True
        except:
            self.finish_ult = False
            # pass
            # self.error_flag = True

    def read_sf(self):
        '''
        read safety factor
        :return:
        '''

        sf_file = 'dlc_sf.txt'
        file_path = os.path.join(os.getcwd(), sf_file)

        if os.path.isfile(file_path):

            with open(file_path, 'r') as f:

                lines = f.readlines()
                for line in lines:

                    if line:
                        list = line.strip().split(':')
                        # print(list)
                        self.dlc_sf[list[0]] = list[1]
                    else:
                        continue

    def get_subgroup(self):
        '''
        get subgroup sorted by mean and half
        :return:
        '''

        group_index = 0
        self.dlc_list.sort()
        # print(self.dlc_list)

        for dlc in self.dlc_list:

            dlc_path = os.path.join(self.run_path, dlc)
            lc_list  = os.listdir(dlc_path)

            for lc in lc_list:

                lc_path = os.path.join(dlc_path, lc)

                if os.path.isdir(lc_path):

                    self.dlc_number += 1
                    self.all_loadcase.append(lc)
                    self.allcase_dlc[lc] = dlc[:5]
                    self.all_lc_path[lc] = lc_path

                    if self.mean_index in lc:
                        group_name = str(lc.split(self.mean_index)[0])+self.mean_index

                        self.lc_group_name[lc] = group_name

                        # if group_name not in self.allgroup_dlc.keys():
                        #     self.allgroup_dlc.append(group_name)

                        if group_name not in self.allgroup_index.keys():
                            group_index +=1
                            self.allgroup_index[group_name] = group_index

                        if group_name not in self.subgroup_list:
                            self.subgroup_list.append(group_name)
                            self.allgroup_types[group_name] = str(1)
                            self.allgroup_dlc[group_name]  = dlc[:5]

                    elif self.half_index in lc:
                        group_name = str(lc.split(self.half_index)[0])+self.half_index

                        self.lc_group_name[lc] = group_name
                        # if group_name not in self.allgroup_list:
                        #     self.allgroup_list.append(group_name)
                        #
                        if group_name not in self.allgroup_index.keys():
                            group_index +=1
                            self.allgroup_index[group_name] = group_index

                        if group_name not in self.subgroup_list:
                            self.subgroup_list.append(group_name)
                            self.allgroup_types[group_name] = str(2)
                            self.allgroup_dlc[group_name]  = dlc[:5]

                    elif '-' in lc:
                        group_name = str(lc.split('-')[0]) if lc.startswith('14') else (str(lc.split('-')[0])+'-')

                        self.lc_group_name[lc] = group_name

                        if group_name not in self.allgroup_dlc:
                            self.allgroup_dlc[group_name] = dlc[:5]

                        if group_name not in self.allgroup_index.keys():
                            self.allgroup_index[group_name] = str(0)

        self.allgroup_list = list(self.allgroup_dlc.keys())
        self.allgroup_list.sort()

        # print(self.allgroup_dlc)
        # print(self.allgroup_index)

    def write_pj(self):

        for key, value in self.cha_dict.items():

            if len(value) == 2:
                # write bpa, bra, baa, bua and tower
                print(key, value)
                for sec in value[1]:

                    # header
                    content = '<?xml version="1.0" encoding="ISO-8859-1" ?>\n' \
                              '<BladedProject ApplicationVersion="4.8.0.41">\n' \
                              '\t<DataModelVersion>0.8</DataModelVersion>\n' \
                              '\t<BladedData dataFormat="project">\n' \
                              '<![CDATA[\n' \
                              'VERSION	4.8.0.34\n' \
                              'MULTIBODY	 1\n' \
                              'CALCULATION	29\n' \
                              'OPTIONS	0\n' \
                              'PROJNAME	\n' \
                              'DATE	\n' \
                              'ENGINEER	\n' \
                              'NOTES	""\n' \
                              'PASSWORD	\n' \
                              'MSTART MAXMIN\n' \
                              'HISTTYPE	1\n'

                    # setting
                    content += 'NCASES\t%s\n' %self.dlc_number
                    content += 'USEGROUPS	1\n'
                    content += 'USESUBGROUPS	1\n'
                    content += 'NUMSUBGROUP	%s\n' %len(self.subgroup_list)
                    content += 'SUBGROUPTYPE %s\n' %' '.join([self.allgroup_types[lc] for lc in self.subgroup_list])
                    content += 'SUBGROUPTEXT	%s\n' %self.mean_index
                    content += 'SUBGROUPTEXT	%s\n' %self.half_index

                    # load case
                    for lc in self.all_loadcase:
                        content += 'CASE	%s\n' %lc
                        content += 'DIRECTORY	%s\n' %self.all_lc_path[lc]
                        content += 'OLDVERSION	0\n'
                        content += 'RUNNAME	%s\n' %lc
                        content += 'GROUP	%s\n' %self.lc_group_name[lc]
                        content += 'SUBGROUP	%s\n' %self.allgroup_index[self.lc_group_name[lc]]

                    # variable
                    content += self.write_pj_variable(key, value[0], sec)

                    # safety factor
                    content += 'NGROUPS	%s\n' %len(self.allgroup_list)

                    for gp in self.allgroup_list:

                        content += 'NAME	"%s"\n' %gp
                        content += 'TYPE	0\n'
                        content += 'SF	 %s\n' %(self.dlc_sf[self.allgroup_dlc[gp]] if self.include_sf else 1.0)

                    content += 'OUTPUTSF	1\n'
                    content += 'SFOVERRIDE	0\n'
                    content += 'MEND\n'
                    # print(content)
                    pj_name = sec if 'Mbr' in sec else ((key+'_'+str('%.3f' %float(sec))) if key != 'ep' else key)
                    print(pj_name)
                    content += '\n' \
                               '0MAXMIN\n' \
                               '		]]>\n' \
                               '	</BladedData>\n' \
                               '<RunConfiguration>\n' \
                               '  <Name>%s</Name>\n' \
                               '  <Calculation>0</Calculation>\n' \
                               '</RunConfiguration>\n\n' \
                               '</BladedProject>' %(pj_name if 'Mbr' not in sec else ("'%s'" %sec))

                    pj_file = pj_name +'.$PJ'
                    # print(pj_file)
                    # file_name = '_'.join([self.file_name[key], str(sec)]) if key != 'tr' \
                    #     else ('_'.join([self.file_name[key], str(sec)]) if 'Mbr' not in key else key)
                    pj_path = os.path.join(self.post_path, pj_name).replace(' ', '_')
                    print(pj_path)
                    if not os.path.isdir(pj_path):
                        os.makedirs(pj_path)

                    with open(os.sep.join([pj_path, pj_file]), 'w+') as f:
                        f.write(content)
                    print('%s pj is done!' %key)

            else:
                # yb, hs, hr
                print(key, value)
                # header
                content = '<?xml version="1.0" encoding="ISO-8859-1" ?>\n' \
                          '<BladedProject ApplicationVersion="4.8.0.41">\n' \
                          '\t<DataModelVersion>0.8</DataModelVersion>\n' \
                          '\t<BladedData dataFormat="project">\n' \
                          '<![CDATA[\n' \
                          'VERSION	4.8.0.34\n' \
                          'MULTIBODY	 1\n' \
                          'CALCULATION	29\n' \
                          'OPTIONS	0\n' \
                          'PROJNAME	\n' \
                          'DATE	\n' \
                          'ENGINEER	\n' \
                          'NOTES	""\n' \
                          'PASSWORD	\n' \
                          'MSTART MAXMIN\n' \
                          'HISTTYPE	1\n'

                # setting
                content += 'NCASES\t%s\n' %self.dlc_number
                content += 'USEGROUPS	1\n'
                content += 'USESUBGROUPS	1\n'
                content += 'NUMSUBGROUP	%s\n' %len(self.subgroup_list)
                content += 'SUBGROUPTYPE %s\n' %' '.join([self.allgroup_types[lc] for lc in self.subgroup_list])
                content += 'SUBGROUPTEXT	%s\n' %self.mean_index
                content += 'SUBGROUPTEXT	%s\n' %self.half_index

                # load case
                for lc in self.all_loadcase:
                    content += 'CASE	%s\n' %lc
                    content += 'DIRECTORY	%s\n' %self.all_lc_path[lc]
                    content += 'OLDVERSION	0\n'
                    content += 'RUNNAME	%s\n' %lc
                    content += 'GROUP	%s\n' %self.lc_group_name[lc]
                    content += 'SUBGROUP	%s\n' %self.allgroup_index[self.lc_group_name[lc]]

                # variable
                content += self.write_pj_variable(key, value)

                # safety factor
                content += 'NGROUPS	%s\n' %len(self.allgroup_list)
                for gp in self.allgroup_list:
                    content += 'NAME	"%s"\n' %gp
                    content += 'TYPE	0\n'
                    content += 'SF	 %s\n' %(self.dlc_sf[self.allgroup_dlc[gp]] if self.include_sf else 1.0)

                content += 'OUTPUTSF	1\n'
                content += 'SFOVERRIDE	0\n'
                content += 'MEND\n'

                content += '\n' \
                           '0MAXMIN\n' \
                           '		]]>\n' \
                           '	</BladedData>\n' \
                           '<RunConfiguration>\n' \
                           '  <Name>%s</Name>\n' \
                           '  <Calculation>0</Calculation>\n' \
                           '</RunConfiguration>\n\n' \
                           '</BladedProject>' %key
                # print(content)

                pj_file = key+'.$PJ'
                # print(pj_file)
                # file_name = self.file_name[key]
                # pj_path = os.path.join(self.post_path, file_name).replace(' ', '_')
                pj_path = os.path.join(self.post_path, key).replace(' ', '_')
                print(pj_path)
                if not os.path.isdir(pj_path):
                    os.makedirs(pj_path)

                with open(os.sep.join([pj_path,pj_file]), 'w+') as f:
                    f.write(content)
                print('%s pj is done!' %key)

    def write_in(self):

        for key, value in self.cha_dict.items():
            print(key, value)
            if len(value) == 2:

                for sec in value[1]:

                    pj_name = sec if 'Mbr' in sec else ((key + '_' + str('%.3f' %float(sec))) if key != 'ep' else key)
                    pj_path = os.path.join(self.post_path, pj_name).replace(' ', '_')

                    content  = 'PTYPE	14\n' \
                               'SDSTAT	1\n' \
                               'OUTSTYLE	B\n'
                    content += 'PATH\t%s\n' %pj_path
                    content += 'RUNNAME	%s\n' %(pj_name if 'Mbr' not in sec else ("'%s'" %sec))
                    content += 'MSTART MAXMIN\n' \
                               'HISTTYPE	1\n'
                    content += 'NCASES\t%s\n' %self.dlc_number

                    if len(self.subgroup_list):
                        content += 'NUMSUBGROUP	%s\n' %len(self.subgroup_list)
                        content += 'SUBGROUPTYPE %s\n' %' '.join([self.allgroup_types[lc] for lc in self.subgroup_list])

                    # load case
                    for lc in self.all_loadcase:
                        content += 'CASE	%s\n' %lc
                        content += 'DIRECTORY	%s\n' %(self.all_lc_path[lc]+os.sep)
                        content += 'RUNNAME	%s\n' %lc
                        content += 'GROUP	%s\n' %self.lc_group_name[lc]
                        sg_index = self.allgroup_index[self.lc_group_name[lc]]
                        content += '' if sg_index == '0' else ('SUBGROUP	%s\n' %sg_index)

                    # write variable
                    result = self.write_in_variable(key, value[0], sec)
                    print(len(result))
                    if len(result) == 2:
                        content += result[0]
                    else:
                        content += result

                    # write safety factor
                    content += 'SFMATRIX	'
                    var_num = len(value[0])

                    for j in range(var_num):
                        i = 0

                        for ind, lc in enumerate(self.all_loadcase):

                            if i == 500:
                                content += '\n'
                                content += '%s,' %(self.dlc_sf[self.allcase_dlc[lc]] if self.include_sf else '1.0')
                                i = 1

                            else:
                                content += '%s,' %(self.dlc_sf[self.allcase_dlc[lc]] if self.include_sf else '1.0')
                                i += 1

                        # the last safety factor without comma
                        if j == var_num - 1:
                            content  = content[:-1]
                            content += '\n'
                        else:
                            content += '\n'

                    content += 'OUTPUTSF 1\n'

                    if len(result) == 2:
                        content += result[1]

                    content += 'MEND'

                    if not os.path.isdir(pj_path):
                        os.makedirs(pj_path)

                    with open(os.sep.join([pj_path,'dtsignal.in']), 'w+') as f:
                        f.write(content)
                    print('%s in is done!' %key)

            else:

                pj_path = os.path.join(self.post_path, key).replace(' ', '_')

                content  = 'PTYPE	14\n' \
                           'SDSTAT	1\n' \
                           'OUTSTYLE	B\n'
                content += 'PATH %s\n' %pj_path
                content += 'RUNNAME	%s\n' %key
                content += 'MSTART MAXMIN\n' \
                           'HISTTYPE	1\n'
                content += 'NCASES\t%s\n' %self.dlc_number
                content += 'NUMSUBGROUP	%s\n' %len(self.subgroup_list)
                content += 'SUBGROUPTYPE %s\n' %' '.join([self.allgroup_types[lc] for lc in self.subgroup_list])

                # load case
                for lc in self.all_loadcase:
                    content += 'CASE	%s\n' %lc
                    content += 'DIRECTORY	%s\n' %(self.all_lc_path[lc]+os.sep)
                    content += 'RUNNAME	%s\n' %lc
                    content += 'GROUP	%s\n' %self.lc_group_name[lc]
                    sg_index = self.allgroup_index[self.lc_group_name[lc]]
                    content += '' if sg_index == '0' else ('SUBGROUP	%s\n' %sg_index)

                # write variable
                result   = self.write_in_variable(key, value)
                if len(result) == 2:
                    content += result[0]
                else:
                    content += result

                # write safety factor
                content += 'SFMATRIX	'
                var_num = len(value)

                for j in range(var_num):
                    i = 0

                    for ind, lc in enumerate(self.all_loadcase):

                        if i == 500:
                            content += '\n'
                            content += '%s,' %(self.dlc_sf[self.allcase_dlc[lc]] if self.include_sf else '1.0')
                            i = 1

                        else:
                            content += '%s,' %(self.dlc_sf[self.allcase_dlc[lc]] if self.include_sf else '1.0')
                            i += 1

                    # the last safety factor without comma
                    if j == var_num-1:
                        content  = content[:-1]
                        content += '\n'
                    else:
                        content += '\n'

                content += 'OUTPUTSF 1\n'

                if len(result) == 2:
                    content += result[1]

                content += 'MEND'

                if not os.path.isdir(pj_path):
                    os.makedirs(pj_path)

                with open(os.sep.join([pj_path,'dtsignal.in']), 'w+') as f:
                    f.write(content)
                print('%s in is done!' %key)

    def write_pj_variable(self, channel, var_list, height=None):

        content = ''

        if channel == 'br':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'

            for var in var_list:
                content += 'ATTRIBF	%22\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s"\n' %var
                content += 'NDIMENS	2\n'
                content += 'DIMFLAG	-1\n'
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
                content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'hr':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'

            for var in var_list:
                content += 'ATTRIBF	%22\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s"\n' %var
                content += 'NDIMENS	2\n'
                content += 'DIMFLAG	-1\n'
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'hs':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'
            for var in var_list:

                content += 'ATTRIBF	%23\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s"\n' %var
                content += 'NDIMENS	2\n'
                content += 'DIMFLAG	-1\n'
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'yb':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'
            for var in var_list:

                content += 'ATTRIBF	%24\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s"\n' %var
                content += 'NDIMENS	2\n'
                content += 'DIMFLAG	-1\n'
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'ep':

            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            ext_list = ['22']*6+['18','19','20']
            for index, var in enumerate(var_list):

                if index<6:
                    content += 'ATTRIBF	%'+'%s\n' %ext_list[index]
                    content += 'VARIAB	"%s"\n' %var
                    content += 'DESCRIPTION	"%s"\n' %var
                    content += 'NDIMENS	2\n'
                    content += 'DIMFLAG	-1\n'
                    content += 'SFDEFAULT	-1\n'
                    content += 'SFVALUE	 0\n'
                else:
                    content += 'ATTRIBF	%' + '%s\n' %ext_list[index]
                    content += 'VARIAB	"%s"\n' %var
                    content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %(var, height)
                    content += 'NDIMENS	3\n'
                    content += 'DIMFLAG	-1\n'
                    content += 'DIM2	 %s\n' %height
                    content += 'SFDEFAULT	-1\n'
                    content += 'SFVALUE	 0\n'
                    content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'tc':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'
            for var in var_list:

                content += 'ATTRIBF	%07\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s"\n' %var
                content += 'NDIMENS	2\n'
                content += 'DIMFLAG	-1\n'
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'tr' and 'Mbr' in height:
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'
            for var in var_list:

                content += 'ATTRIBF	%25\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, %s"\n' %(var, height)
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	0\n'
                content += "AXITICK2	'%s'\n" %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'tr' and 'Mbr' not in height:
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	0\n'
            for var in var_list:

                content += 'ATTRIBF	%25\n'
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, Tower station height= %sm"\n' %(var, float(height))
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	-1\n'
                content += 'DIM2	 %s\n' %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'bps':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index,var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(15+index%3)
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %(var, height)
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	-1\n'
                content += 'DIM2	 %s\n' %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
                content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'brs':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(41+index%3)
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %(var, height)
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	-1\n'
                content += 'DIM2	 %s\n' %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
                content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'bas':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(59+index%3)
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %(var, height)
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	-1\n'
                content += 'DIM2	 %s\n' %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
                content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'

        elif channel == 'bus':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'BLADEFLAG	-1\n'
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(62+index%3)
                content += 'VARIAB	"%s"\n' %var
                content += 'DESCRIPTION	"%s, Distance along blade= %sm"\n' %(var, height)
                content += 'NDIMENS	3\n'
                content += 'DIMFLAG	-1\n'
                content += 'DIM2	 %s\n' %height
                content += 'SFDEFAULT	-1\n'
                content += 'SFVALUE	 0\n'
                content += 'BLCHAN	0\n'
            content += 'EXTRAVARS\t0\n'
        # print(content)
        return content

    def write_in_variable(self, channel, var_list, height=None):
        # print(channel)
        content  = ''
        content1 = ''

        if channel == 'br':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'

            for var in var_list:
                content += 'ATTRIBF	%22\n'
                content += "VARIAB	'%s'\n" %var
                content += "FULL_NAME	'%s'\n" %var

            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade root Mx'\n"
            content1 += "GRPNAME\t'Blade root My'\n"
            content1 += "GRPNAME\t'Blade root Mxy'\n"
            content1 += "GRPNAME\t'Blade root Mz'\n"
            content1 += "GRPNAME\t'Blade root Fx'\n"
            content1 += "GRPNAME\t'Blade root Fy'\n"
            content1 += "GRPNAME\t'Blade root Fxy'\n"
            content1 += "GRPNAME\t'Blade root Fz'\n"

            return content, content1

        elif channel == 'hr':
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:
                content += 'ATTRIBF	%22\n'
                content += "VARIAB	'%s'\n" %var
                content += "FULL_NAME	'%s'\n" %var

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'hs':
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:
                content += 'ATTRIBF	%23\n'
                content += "VARIAB	'%s'\n" %var
                content += "FULL_NAME	'%s'\n" %var

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'yb':
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:

                content += 'ATTRIBF	%24\n'
                content += "VARIAB	'%s'\n" %var
                content += "FULL_NAME	'%s'\n" %var

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'tc':
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:

                content += 'ATTRIBF	%07\n'
                content += "VARIAB	'%s'\n" %var
                content += "FULL_NAME	'%s'\n" %var

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'ep':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            ext_list = ['22']*6 + ['18', '19', '20']
            for index, var in enumerate(var_list):

                if index < 6:
                    content += 'ATTRIBF	%'+'%s\n' %ext_list[index]
                    content += "VARIAB	'%s'\n" %var
                    content += "Full_NAME	'%s'\n" %var
                else:
                    content += 'ATTRIBF	%'+'%s\n' %ext_list[index]
                    content += "VARIAB	'%s'\n" %var
                    content += 'Full_NAME	"%s, Distance along blade= %sm"\n' %(var, height)
            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade root Mx'\n"
            content1 += "GRPNAME\t'Blade root My'\n"
            content1 += "GRPNAME\t'Blade x-deflection (perpendicular to rotor plane), Distance along blade= %sm'\n" %(height)
            return content, content1

        elif channel == 'tr' and 'Mbr' in height:
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:

                content += 'ATTRIBF	%25\n'
                content += "VARIAB	'%s'\n" %var
                content += "AXITICK2	'%s'\n" %height
                content += "FULL_NAME	'%s, %s'\n" %(var, height)

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'tr' and 'Mbr' not in height:
            content += 'NVARS	%s\n' %len(var_list)

            for var in var_list:

                content += 'ATTRIBF	%25\n'
                content += "VARIAB	'%s'\n" %var
                content += 'DIM2	 %s\n' %height
                content += "FULL_NAME	'%s, Tower station height= %sm'\n" %(var, float(height))

            content += 'EXTRAVARS\t0\n'
            return content

        elif channel == 'bps':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index,var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(15+index%3)
                content += "VARIAB	'%s'\n" %var
                content += 'DIM2	 %s\n' %height
                content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %(var, height)

            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade Mx (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade My (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mxy (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mz (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fx (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fy (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fxy (Principal axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fz (Principal axes), Distance along blade= %sm'\n" %(height)

            return content, content1

        elif channel == 'brs':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(41+index%3)
                content += "VARIAB	'%s'\n" %var
                content += 'DIM2	 %s\n' %height
                content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %(var, height)

            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade Mx (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade My (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mxy (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mz (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fx (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fy (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fxy (Root axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fz (Root axes), Distance along blade= %sm'\n" %(height)

            return content, content1

        elif channel == 'bas':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(59+index%3)
                content += "VARIAB	'%s'\n" %var
                content += 'DIM2	 %s\n' %height
                content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %(var, height)

            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade Mx (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade My (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mxy (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mz (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fx (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fy (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fxy (Aerodynamic axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fz (Aerodynamic axes), Distance along blade= %sm'\n" %(height)

            return content, content1

        elif channel == 'bus':
            content += 'NVARS	%s\n' %len(var_list)
            content += 'COMBINE	3\n'
            content += 'COMBINEMETHOD	0\n'
            for index, var in enumerate(var_list):

                content += 'ATTRIBF	%'+'%s\n' %int(62+index%3)
                content += "VARIAB	'%s'\n" %var
                content += 'DIM2	 %s\n' %height
                content += "FULL_NAME	'%s, Distance along blade= %sm'\n" %(var, height)

            content += 'EXTRAVARS\t0\n'

            content1 += "GRPNAME\t'Blade Mx (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade My (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mxy (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Mz (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fx (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fy (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fxy (User axes), Distance along blade= %sm'\n" %(height)
            content1 += "GRPNAME\t'Blade Fz (User axes), Distance along blade= %sm'\n" %(height)

            return content, content1

        return content

if __name__ == '__main__':

    run_path = r'\\172.20.0.4\fs03\CE\W6250-172-14m\run_0430_0.8'

    # ultimate(run_path)

