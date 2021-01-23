#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/1/2020 8:04 PM
# @Author  : CE
# @File    : get_variable.py

import os
import itertools

class Get_Variable(object):

    def __init__(self, project):

        self.prj_path = project

        # info stored in project file
        self.twr_height    = []
        self.bld_radius    = []
        self.twr_output    = []
        self.bpa_output    = []    # blade principle axis

        self.bpa_out_type  = None
        self.bra_out_type  = None  # blade root axis
        self.baa_out_type  = None  # blade aerodynamic axis
        self.bua_out_type  = None  # blade user axis

        self.twr_model     = None

        # variable stored in % file
        self.br_var_list   = []
        self.yb_var_list   = []
        self.hr_var_list   = []
        self.hs_var_list   = []
        self.tr_var_list   = []
        self.ep_var_list   = []
        self.tc_var_list   = []
        self.bps_var_list  = []
        self.brs_var_list  = []
        self.bas_var_list  = []
        self.bus_var_list  = []
        
        # fatigue variable
        self.br1_fat_list  = []
        self.br2_fat_list  = []
        self.br3_fat_list  = []
        self.hr_fat_list   = []
        self.hs_fat_list   = []
        self.tr_fat_list   = []
        self.bps1_fat_list = []
        self.bps2_fat_list = []
        self.bps3_fat_list = []
        self.brs1_fat_list = []
        self.brs2_fat_list = []
        self.brs3_fat_list = []
        self.bas1_fat_list = []
        self.bas2_fat_list = []
        self.bas3_fat_list = []
        self.bus1_fat_list = []
        self.bus2_fat_list = []
        self.bus3_fat_list = [] 

        self.get_variable()

    def get_variable(self):

        pj_path, pj_file = os.path.split(self.prj_path)
        # print(pj_path, pj_file)

        with open(self.prj_path) as f:
            lines = f.readlines()
            for line in lines:

                if line.startswith('DIST') and 'DISTMODE' not in line:
                    self.bld_radius = line.strip().replace(' ', '').split('\t',1)[1].split(',')
                    self.bld_radius = [k for k, g in itertools.groupby(self.bld_radius)]
                    # print(len(self.bld_radius))
                    print('blade radius:', self.bld_radius)

                if line.startswith('TMODEL'):
                    self.twr_model = line.strip().split()[1]
                    print('tower model:', self.twr_model)

                if line.startswith('TJ'):
                    self.twr_height = line.strip().replace(' ','').split('\t',1)[1].split(',')
                    # print(len(self.twr_height))
                    print('tower height:', self.twr_height)

                if line.startswith('NTE') and self.twr_model == '3':
                    self.twr_height = [str(i) for i in range(1, int(line.split()[-1])+1)]
                    print('tower height:', self.twr_height)

                if line.startswith('BLOADS_STS'):
                    self.bpa_output = line.strip().split('\t',1)[1].split(',')
                    print('principal axis:', self.bpa_output)

                if line.startswith('BMFLAP'):
                    self.bpa_out_type = line.strip().split('\t')[1]
                    print('principal output:', self.bpa_out_type)

                if line.startswith('UNTWIST'):
                    self.bra_out_type = line.strip().split('\t')[1]
                    print('root axes output:', self.bra_out_type)

                if line.startswith('BL_AERO'):
                    self.baa_out_type = line.strip().split('\t')[1]
                    print('aerodynamic output:', self.baa_out_type)

                if line.startswith('BL_USER'):
                    self.bua_out_type = line.strip().split('\t')[1]
                    print('user output:', self.bua_out_type)

                # if self.twr_model == 2:
                if line.startswith('TLOADS_STS'):
                    self.twr_output = line.strip().split('\t',1)[1].split(',')
                    print('tower output:', self.twr_output)

                # elif self.twr_model == 3:
                if line.startswith('TLOADS_ELS'):
                    self.twr_output = line.strip().split()[1:]
                    print('tower output:', self.twr_output)

        pj_name = pj_file.split('.$')[0]

        # blade root
        br_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))
        self.br_flag = True if os.path.isfile(br_file) else False
        if self.br_flag:
            with open(br_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[8:]
                self.br_var_list = [var_list[8*j+i] for i in range(8) for j in range(3)]
            # print(self.br_var_list)
            
            self.br_fat_list = [var for var in self.br_var_list if '1' in var and 'xy' not in var]

        # blade root axis
        bs1_file = os.path.join(pj_path, '.%'.join([pj_name,'41']))
        bs2_file = os.path.join(pj_path, '.%'.join([pj_name,'42']))
        bs3_file = os.path.join(pj_path, '.%'.join([pj_name,'43']))
        self.bs_flag = True if os.path.isfile(bs1_file) and os.path.isfile(bs2_file) and os.path.isfile(bs3_file) else False
        if self.bs_flag:
            with open(bs1_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

            with open(bs2_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            with open(bs3_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            self.brs_var_list= [var_list[8*j+i] for i in range(8) for j in range(3)]
            # print(self.brs_var_list)

            self.brs_fat_list = [[var for var in self.brs_var_list if '1' in var and 'xy' not in var],
                                 [var for var in self.brs_var_list if '2' in var and 'xy' not in var],
                                 [var for var in self.brs_var_list if '3' in var and 'xy' not in var]]

        # blade principle axis
        bp1_file = os.path.join(pj_path, '.%'.join([pj_name,'15']))
        bp2_file = os.path.join(pj_path, '.%'.join([pj_name,'16']))
        bp3_file = os.path.join(pj_path, '.%'.join([pj_name,'17']))
        self.bp_flag = True if os.path.isfile(bp1_file) and os.path.isfile(bp2_file) and os.path.isfile(bp3_file) else False
        if self.bp_flag:
            with open(bp1_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

            with open(bp2_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            with open(bp3_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            self.bps_var_list = [var_list[8*j+i] for i in range(8) for j in range(3)]
            # print(self.bps_var_list)

            self.bps_fat_list = [[var for var in self.bps_var_list if '1' in var and 'xy' not in var],
                                 [var for var in self.bps_var_list if '2' in var and 'xy' not in var],
                                 [var for var in self.bps_var_list if '3' in var and 'xy' not in var]]

        # blade aerodynamic axis
        ba1_file = os.path.join(pj_path, '.%'.join([pj_name,'59']))
        ba2_file = os.path.join(pj_path, '.%'.join([pj_name,'60']))
        ba3_file = os.path.join(pj_path, '.%'.join([pj_name,'61']))
        self.ba_flag = True if os.path.isfile(ba1_file) and os.path.isfile(ba2_file) and os.path.isfile(ba3_file) else False
        if self.ba_flag:
            with open(ba1_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

            with open(ba2_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            with open(ba3_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            self.bas_var_list= [var_list[8 * j + i] for i in range(8) for j in range(3)]
            # print(self.bas_var_list)

            self.bas_fat_list = [[var for var in self.bas_var_list if '1' in var and 'xy' not in var],
                                 [var for var in self.bas_var_list if '2' in var and 'xy' not in var],
                                 [var for var in self.bas_var_list if '3' in var and 'xy' not in var]]

        # blade user axis
        bu1_file = os.path.join(pj_path, '.%'.join([pj_name,'62']))
        bu2_file = os.path.join(pj_path, '.%'.join([pj_name,'63']))
        bu3_file = os.path.join(pj_path, '.%'.join([pj_name,'64']))
        self.bu_flag = True if os.path.isfile(bu1_file) and os.path.isfile(bu2_file) and os.path.isfile(bu3_file) else False
        if self.bu_flag:
            with open(bu1_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

            with open(bu2_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            with open(bu3_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

            self.bus_var_list= [var_list[8 * j + i] for i in range(8) for j in range(3)]
            # print(self.bus_var_list)

            self.bus_fat_list = [[var for var in self.bus_var_list if '1' in var and 'xy' not in var],
                                 [var for var in self.bus_var_list if '2' in var and 'xy' not in var],
                                 [var for var in self.bus_var_list if '3' in var and 'xy' not in var]]

        # hub rotating
        hr_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))
        self.hr_flag = True if os.path.isfile(hr_file) else False
        if self.hr_flag:
            with open(hr_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

                self.hr_var_list = var_list[:8]
            # print(self.hr_var_list)

            self.hr_fat_list = [var for var in self.hr_var_list if 'yz' not in var]

        # hub stationary
        hs_file = os.path.join(pj_path, '.%'.join([pj_name,'23']))
        self.hs_flag = True if os.path.isfile(hs_file) else False
        if self.hs_flag:
            with open(hs_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                self.hs_var_list = var_line.strip().split('\t')[1][1:-1].split("' '")
            # print(self.hs_var_list)

            self.hs_fat_list = [var for var in self.hs_var_list if 'yz' not in var]

        # yaw bearing
        yb_file = os.path.join(pj_path, '.%'.join([pj_name,'24']))
        self.yb_flag = True if os.path.isfile(yb_file) else False
        if self.yb_flag:
            with open(yb_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                self.yb_var_list = var_line.strip().split('\t')[1][1:-1].split("' '")
            # print(self.yb_var_list)

            self.yb_fat_list = [var for var in self.yb_var_list if 'xy' not in var]

        # tower
        tr_file = os.path.join(pj_path, '.%'.join([pj_name,'25']))
        self.tr_flag = True if os.path.isfile(tr_file) else False
        if self.tr_flag:
            with open(tr_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                self.tr_var_list = var_line.strip().split('\t')[1][1:-1].split("' '")
            # print(self.tr_var_list)

            self.tr_fat_list = [var for var in self.tr_var_list if 'yz' not in var]
            self.tr_fat_list = [var for var in self.tr_fat_list if 'xy' not in var]

        # extrapolation
        ep1_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))
        ep2_file = os.path.join(pj_path, '.%'.join([pj_name,'18']))
        ep3_file = os.path.join(pj_path, '.%'.join([pj_name,'19']))
        ep4_file = os.path.join(pj_path, '.%'.join([pj_name,'20']))
        self.ep_flag = True if os.path.isfile(ep1_file) and\
                               os.path.isfile(ep2_file) and\
                               os.path.isfile(ep3_file) and\
                               os.path.isfile(ep4_file) else False
        if self.ep_flag:
            with open(ep1_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[8:]
                self.ep_var_list = [var_list[8*j+i] for i in range(2) for j in range(3)]

            with open(ep2_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[0]
                # print(var_list)
                self.ep_var_list.append(var_list)

            with open(ep3_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[0]
                # print(var_list)
                self.ep_var_list.append(var_list)

            with open(ep4_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[0]
                # print(var_list)
                self.ep_var_list.append(var_list)
            # print(self.ep_var_list)

        # tower clearance
        tc_file = os.path.join(pj_path, '.%'.join([pj_name,'07']))
        self.tc_flag = True if os.path.isfile(tc_file) else False
        if self.tc_flag:
            with open(tc_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                self.tc_var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[-3:]
            # print(self.tc_var_list)

if __name__ == '__main__':

    pj_path  = r"\\172.20.0.4\fs02\CE\W2500-135-90\run_50%\DLC12\12_aa-01\12_aa-01.$PJ"
    Get_Variable(pj_path)