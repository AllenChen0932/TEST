#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/1/2020 8:04 PM
# @Author  : CE
# @File    : Get_Variable.py

import os
import itertools

class Get_Variables(object):

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

        # ultimate variable
        self.br_var_list   = []
        self.br_mxy        = []
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
        self.hr_all_list   = []
        self.brs_mxy_list  = {}
        
        # fatigue variable
        self.br1_fat_list  = []
        self.br2_fat_list  = []
        self.br3_fat_list  = []
        self.br_mxy_list   = []
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

        self.var_ext = {}
        self.get_var_extension()
        self.get_variable()

    def get_var_extension(self):

        pj_path, pj_file = os.path.split(self.prj_path)
        file_list = [file for file in os.listdir(pj_path) if file.split("%")[-1].isdigit()]

        for file in file_list:
            file_path = os.path.join(pj_path, file)
            with open(file_path) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split("'")[1::2]

                for var in var_list:
                    self.var_ext[var] = file.split('%')[-1]

    def get_variable(self):

        pj_path, pj_file = os.path.split(self.prj_path)
        # print(pj_path, pj_file)
        tower_flag = False

        with open(self.prj_path) as f:
            lines = f.readlines()
            for line in lines:

                if line.startswith('DIST') and 'DISTMODE' not in line:
                    self.bld_radius = line.strip().replace(' ', '').split('\t',1)[1].split(',')
                    self.bld_radius = [k for k, g in itertools.groupby(self.bld_radius)]
                    # print(len(self.bld_radius))
                    # print('blade radius:', self.bld_radius)

                if line.startswith('TMODEL'):
                    self.twr_model = line.strip().split()[1]
                    # print('tower model:', self.twr_model)

                if 'TGEOM' in line:
                    tower_flag = True
                if tower_flag and 'MEND' in line:
                    tower_flag = False

                if line.startswith('TJ') and tower_flag:
                    self.twr_height = line.strip().split()[1].split(',')
                    # print(self.twr_height)
                    # print(len(self.twr_height))
                    # print('tower height:', self.twr_height)

                if line.startswith('NTE') and self.twr_model == '3':
                    self.twr_height = [str(i) for i in range(1, int(line.split()[-1])+1)]
                    # print('tower height:', self.twr_height)

                if line.startswith('BLOADS_STS'):
                    self.bpa_output = line.strip().split('\t',1)[1].split(',')
                    # print('principal axis:', self.bpa_output)

                if line.startswith('BMFLAP'):
                    self.bpa_out_type = line.strip().split('\t')[1]
                    # print('principal output:', self.bpa_out_type)

                if line.startswith('UNTWIST'):
                    self.bra_out_type = line.strip().split('\t')[1]
                    # print('root axes output:', self.bra_out_type)

                if line.startswith('BL_AERO'):
                    self.baa_out_type = line.strip().split('\t')[1]
                    # print('aerodynamic output:', self.baa_out_type)

                if line.startswith('BL_USER'):
                    self.bua_out_type = line.strip().split('\t')[1]
                    # print('user output:', self.bua_out_type)

                # if self.twr_model == 2:
                if line.startswith('TLOADS_STS'):
                    self.twr_output = line.strip().split('\t',1)[1].split(',')
                    # print('tower output:', self.twr_output)

                # elif self.twr_model == 3:
                if line.startswith('TLOADS_ELS'):
                    self.twr_output = line.strip().split()[1:]
                    # print('tower output:', self.twr_output)

        pj_name = os.path.splitext(pj_file)[0]

        # blade root
        br_file = os.path.join(pj_path, '.%'.join([pj_name,'22']))
        self.br_flag = True if os.path.isfile(br_file) else False
        if self.br_flag:
            with open(br_file) as f:
                lines    = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[8:]

            if len(var_list) == 24:
                self.br_var_list = [var_list[8*j+i] for i in range(8) for j in range(3)]
                self.br_fat_list = [[var for var in var_list if '1' in var and 'xy' not in var],
                                    [var for var in var_list if '2' in var and 'xy' not in var],
                                    [var for var in var_list if '3' in var and 'xy' not in var]]
            elif len(var_list) == 16:
                self.br_var_list = var_list[0:8]
                self.br_fat_list = [[var for var in var_list if '1' in var and 'xy' not in var],
                                    [var for var in var_list if '2' in var and 'xy' not in var]]
            else:
                self.br_var_list = var_list[0:8]
                self.br_fat_list = [[var for var in var_list if '1' in var and 'xy' not in var]]
            # print(self.br_fat_list)

        # blade root mxy angle
        br_file = os.path.join(pj_path, '.%'.join([pj_name,'700']))
        self.br_mxy_flag = True if os.path.isfile(br_file) else False
        if self.br_mxy_flag:
            with open(br_file) as f:
                lines    = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split("'")[1::2]
                self.br_mxy_list = var_list
                # print('br_mxy:', var_list)

        # blade root axes mxy for each section
        brs_mxy_list = [file for file in os.listdir(pj_path) if '.%9' in file]
        self.brs_mxy_flag = True if brs_mxy_list else False
        if self.brs_mxy_flag:
            for file in brs_mxy_list:
                file_path = os.path.join(pj_path, file)
                with open(file_path) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]
                    # print(var_line.strip().split('\t')[1])
                    var_list = var_line.strip().split("'")[1::2]
                    sec_name = '_'.join(var_list[0].split('_')[:2])
                    self.brs_mxy_list[sec_name] = var_list

        # blade root axes mxy angle
        brs1_file = os.path.join(pj_path, '.%'.join([pj_name, '800']))
        self.brs1_mxy_flag = True if os.path.isfile(brs1_file) else False
        if self.brs1_mxy_flag:
            with open(brs1_file) as f:
                lines    = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split("'")[1::2]
                self.brs1_mxy_list = var_list
                # print('brs_mxy:',var_list)

        brs2_file = os.path.join(pj_path, '.%'.join([pj_name, '810']))
        self.brs2_mxy_flag = True if os.path.isfile(brs2_file) else False
        if self.brs2_mxy_flag:
            with open(brs2_file) as f:
                lines    = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split("'")[1::2]
                self.brs2_mxy_list = var_list
                # print('brs_mxy:',var_list)

        brs3_file = os.path.join(pj_path, '.%'.join([pj_name, '820']))
        self.brs3_mxy_flag = True if os.path.isfile(brs3_file) else False
        if self.brs3_mxy_flag:
            with open(brs3_file) as f:
                lines    = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]
                # print(var_line.strip().split('\t')[1])
                var_list = var_line.strip().split("'")[1::2]
                self.brs3_mxy_list = var_list
                # print('brs_mxy:',var_list)

        # blade root axes
        bs1_file = os.path.join(pj_path, '.%'.join([pj_name,'41']))
        bs2_file = os.path.join(pj_path, '.%'.join([pj_name,'42']))
        bs3_file = os.path.join(pj_path, '.%'.join([pj_name,'43']))
        self.bs_flag = any([os.path.isfile(bs1_file), os.path.isfile(bs2_file), os.path.isfile(bs3_file)])
        if self.bs_flag:
            if os.path.isfile(bs1_file):
                with open(bs1_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")
                self.brs_var_list = var_list
                self.brs_fat_list = [[var for var in self.brs_var_list if '1' in var and 'xy' not in var]]

                if os.path.isfile(bs2_file):
                    with open(bs2_file) as f:
                        lines = f.readlines()
                        var_line = [line for line in lines if line.startswith('VARIAB')][0]

                        var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                    if os.path.isfile(bs3_file):
                        with open(bs3_file) as f:
                            lines = f.readlines()
                            var_line = [line for line in lines if line.startswith('VARIAB')][0]

                            var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                        self.brs_var_list = [var_list[8*j+i] for i in range(8) for j in range(3)]
                        self.brs_fat_list = [[var for var in self.brs_var_list if '1' in var and 'xy' not in var],
                                             [var for var in self.brs_var_list if '2' in var and 'xy' not in var],
                                             [var for var in self.brs_var_list if '3' in var and 'xy' not in var]]
            # print('brs:',self.brs_var_list)

        # blade principle axis
        bp1_file = os.path.join(pj_path, '.%'.join([pj_name,'15']))
        bp2_file = os.path.join(pj_path, '.%'.join([pj_name,'16']))
        bp3_file = os.path.join(pj_path, '.%'.join([pj_name,'17']))
        self.bp_flag = any([os.path.isfile(bp1_file), os.path.isfile(bp2_file), os.path.isfile(bp3_file)])
        if self.bp_flag:
            if os.path.isfile(bp1_file):
                with open(bp1_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

                self.bps_var_list = var_list
                self.bps_fat_list = [[var for var in self.bps_var_list if '1' in var and 'xy' not in var]]

                if os.path.isfile(bp2_file):
                    with open(bp2_file) as f:
                        lines = f.readlines()
                        var_line = [line for line in lines if line.startswith('VARIAB')][0]

                        var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                    if os.path.isfile(bp3_file):
                        with open(bp3_file) as f:
                            lines = f.readlines()
                            var_line = [line for line in lines if line.startswith('VARIAB')][0]

                            var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                        self.bps_var_list = [var_list[8*j+i] for i in range(8) for j in range(3)]
                        self.bps_fat_list = [[var for var in self.bps_var_list if '1' in var and 'xy' not in var],
                                             [var for var in self.bps_var_list if '2' in var and 'xy' not in var],
                                             [var for var in self.bps_var_list if '3' in var and 'xy' not in var]]

        # blade aerodynamic axis
        ba1_file = os.path.join(pj_path, '.%'.join([pj_name,'59']))
        ba2_file = os.path.join(pj_path, '.%'.join([pj_name,'60']))
        ba3_file = os.path.join(pj_path, '.%'.join([pj_name,'61']))
        self.ba_flag = any([os.path.isfile(ba1_file), os.path.isfile(ba2_file), os.path.isfile(ba3_file)])
        if self.ba_flag:

            if os.path.isfile(ba1_file):
                with open(ba1_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

                self.bas_var_list = var_list
                self.bas_fat_list = [[var for var in self.bas_var_list if '1' in var and 'xy' not in var]]

                if os.path.isfile(ba2_file):
                    with open(ba2_file) as f:
                        lines = f.readlines()
                        var_line = [line for line in lines if line.startswith('VARIAB')][0]

                        var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                    if os.path.isfile(ba3_file):

                        with open(ba3_file) as f:
                            lines = f.readlines()
                            var_line = [line for line in lines if line.startswith('VARIAB')][0]

                            var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                        self.bas_var_list= [var_list[8*j+i] for i in range(8) for j in range(3)]
                        self.bas_fat_list = [[var for var in self.bas_var_list if '1' in var and 'xy' not in var],
                                             [var for var in self.bas_var_list if '2' in var and 'xy' not in var],
                                             [var for var in self.bas_var_list if '3' in var and 'xy' not in var]]

        # blade user axis
        bu1_file = os.path.join(pj_path, '.%'.join([pj_name,'62']))
        bu2_file = os.path.join(pj_path, '.%'.join([pj_name,'63']))
        bu3_file = os.path.join(pj_path, '.%'.join([pj_name,'64']))
        self.bu_flag = any([os.path.isfile(bu1_file), os.path.isfile(bu2_file), os.path.isfile(bu3_file)])
        if self.bu_flag:
            if os.path.isfile(bu1_file):
                with open(bu1_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")

                self.bus_var_list = var_list
                self.bus_fat_list = [[var for var in self.bus_var_list if '1' in var and 'xy' not in var]]

                if os.path.isfile(bu2_file):
                    with open(bu2_file) as f:
                        lines = f.readlines()
                        var_line = [line for line in lines if line.startswith('VARIAB')][0]

                        var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                    if os.path.isfile(bu3_file):
                        with open(bu3_file) as f:
                            lines = f.readlines()
                            var_line = [line for line in lines if line.startswith('VARIAB')][0]

                            var_list.extend(var_line.strip().split('\t')[1][1:-1].split("' '"))

                        self.bus_var_list= [var_list[8*j+i] for i in range(8) for j in range(3)]
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
                self.hr_all_list = var_list
            # print(self.hr_var_list)

            self.hr_fat_list = [var for var in self.hr_var_list if 'yz' not in var]
            # print(self.hr_fat_list)

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

        # nacelle acceleration
        nac_acc = os.path.join(pj_path, '.%'.join([pj_name,'26']))
        self.nac_flag = True if os.path.isfile(nac_acc) else False
        if self.nac_flag:
            with open(nac_acc) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                var_list = var_line.strip().split("'")[1::2][-6:]

            self.nac_acc_list = var_list

        # rotor speed
        rs_file = os.path.join(pj_path, '.%'.join([pj_name,'05']))
        self.rs_flag = True if os.path.exists(rs_file) else False
        # mean pitch angle and hub wind speed magnitude
        ws_file = os.path.join(pj_path, '.%'.join([pj_name,'07']))
        self.ws_flag = True if os.path.exists(ws_file) else False
        # drive train
        self.mb_flag = all([self.ws_flag, self.rs_flag, self.hs_flag])    # flag for main bearing statistics

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
        self.ep_flag = all([os.path.isfile(ep1_file),
                            any([os.path.isfile(ep2_file), os.path.isfile(ep3_file),os.path.isfile(ep4_file)])])
        if self.ep_flag:
            if os.path.isfile(ep1_file):
                with open(ep1_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[8:]
                    if len(var_list)==24:
                        self.ep_var_list = [var_list[8*j+i] for i in range(2) for j in range(3)]
                    else:
                        self.ep_var_list = [var_list[:2]]

            if os.path.isfile(ep2_file):
                with open(ep2_file) as f:
                    lines = f.readlines()
                    var_line = [line for line in lines if line.startswith('VARIAB')][0]

                    var_list = var_line.strip().split('\t')[1][1:-1].split("' '")[0]
                    # print(var_list)
                    self.ep_var_list.append(var_list)

            if os.path.isfile(ep3_file) and os.path.isfile(ep4_file):

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

        dt_file = os.path.join(pj_path, '.%'.join([pj_name,'05']))
        self.dt_flag = True if os.path.isfile(dt_file) else False
        if self.dt_flag:
            with open(dt_file) as f:
                lines = f.readlines()
                var_line = [line for line in lines if line.startswith('VARIAB')][0]

                self.dt_var_list = var_line.strip().split("'")[1::2]

if __name__ == '__main__':

    pj_path  = r"\\172.20.0.4\fs02\CE\W2500-135-90\run_50%\DLC12\12_aa-01\12_aa-01.$PJ"
    # Get_Variables(pj_path)