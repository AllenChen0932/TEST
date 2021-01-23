#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/6/2020 12:32 PM
# @Author  : CE
# @File    : project2joblist.py

import os
import time
import shutil
import multiprocessing as mp
from PyQt5.QtCore import pyqtSignal

from read_pj import read_pj

class Offshore_Joblist(object):

    emit_signal = pyqtSignal(str)

    def __init__(self, run_path):

        self.run_path = run_path

        self.dlc_list = {}     #dlc: load cases(design load case under run path)
        self.dlc_path = {}     #dlc: run path, for generating in template
        self.lcs_path = {}     #lcs: run path, for all load case
        self.npj_list = []     #lc list without project

        self.dlc_in   = {}     #dlc: in path
        self.in_cont  = {}     #in: content

    def generate_in(self, prj_path, result_path):

        dtbladed = r"C:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"
        command  = r'"%s" -Prj "%s" -RunDir "%s"' % (dtbladed, prj_path, result_path)
        print('begin to generate in......')
        print('start to process: %s' %time.strftime('%H:%M:%S',time.localtime()))
        os.popen(command).read()
        print('process finished: %s' %time.strftime('%H:%M:%S',time.localtime()))

    def get_pj(self):

        # design load case file under run path
        if 'DLC' not in ''.join(os.listdir(self.run_path)):
            raise Exception('No DLC under path %s' %self.run_path)
        dlc_list = [file for file in os.listdir(self.run_path)
                        if os.path.isdir(os.path.join(self.run_path, file)) and file.upper().startswith('DLC22')]

        for dlc in dlc_list:
            # dlc path
            dlc_path = os.path.join(self.run_path, dlc)

            if os.path.isdir(dlc_path):

                # load case under each design load case
                self.dlc_list[dlc] = [file for file in os.listdir(dlc_path) if
                                      os.path.isdir(os.path.join(dlc_path, file))]

                # first load case under each design load case for create template
                lc      = os.listdir(dlc_path)[0]
                lc_path = os.path.join(dlc_path, lc)
                pj_path = os.path.join(lc_path, [file for file in os.listdir(lc_path) if '.$PJ' in file.upper()][0])
                self.dlc_path[dlc] = pj_path

                # all load case under each design load case
                lc_list = os.listdir(dlc_path)
                for lc in lc_list:

                    lc_path = os.path.join(dlc_path, lc)
                    pj_path = os.path.join(lc_path, [file for file in os.listdir(lc_path) if '.$PJ' in file.upper()][0])

                    if pj_path:
                        self.lcs_path[lc] = [pj_path, lc_path]
                    else:
                        self.npj_list.append(lc)
        # print(self.dlc_path)

        if not self.dlc_path:
            raise Exception('%s\nNo Pj exist! Please choose a right run path!'%self.run_path)

    def create_template(self):

        current_dir = os.getcwd()
        # print(current_dir)
        temp_path   = os.path.join(current_dir, 'temp')

        if os.path.isdir(temp_path):
            shutil.rmtree(temp_path)

        # pool = mp.Pool(processes=3)
        for dlc, pj_path in self.dlc_path.items():

            lc_path = os.path.join(temp_path, dlc)
            if not os.path.isdir(lc_path):
                os.makedirs(lc_path)
                shutil.copy(pj_path, lc_path)
                temp_pj = os.path.join(lc_path, os.path.split(pj_path)[1])

                # pool.apply_async(self.generate_in, args=(temp_pj, lc_path))
                self.generate_in(temp_pj, lc_path)
                temp_in = os.path.join(lc_path, 'dtbladed.in')
                self.dlc_in[dlc] = temp_in

        # pool.close()
        # pool.join()

    def write_in(self, dlc, lc_list):

        key_flag   = None
        time_start = time.time()

        if 'DLC1' in dlc:

            kw_dlc1x = ['CURRENT', 'WINDSEL', 'WAVE', 'TIDE', 'INITCON', 'SIMCON']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                # print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc1x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += line
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc
                            continue

                        else:
                            content += line
                            # continue

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC2' in dlc:

            for lc in lc_list:

                kw_dlc2x = ['WINDSEL', 'WINDV', 'WAVE', 'TIDE', 'SIMCON', 'RTOL', 'DISCON', 'SAFETYSYSTEM', 'CURRENT', 'control']

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                # print(lc, dlc_in[0], dlc)

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key      = line.split()[1]
                            key_flag = True if key in kw_dlc2x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:

                                    if key == 'FAULTTIME':
                                        if kw['FAULTENABLED']=='True':
                                            content += '%s\t%s\n' %('FAULTTIME', kw['FAULTTIME'])
                                        else:
                                            continue

                                    elif key == 'FAULTVAL':
                                        if kw['FAULTENABLED']=='True':
                                            content += '%s\t%s\n' %('FAULTVAL',  kw['FAULTVAL'])
                                        else:
                                            continue

                                    elif key == 'FAILTIME':
                                        content += 'FAILTIME\t%s\n' %kw[key]
                                        content += 'RECOV\t%s\n' %kw[key]

                                    elif key == 'ATOLAZ':
                                        content += 'ATOLAZ\t%s\n' %kw[key]
                                        content += 'FAILRATE\t%s\n' %kw[key]

                                    elif key == 'DIRTYPE':
                                        content += 'DIRTYPE\t%s\n' %kw[key]
                                        if 'DIRRATE' in kw:
                                            content += 'DIRRATE\t%s\n' %kw['DIRRATE']
                                        else:
                                            continue

                                    else:
                                        content += '%s%s\n' %(line.split()[0], '\t' + kw[key] if kw[key] else '')
                            else:

                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC3' in dlc:

            kw_dlc3x = ['WINDSEL', 'WAVE', 'TIDE', 'SIMCON', 'Control']
            control_flag = None

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc3x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                            continue

                        elif line.startswith('        <AdditionalParameters>'):
                            content     += kw['Control']
                            control_flag = True
                            # continue
                        elif control_flag:
                            continue

                        elif '&lt;/Controller&gt;</AdditionalParameters>' in line:
                            control_flag = False

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC4' in dlc:

            kw_dlc4x = ['WINDSEL', 'WAVE', 'TIDE', 'INITCON', 'SIMCON', 'CURRENT']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc4x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line
                            # continue

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC5' in dlc:

            kw_dlc5x = ['WINDSEL', 'WAVE', 'TIDE', 'SAFETYSYSTEM', 'CURRENT']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc5x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line
                            # continue

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC6' in dlc:

            kw_dlc6x = ['WINDSEL', 'WAVE', 'TIDE', 'CURRENT', 'RTOL', 'SIMCON']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc6x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line
                            # continue

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC7' in dlc:

            kw_dlc7x = ['WINDSEL', 'WAVE', 'TIDE', 'CURRENT', 'RTOL', 'SIMCON']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()

                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc7x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line
                            # continue

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        if 'DLC8' in dlc:

            kw_dlc8x = ['WINDSEL', 'WAVE', 'TIDE', 'CURRENT', 'RTOL', 'INITCON', 'SIMCON']

            for lc in lc_list:

                content = ''
                pj_path = self.lcs_path[lc]
                kw      = read_pj(pj_path[0])
                dlc_in  = [key for key in self.dlc_in.keys() if dlc in key]
                print(dlc_in[0])

                with open(self.dlc_in[dlc_in[0]]) as f:

                    lines = f.readlines()
                    for line in lines:

                        if line.startswith('PATH'):
                            content += 'PATH\t%s\n' %pj_path[1]
                            continue

                        if line.startswith('RUNNAME'):
                            content += 'RUNNAME\t%s\n' %lc
                            continue

                        if line.startswith('MSTART'):
                            key = line.split()[1]
                            key_flag = True if key in kw_dlc8x else False
                            content += line
                            continue

                        if key_flag:
                            if 'MEND' not in line:
                                key = line.split()[0]
                                # print(kw[key])
                                if key in kw:
                                    if kw[key]:
                                        content += '%s\t%s\n' %(line.split()[0], kw[key])
                                    else:
                                        content += '%s\n' % (line.split()[0])
                                else:
                                    content += line
                            else:
                                key_flag = False
                                content += 'MEND\n'
                            continue

                        elif line.startswith('  <Name>'):
                            content += '  <Name>%s</Name>\n' %lc

                        else:
                            content += line

                    in_path = os.path.join(pj_path[1], 'dtbladed.in')

                    if os.path.isfile(in_path):
                        os.remove(in_path)

                    with open(in_path, 'w+') as f:
                        f.write(content)

        time_stop = time.time()
        print('%s cost %ss' % (dlc, time_stop-time_start))

    def run(self):
        print('start to transfer: %s' %time.strftime('%H:%M:%S',time.localtime()))
        self.get_pj()
        self.create_template()

        res = []
        pool = mp.Pool(processes=10)
        for dlc, lc_list in self.dlc_list.items():
            res.append(pool.apply_async(self.write_in, args=(dlc, lc_list)))

        pool.close()
        pool.join()
        print('transfer finished: %s' %time.strftime('%H:%M:%S',time.localtime()))

        return res

if __name__ == '__main__':

    run_path = r'\\172.20.0.4\fs03\CE\W6250-172-14m\run_0429'

    start = time.time()
    Offshore_Joblist(run_path).run()
    stop = time.time()
    print('multi thread cost %ss' %(stop-start))
