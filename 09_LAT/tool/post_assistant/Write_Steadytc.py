#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 7/20/2020 8:48 PM
# @Author  : CE
# @File    : Steady_TC.py

import os

def Steadytc(pj_path, post_path):
    content   = ''

    with open(pj_path) as f:
        lines = f.readlines()
        spload_flag = False
        spload_cont = True

        for line in lines:
            if line.startswith('CALCULATION'):
                content += 'CALCULATION	8\n'
                continue

            if line.startswith('OPTIONS'):
                content += 'OPTIONS	4277\n'
                continue

            if line.startswith('RHO') and 'RHOW' not in line:
                content += 'RHO	 1E-9\n'
                continue

            if line.startswith('GRAVITY'):
                content += 'GRAVITY	 1E-9\n'
                continue

            if line.startswith('SPLOAD'):
                spload_flag = True
                continue

            if 'MEND' in line and spload_flag:
                spload_flag = False
                continue

            if spload_flag:
                continue

            if line.startswith('  <Name>'):
                content += '  <Name>sparked</Name>\n'
                continue

            elif line.startswith('  <Calculation>'):
                content += '  <Calculation>8</Calculation>\n'

            elif line.startswith('0WINDSEL') or line.startswith('0SPLOAD'):
                pass

            elif line.startswith('0') and spload_cont:
                content += 'MSTART SPLOAD\n' \
                           'WSPEED	 1e-09\n' \
                           'AZIMUTH	 0\n' \
                           'YAW	 0\n' \
                           'INC	 0\n' \
                           'PITCH	 0\n' \
                           'SWEEP	A\n' \
                           'END	 6.28318\n' \
                           'STEP	 1.74532777777778E-02\n' \
                           'MEND\n\n'
                content += line
                spload_cont = False

            elif line.startswith('		]]>'):
                content += '0SPLOAD\n\n' \
                           '		]]>\n'
            else:
                content += line

    # file_path = os.path.join(post_path, 'steady')
    if not os.path.isdir(post_path):
        os.makedirs(post_path)

    prj_path = os.path.join(post_path, 'tc.$PJ')

    with open(prj_path, 'w+') as f:
        f.write(content)

    bladed_path = None
    if os.path.isfile(r"C:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"):
        bladed_path  = r"C:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"
    elif os.path.isfile(r"D:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"):
        bladed_path = r"D:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"
    elif os.path.isfile(r"E:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"):
        bladed_path = r"E:\Program Files (x86)\DNV GL\Bladed 4.8\Bladed_m72.exe"

    command = '"%s" -Prj "%s" -RunDir "%s" -ResultPath "%s"' %(bladed_path, prj_path, post_path, post_path)

    os.popen(command)

    in_path = os.path.join(post_path, 'dtbladed.in')
    content = ''
    with open(in_path, 'w+') as f:
        lines = f.readlines()
        for line in lines:
            if line.startswith('OPTNS'):
                content += 'OPTNS	4276\n'
            else:
                content += line
        f.write(content)

if __name__ == '__main__':

    pj_path   = r"\\172.20.4.132\fs02\CE\V3\loop06\run_1121\DLC12\12_aa-01\12_aa-01.$PJ"
    post_path = r'C:\Users\10700700\Desktop\py\42_TC\Steady'
    Steadytc(pj_path, post_path)


