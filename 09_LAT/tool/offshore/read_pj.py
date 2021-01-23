#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 5/28/2020 8:34 AM
# @Author  : CE
# @File    : read_pj.py


module_modify = ['WINDSEL','SIMCON','WINDV','DISCON','CURRENT','INITCON', 'RTOL',
                 'SAFETYSYSTEM','TIDES','WAVE','CONTROLFAIL','PITCHFAIL']

def read_pj(pj_path):

    kw_value = {}
    kw_flag  = False
    keywords = None

    control      = ''
    control_flag = None
    # keywords = None

    with open(pj_path) as f:

        lines = f.readlines()

        for line in lines:

            if line.startswith('MSTART'):
                keywords = line.strip().split()[1]

                if keywords in module_modify:
                    kw_flag = True

            if kw_flag and 'MEND' not in line:
                temp  = line.strip().split('\t') if '\t' in line else line.strip().split()

                if keywords != 'SAFETYSYSTEM':
                    kw    = temp[0]
                    value = temp[1] if len(temp)>1 else ''

                    kw_value[kw] = value
                    # print(kw,value)
                elif 'TIMEVAL'in line:
                    kw    = temp[0]
                    value = temp[1] if len(temp)>1 else ''

                    kw_value[kw] = value

            if 'MEND' in line:
                kw_flag = False

            if line.startswith('        <AdditionalParameters>'):
                # control += line
                control_flag = True

            if control_flag:
                control += line

            if '&lt;/Controller&gt;</AdditionalParameters>' in line:
                # control += line
                control_flag = False

        kw_value['Control'] = control

    return kw_value

if __name__ == '__main__':

    pj_path = r"\\172.20.0.4\fs03\CE\W6250-172-14m\run_0429\test_in\checked\powprod.$PJ"

    read_pj(pj_path)