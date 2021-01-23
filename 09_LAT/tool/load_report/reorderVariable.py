# -*- coding: utf-8 -*-
# @Time    : 2020/07/18 9:55
# @Author  : CJG
# @File    : reorderVariable.py

def reorderVariable(varlist, reqlist):
    varout = []
    for req in reqlist:
        for var in varlist:
            if req in var.strip():
                varout.append(var)
                # varlist.remove(var)  # should NOT be done!
                # print(varlist)
                break
    # print(varout)
    return varout

if __name__ == '__main__':

    varlist = ['Blade root 1 Mx','Blade root 1 My','Blade root 1 Mxy','Blade root 1 Mz',
               'Blade root 1 Fx','Blade root 1 Fy','Blade root 1 Fxy','Blade root 1 Fz']
    reqlist = ('Mx','My','Mz','Fx','Fy','Fz')
    reorderVariable(varlist, reqlist)
    # print(varlist)
