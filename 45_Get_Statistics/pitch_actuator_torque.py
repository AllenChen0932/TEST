#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020/11/25 7:23
# @Author  : CE
# @File    : pitch_actuator_torque.py

import os
import Read_Bladed_v2 as rb


class PitchSystem(object):

    def __init__(self, dlc_path, diameter=3.87, fri_coef=0.03, sta_fri=25, efficiency=0.8736, trans_ratio=2968):
        self.dlc_path  = dlc_path
        self.fric_coef = fri_coef
        self.sta_fric  = sta_fri
        self.diameter  = diameter
        self.efficiency  = efficiency
        self.trans_ratio = trans_ratio

        self.get_power()

    def get_data(self, lc_path, lc_name):

        loads = rb.read_bladed(lc_name, lc_path, '22')
        rate  = rb.read_bladed(lc_name, lc_path, '08')

        omega = rate.data[:,3]   # rad/s
        Mxy = loads.data[:,10]
        Mz1  = loads.data[:,11]
        Fz  = loads.data[:,15]
        Fxy = loads.data[:,14]
        Mf  = float(self.fric_coef)*0.5*(4.4*Mxy+Fz*self.diameter+3.81*Fxy*self.diameter)+self.sta_fric
        # res = np.column_stack((omega, Mf))
        # Mb = Mz+np.where(res[:,0]>0.2, res[:,1], -res[:,1])
        # Mb = Mz+Mf
        # power1 = abs(Mb*omega)


        omega = rate.data[:,4]   # rad/s
        Mxy = loads.data[:,18]
        Mz2  = loads.data[:,19]
        Fz  = loads.data[:,23]
        Fxy = loads.data[:,22]
        Mf  = float(self.fric_coef)*0.5*(4.4*Mxy+Fz*self.diameter+3.81*Fxy*self.diameter)+self.sta_fric
        # res = np.column_stack((omega, Mf))
        # Mb = Mz+np.where(res[:,0]>0.2, res[:,1], -res[:,1])
        # Mb = Mz+Mf
        # power2 = abs(Mb*omega)

        omega = rate.data[:,5]   # rad/s
        Mxy = loads.data[:,26]
        Mz3  = loads.data[:,27]
        Fz  = loads.data[:,31]
        Fxy = loads.data[:,30]
        Mf  = float(self.fric_coef)*0.5*(4.4*Mxy+Fz*self.diameter+3.81*Fxy*self.diameter)+self.sta_fric
        # # res = np.column_stack((omega, Mf))
        # # Mb = Mz+np.where(res[:,0]>0.2, res[:,1], -res[:,1])
        # Mb = Mz+Mf
        # power3 = abs(Mb*omega)
        # print(Mz1.max(), Mz2.max(), Mz3.max())

        # print(np.average(power1+power2+power3)/1000/self.efficiency, np.max(power1+power2+power3)/1000/self.efficiency)
        # return np.average(power1+power2+power3)/1000,np.max(power1+power2+power3)/1000

    def get_pitch_actuator_power(self, lc_path, lc_name):

        pitch_system = rb.read_bladed(lc_name, lc_path, '08')

        omega  = pitch_system.data[:, -3]  # rad/s
        torque1 = pitch_system.data[:, -9]
        power1 = abs(torque1*omega)

        omega  = pitch_system.data[:, -2]  # rad/s
        torque2 = pitch_system.data[:, -8]
        power2 = abs(torque2*omega)

        omega  = pitch_system.data[:, -1]  # rad/s
        torque3 = pitch_system.data[:, -7]
        power3 = abs(torque3*omega)
        print(torque1.max(), torque2.max(), torque3.max())
        # print(np.average(power1+power2+power3)/1000, np.max(power1+power2+power3)/1000)

    def get_power(self):
        lc_list = os.listdir(self.dlc_path)
        # print(lc_list)
        for lc in lc_list:
            lc_path = os.path.join(self.dlc_path, lc)
            # print(lc)
            self.get_data(lc_path, lc)
            self.get_pitch_actuator_power(lc_path, lc)
            # print(res)

if __name__ == '__main__':
    path = r'\\172.20.4.132\fs02\CE\V3\loop06\test\pitch_power\run_1121\DLC12'

    PitchSystem(path)