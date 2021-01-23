# ！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2019/10/26 11:14
# @Author  : CE
# @File    : joblist.py
'''
1 本程序只适用于通过additional controller parameters对话框来设置控制器参数，不适用于有多个控制器时的模型；
2 有多个控制器的模型需要在每个控制器中定义单独的参数文件；
'''

import os
import multiprocessing as mp

# 修改joblist中输出路径
class modify_joblist(object):
    def __init__(self,
                 joblist,
                 new_list,
                 job_path,
                 run_path,
                 dll_path,
                 xml_path,
                 newlist_path,
                 add_con_para,
                 other_con):
        '''
        定义输入输出
        :param joblist: 原始joblist的引用路径
        :param new_list: 新joblist的名称
        :param jobs_path: 将原来的jobs按照工况名生成新的jobs文件的路径
        :param run_path: 模型中仿真输出结果所在的run的路径
        :param dll_path: 新的dll文件的引用路径
        :param xml_path: 新的xml文件的引用路径

        '''

        self.job_list = joblist.replace('/','\\')
        self.new_list = new_list
        self.job_path = job_path.replace('/','\\') if job_path else None
        self.dll_path = dll_path.replace('/','\\') if dll_path else None
        self.xml_path = xml_path.replace('/','\\') if xml_path else None
        self.run_path = run_path.replace('/','\\')
        self.njl_path = newlist_path
        self.add_cont = add_con_para
        self.oth_cont = other_con

        self.old_path = None
        self.dlc_jobpath = {}  # dlc name：job old path
        self.rela_path = []

        self.get_jobpath()
        self.create_jobs()
        self.modify_joblist()

    def get_jobpath(self):
        '''
        读取joblist中包含的所有jobs及其路径，并记录到字典self.dlc_jobpath中去
        :return:
        '''
        dlc_name = None

        print('Start to read joblist and get dlc name and path!')
        with open(self.job_list, 'r') as f:
            for line in f.readlines():
                if '<ResultDir>' in line:
                    temp     = line.split('>')
                    res_dir  = temp[1].split('<')[0]
                    dlc_name = res_dir.split('\\')[-1]

                # 读取batch所引用的job的路径
                if '<InputFileDir>' in line:
                    temp = line.split('>')
                    # 原始文件中job的路径
                    old_path = temp[1].split('<')[0]
                    # 原始的dlc所对应的job的原始路径
                    self.dlc_jobpath[dlc_name] = old_path
        print('Read joblist is done!')

    def get_control_para(self, control_parameters):
        # 文本框中包含定义整个参数的语句
        # 功能可以修改也可以增加

        def part_func(list, n):
            '''
            将list按照固定长度进行分割
            '''
            new_list = []
            for i in range(0, len(list), n):
                # yield list[i:i + n]
                temp = list[i:i + n]
                new_list.append(temp)
            return new_list

        para_name  = []
        para_value = []
        # para_cont  = []  # 记录变量所包含的内容

        add_cont_para = control_parameters

        add_cont_para = add_cont_para.replace('<', '&lt;')
        add_cont_para = add_cont_para.replace('>', '&gt;')
        # print(add_cont_para)

        para_cont = add_cont_para.split('\n')
        for i in range(len(para_cont)):
            para_cont[i] = para_cont[i].strip() + '\n'
        # print(para_cont)

        for i in range(len(para_cont)):
            if i%5 == 1:
                para_name.append(para_cont[i])

            if i%5 == 3:
                para_value.append(para_cont[i])

        return para_name, para_value, part_func(para_cont, 5)

    def get_content_para(self, other_parameters):
        '''
        # Python是否可以按照字符串进行分割，比如“->”
        # 或者是按照替换其中的数字即可
        # 此部分有限制，.$PJ与.in文件并非所有的建模语句都相同，目前只涉及到相同部分的修改
        '''
        cont_value = {}
        cont_para  = other_parameters.split(';')

        for i in range(len(cont_para)):
            temp  = cont_para[i].split(',')
            old   = temp[0].strip()
            new   = temp[1].strip()

            cont_value[old] = new
        # print(cont_value)
        return cont_value

    # @pysnooper.snoop()
    def modify_prj(self, prj_path, dll_path, xml_path, control_parameters, other_parameters):
        '''
        读取和修改prj文件的路径
        :param prj_path: 原始prj文件的引用路径
        :param dll_path: 新的dll文件的引用路径
        :param xml_path: 新的xml文件的引用路径
        :return: 修改过的xml的内容
        '''

        content = ''

        # 记录控制器的个数，可能存在风速控制器
        # 默认第一个是风机控制器
        num_con = 1

        # 记录第一个<AdditionalParameters>
        # 在每个控制器中还可以通过<AdditionalParameters>定义parameter
        # 默认修改第一个参数
        # 如果加上占位符，可以不用改参数
        num_para = 1

        # 记录additional parameter介绍的位置
        para_index = dict()

        with open(prj_path, 'r', encoding='ISO-8859-1') as f:
            # print('  Open .$PJ sucessfully!')
            content_ori = f.readlines()

        for i in range(len(content_ori)):
            # 如果模型通过READ读入参数，那么修改读入参数文件的路径
            if '<AdditionalParameters>' in content_ori[i] and xml_path and num_para < 2:

                line = content_ori[i].split('>')[0] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                content_ori[i] = line

                num_para += 1
            # 如果之前的模型没有通过READ读取参数，那么可以通过以下命令增加参数文件的路径
            if '<AdditionalParameters />' in content_ori[i] and xml_path and num_para < 2:
                line = content_ori[i][:-4] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                # print(line)
                content_ori[i] = line

                num_para += 1
            if '<Filepath>' in content_ori[i] and dll_path and num_con < 2:

                temp = content_ori[i].split('>')
                line = temp[0] + '>' + dll_path + '</Filepath>' + '\n'

                content_ori[i] = line
                num_con += 1
                # 修改xml在修改dll之前，如果dll之前没有读取xml,dll之后的parameters也不用修改
                num_para += 1
            # 记录所有控制参数的名称和位置
            if '&lt;Name&gt;P_' in content_ori[i]:
                content_ori[i] = content_ori[i].strip() + '\n'
                para_index[content_ori[i]] = i

                # 修改parameter中的参数
                # if '&lt;InitialValue&gt;9999&lt;/InitialValue&gt;' in line:
                #
                #     line = '&lt;InitialValue&gt;750&lt;/InitialValue&gt;'
        # print('  Modifying dll and xml is done!\n')

        if control_parameters:
            # print('  Start to modify additional control parameter...')
            control = self.get_control_para(control_parameters)

            index_para = {v: k for k, v in para_index.items()}

            for j in range(len(control[0])):

                if control[0][j] in para_index.keys():

                    content_ori[para_index[control[0][j]] + 2] = control[1][j]

                else:

                    for i in range(len(control[2][j])):

                        content_ori.insert(max(para_index.values()) + 4 + i, control[2][j][i])

        if other_parameters:

            cont_value = self.get_content_para(other_parameters)
            # print(cont_value)

            for i in range(len(content_ori)):

                for k in cont_value.keys():
                    # print(k)
                    # 可能会出现相同的名称，如RHO和RHOW
                    vname = k + '\t'

                    if vname in content_ori[i]:

                        if '->' not in cont_value[k]:

                            content_ori[i] = k + '\t ' + cont_value[k] + '\n'

                        else:

                            temp = cont_value[k].split('->')

                            content_ori[i] = content_ori[i].replace(temp[0], temp[1])

        for line in content_ori:

            content += line
            # print(line)
        # print(len(content_ori))

        return content

    def modify_in(self, in_path, dll_path, xml_path, control_parameters, other_parameters):
        '''
        读取并修改in文件的内容
        :param in_path: 原始in文件的引用路径
        :param dll_path: 新的dll文件的引用路径
        :param xml_path: 新的xml文件的引用路径
        :return: 修改过的in文件内容
        '''

        content = ''

        # 记录控制器的个数，可能存在风速控制器
        # 默认第一个是风机控制器
        num_con = 1

        # 记录第一个<AdditionalParameters>
        # 在每个控制器中还可以通过<AdditionalParameters>定义parameter
        # 默认修改第一个参数
        # 如果加上占位符，可以不用改参数
        num_para = 1

        # 记录additional parameter介绍的位置
        para_index = dict()

        with open(in_path, 'r', encoding='ISO-8859-1') as f:

            # print('  Open .in sucessfully!')
            content_ori = f.readlines()

        for i in range(len(content_ori)):

            # 如果模型通过READ读入参数，那么修改读入参数文件的路径
            if '<AdditionalParameters>' in content_ori[i] and xml_path and num_para < 2:

                line = content_ori[i].split('>')[0] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                content_ori[i] = line

                num_para += 1

            # 如果之前的模型没有通过READ读取参数，那么可以通过以下命令增加参数文件的路径
            if '<AdditionalParameters />' in content_ori[i] and xml_path and num_para < 2:

                line = content_ori[i][:-4] + '>' + 'READ ' + xml_path + '</AdditionalParameters>' + '\n'
                # print(line)
                content_ori[i] = line

                num_para += 1

            if '<Filepath>' in content_ori[i] and dll_path and num_con < 2:

                temp = content_ori[i].split('>')
                line = temp[0] + '>' + dll_path + '</Filepath>' + '\n'

                content_ori[i] = line
                num_con += 1
                # 修改xml在修改dll之前，如果dll之前没有读取xml,dll之后的parameters也不用修改
                num_para += 1

            # 记录所有控制参数的名称和位置
            if '&lt;Name&gt;P_' in content_ori[i]:

                para_index[content_ori[i]] = i

                # 修改parameter中的参数
                # if '&lt;InitialValue&gt;9999&lt;/InitialValue&gt;' in line:
                #
                #     line = '&lt;InitialValue&gt;750&lt;/InitialValue&gt;'
        # print('  Modifying dll and xml is done!\n')

        if control_parameters:
        # print('  Start to modify additional control parameter...')
            control = self.get_control_para(control_parameters)

            index_para = {v: k for k, v in para_index.items()}

            for j in range(len(control[0])):

                if control[0][j] in para_index.keys():

                    content_ori[para_index[control[0][j]] + 2] = control[1][j]

                else:

                    for i in range(len(control[2][j])):

                        content_ori.insert(max(para_index.values()) + 4 + i, control[2][j][i])

        if other_parameters:

            # print('  Start to modify the project content...')
            cont_value = self.get_content_para(other_parameters)

            for i in range(len(content_ori)):

                for k in cont_value.keys():

                    # 可能会出现相同的名称，如RHO和RHOW
                    vname = k + '\t'

                    if vname in content_ori[i]:

                        if '->' not in cont_value[k]:

                            content_ori[i] = k + '\t ' + cont_value[k] + '\n'

                        else:

                            temp = cont_value[k].split('->')

                            content_ori[i] = content_ori[i].replace(temp[0], temp[1])

        content = ''.join(content_ori)

        with open(in_path, 'w+', encoding='ISO-8859-1') as f:
            f.write(content)

    def write_txt(self, output_path, content, file):
        '''
        将修改过的pj或in写入新的路径下
        :param output_path: pj和in所在文件
        :param content: 修改过的pj和in文件的内容
        :param file: 输出文件名称，如powprod.$PJ
        :return:
        '''

        file_path = output_path + os.sep + file
        with open(file_path, 'w+', encoding='ISO-8859-1') as f:
            f.write(content)

    def create_jobs(self):
        '''
        将原始的jobs按照工况进行命名，然后保存到self.job_path下面
        :return:
        '''
        print('Start to create jobs...')
        pool = mp.Pool(processes=mp.cpu_count())
        for key, value in self.dlc_jobpath.items():

            if self.job_path:
                path = os.sep.join([self.job_path, key])
            else:
                path = os.sep.join([(value.split('\\')[:-1]), key])

            if not os.path.exists(path):
                os.makedirs(path)

            file_list = os.listdir(value)
            for file in file_list:
                if '.$PJ' in file:
                    prj_path = value + os.sep + file
                    cont_prj = self.modify_prj(prj_path, self.dll_path, self.xml_path, self.add_cont, self.oth_cont)
                    self.write_txt(path, cont_prj, file)

                if file.endswith('.in'):
                    in_path = value + os.sep + file
                    cont_in = self.modify_in(in_path, self.dll_path, self.xml_path, self.add_cont, self.oth_cont)
                    self.write_txt(path, cont_in, file)
        print('Jobs is done!')

    def modify_joblist(self):
        '''
        修改joblist的内容
        :return:
        '''
        print('Start to cteate joblist...')
        content = ''

        with open(self.job_list, 'r') as f:
            for line in f.readlines():
                # 修改job的仿真结果路径
                if '<ResultDir>' in line and self.run_path:
                    temp = line.split('>')
                    old_path = temp[1].split('<')[0]
                    dlc_path = '\\'.join(old_path.split('\\')[-2:])
                    new_path = self.run_path.replace('/', '\\') + os.sep + dlc_path

                    line = temp[0] + '>' + new_path + '</ResultDir>' + '\n'

                if '<InputFileDir>' in line:
                    temp = line.split('>')
                    old_path = temp[1].split('<')[0]
                    dlc_name = [k for k, v in self.dlc_jobpath.items() if old_path == v]

                    if self.job_path:
                        new_path = os.sep.join([self.job_path, dlc_name[0]])
                        self.rela_path = '\\'.join(new_path.split('\\')[-2:])

                        line = temp[0] + '>' + new_path + '</InputFileDir>' + '\n'

                    else:
                        new_path = '\\'.join(old_path.split('\\')[:-1]) + os.sep + dlc_name[0]
                        self.rela_path = '\\'.join(new_path.split('\\')[-2:])

                        line = temp[0] + '>' + new_path + '</InputFileDir>' + '\n'
                        # print('reletive path:', self.rela_path)

                # 相对路径似乎不起作用
                if '<InputFileDirRelativeToBatch>' in line:
                    temp = line.split('>')
                    line = temp[0] + '>' + self.rela_path + '</InputFileDirRelativeToBatch>' + '\n'

                content += line

        if not self.njl_path:
            job_list_path = os.sep.join(self.job_list.split(os.sep)[:-1])
        else:
            job_list_path = self.njl_path

        if self.new_list:
            new_list = self.new_list + '.joblist'
        else:
            new_list = self.job_list.split('\\')[-1].split('.')[0] + '_new.joblist'

        self.write_txt(job_list_path, content, new_list)
        print('Joblist is done')

if __name__ == '__main__':

    # 使用方法：
    # 1. 依次定义joblist/dll_path/jobs_path/result_dir/xml_path/new_joblist的路径；
    # 2. 其中joblist、jobs_path、result_dir和new_joblist是必须定义的；
    # 3. dll_path和xml_path可以根据需求进行定义，如不用修改直接定义空即可；
    # 4. 定义路径时，请注意引号前的r，该字符不可或缺。

    # start = time.time()

    # 原joblist的引用路径
    joblist    = r"E:\python/01_Joblist/batch/Job Lists/dlc12.joblist"
    jobs_path  = r"E:\python\01_Joblist\batch\Jobs test"
    result_dir = r"E:\python\01_Joblist"
    new_name   = 'test'
    dll_path   = r"E:\python\01_Joblist\Controller_loop3.2_V0.0\Win32Discon_loop3.2_v0.0.dll"
    xml_path   = r"E:\python\01_Joblist\Controller_loop3.2_V0.0\Parameters.xml"
    newlist_path = r'E:\python\01_Joblist\batch'


    modify_joblist(joblist, new_joblist, jobs_path, result_dir, dll_path, xml_path, None, None, None)
