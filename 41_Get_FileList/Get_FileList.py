#！usr/bin/python3
# -*- coding: utf-8 -*-
# @Time    : 2020/12/19 14:43
# @Author  : CE
# @File    : Get_FileList.py

import os
import time
import collections

def get_all_dirs_deep(path):
    "Deep First Search"
    file_lists  = []
    stack = []
    stack.append(path)

    while len(stack)!=0:
        current_path = stack.pop()
        file_list = os.listdir(current_path)

        for file in file_list:
            file_path = os.path.join(current_path, file)

            if os.path.isdir(file_path):
                stack.append(file_path)
            else:
                file_lists.append(file_path)

def get_all_dirs_wide(path):
    "Wide First Search"
    queue = collections.deque()
    queue.append(path)

    file_lists = []

    while len(queue)!=0:
        current_path = queue.popleft()
        file_list = os.listdir(current_path)

        for file in file_list:
            file_path = os.path.join(current_path, file)

            if os.path.isdir(file_path):
                queue.append(file_path)
            else:
                file_lists.append(file_path)

def get_all_dirs_walk(path):
    file_lists = []
    for root,dirs,files in os.walk(path):
        for file in files:
            file_path = os.path.join(root, file)
            file_lists.append(file_path)
    print(len(file_lists))

path = r'\\172.20.4.132\fs02\CE\V3\loop06\run_1121'

start = time.time()
get_all_dirs_deep(path)
finish = time.time()
print('deep total time:', finish-start)

start = time.time()
get_all_dirs_wide(path)
finish = time.time()
print('wide total time:', finish-start)

start = time.time()
get_all_dirs_walk(path)
finish = time.time()
print('walk total time:', finish-start)

