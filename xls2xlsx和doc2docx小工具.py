# -*- coding: utf-8 -*-

import os, shutil, sys
from changeOffice import Change

def mkdir(path):
    path = path.strip()  # 去除首位空格
    path = path.rstrip("\\")  # 去除尾部 \ 符号
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
        print(path + ' 创建成功')
        return True
    else:
        print(path + ' 目录已存在')
        return False

mkdir('.\\备份')
for fileName in os.listdir('.\\#转换'):
    if '工具' not in fileName:
        if 'xls' in fileName or 'xlsx' in fileName or 'doc' in fileName or 'docx' in fileName:
            shutil.copy('.\\#转换\\' + fileName, '.\\备份')
            pass

#转换文件，可能转出的文件读写空值，那么还得利用WPS或者LIBRE OFFICE
Goal_dir = Change(".\\#转换")
Goal_dir.doc2docx()
Goal_dir.xls2xlsx()
Goal_dir.ppt2pptx()

for fileName in os.listdir('.\\备份'):
    if '工具' not in fileName:
        if 'xls' in fileName or 'xlsx' in fileName or 'doc' in fileName or 'docx' in fileName:
            shutil.copy('.\\备份\\' + fileName, '.\\#转换')
            pass

isExists = os.path.exists('.\\备份')
if isExists:
    shutil.rmtree('.\\备份')

input('转换成功')

