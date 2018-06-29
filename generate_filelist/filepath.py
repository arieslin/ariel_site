# -*- coding:utf-8 -*-
"""
crreated on 2018-4-14
@author: linwei
"""
import os
import xlwt

def list_files(start_path):
    for root, dirs, files in os.walk(start_path):
        level = root.replace(start_path, "").count(os.sep)
        dir_indent = "|    " * (level - 1) + "|=="
        file_indent = "|    " * level + "|--"
        if not level:
            print(".")
        else:
            print ('{} {}'.format(dir_indent, os.path.basename(root)))
        for f in files:
            print ('{} {}'.format(file_indent, f))

def get_path(path):
    pah_lists = []
    for root, dirs, files in os.walk(path):
        for dir in dirs:
            pah_lists.append(os.path.join(root,dir))
        for file in files:
            if file not in [ "filepath.py", "filepath.bat", 'filepath.exe']:
                pah_lists.append(os.path.join(root,file))
    return pah_lists

def save_xls(path):
    filename = path +"\\file_path.xls"
    if os.path.exists(filename):
        os.remove(filename)
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('file_path')
    for index, value in enumerate(get_path(path)):
        worksheet.write(index, 0, value.decode('gbk'))
    workbook.save(filename)

def save_txt(path):
    filename = path +"\\file_path.txt"
    if os.path.exists(filename):
        os.remove(filename)
    with open(filename, 'w+') as file:
        for path in get_path(path):
            file.write(path)
            file.write('\n')

if __name__ == "__main__":
    path = os.getcwd()
    save_xls(path)
