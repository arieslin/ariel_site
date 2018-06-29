# -*- coding:utf-8 -*-
import xlrd
from xlwt import *
from xlutils import copy
import os
from grey_filling import *
from doppler_filling import *


#解析出每个探头的精度指标的列表
def get_probe_info(file, list):
    workbook = xlrd.open_workbook(file)
    original = workbook.sheet_by_name(u"Sheet1")
    for row in range(1,original.nrows):
        info_map = {}
        for col in  range(original.ncols):
            info_map.update({original.row_values(0)[col] : original.row_values(row)[col]})
        list.append(info_map)

#对每个sheet页面进行填充。
def fill_probe_info(list, module_path, result_path):
    workbook = xlrd.open_workbook(module_path, formatting_info=True)
    sheets = workbook.nsheets
    print sheets
    style = set_style()
    wtbook = copy.copy(workbook)
    for i in range(30):
        sheet_index = wtbook.get_sheet(i)
        fun_name = 'fill_sheet' + str(i)
        eval(fun_name)(list, sheet_index, style)
    wtbook.save(result_path)

#构造精度表格的样式
def set_style():
    borders = Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    al = Alignment()
    al.horz = Alignment.HORZ_CENTER
    al.vert = Alignment.VERT_CENTER

    style = XFStyle()
    style.borders = borders
    style.alignment = al
    return  style

def main():
    original_path = "original.xls"
    probe_info_list = []
    module_path = "module.xls"

    result_path = "result.xls"
    if os.path.exists(result_path):
        os.remove(result_path)

    get_probe_info(original_path, probe_info_list)
    print probe_info_list
    fill_probe_info(probe_info_list, module_path, result_path)

main()