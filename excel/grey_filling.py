# -*- coding:utf-8 -*-
import xlrd
import xlwt
from xlwt import *
import xlutils
from xlutils import copy
import os

def generate_test_depth(depth):
    depth = int(depth)
    depth_select = [10, 30, 50, 70, 120, 160, 200]

    if depth <= 40:
        return depth_select[0], depth_select[1]
    elif 40 < depth <= 60:
        return depth_select[0], depth_select[1], depth_select[2]
    elif 60 < depth <= 110:
        return depth_select[1], depth_select[2], depth_select[3]
    elif 110 < depth <= 150:
        return depth_select[1], depth_select[2], depth_select[3], depth_select[4]
    elif 150 < depth <= 190:
        return depth_select[1], depth_select[2], depth_select[3], depth_select[4], depth_select[5]
    elif 190 < depth <= 240:
        return depth_select[1], depth_select[2], depth_select[3], depth_select[4], depth_select[5], depth_select[6]
generate_test_depth(80)



#tc-53992
def fill_sheet12(list, sheet, style):
    depth_index = 2
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Expected Volume'] != '':
            if 'L' in map['Probe']:
                theory_value = ['8.7 ']
                width = 2
                sheet.write_merge(depth_index, depth_index + width - 1, 12, 12, Formula(
                    'IF(AND(L{0}<=$A$3),"Pass","Fail")'.format(depth_index + 1, depth_index + 3, )), style)
            else:
                theory_value = ['8.7 ', '63.3']
                width = 4
                sheet.write_merge(depth_index, depth_index + width - 1, 12, 12, Formula('IF(AND(L{0}<=$A$3,L{1}<=$A$3),"Pass","Fail")'.format(depth_index + 1, depth_index + 3, )), style)
            sheet.write_merge(depth_index, depth_index + width - 1, 1, 1, map['Probe'], style)
            for row, value in enumerate(theory_value):
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 2, 2, value, style)
                for col in range(3, 10):
                    sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, col, col, None, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 11, 11, Formula('MAX(ABS(K{0}/C{0}-1),ABS(K{1}/C{0}-1))'.format(depth_index+row*2+1, depth_index+row*2+2)), style)
            for row in range(width):
                sheet.write(depth_index+row, 10, None, style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, str(map['Expected Volume'] * 100) + u'%', style)

