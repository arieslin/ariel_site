# -*- coding:utf-8 -*-
import xlrd
from xlwt import *
from xlutils import copy
import os

#
def parse_detect_value(str):
    list= []
    list.append(str.split(',')[0].split(':'))
    list.append(str.split(',')[1].split(':'))
    print list
    return list

#tc-53997
def fill_sheet16(list, sheet, style):
    depth_index = 2
    width = 4
    theory_value = [['500', '240'], ['500', '240']]
    scan_speed = ['2', '3']
    for index, map in enumerate(list):
        if map['PW detect depth'] != '' and map['Probe'] != '':
            sheet.write_merge(depth_index, depth_index + width - 1, 2, 2, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + width - 1, 11, 11, Formula('IF(AND(J{0}<=$A$3,J{1}<=$A$3,K{0}<=$B$3,K{1}<=$B$3),"Pass","Fail")'
                                                                                    .format(depth_index + 1, depth_index + 3)), style)
            pw_freq = parse_detect_value(map['PW detect depth'])
            for row, freq in enumerate(pw_freq):
                sheet.write_merge(depth_index + row*2, depth_index + row*2 + width/2 - 1, 3, 3, freq[0], style)
                for i, speed in enumerate(scan_speed):
                    sheet.write(depth_index + i + row*2, 4, speed, style)
                    sheet.write(depth_index + i + row*2, 7, None, style)
                    sheet.write(depth_index + i + row*2, 8, None, style)
            for row, theory in enumerate(theory_value):
                for col, value in enumerate(theory):
                    sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + width / 2 - 1, 5 + col, 5 + col, value, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + width / 2 - 1, 9, 9, Formula('MAX(ABS(H{0}/F{0}-1),ABS(H{1}/F{0}-1))'
                                                                                                              .format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)), style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + width / 2 - 1, 10, 10, Formula('MAX(ABS(I{0}/G{0}-1),ABS(I{1}/G{0}-1))'
                                                                                                                .format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)), style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, str(map['Expected Time'] * 100) + u'%', style)
    sheet.write_merge(2, depth_index - 1, 1, 1, str(map['Expected Heart Rate'] * 100) + u'%', style)


