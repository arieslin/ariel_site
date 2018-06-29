# -*- coding:utf-8 -*-
import xlrd
from xlwt import *
from xlutils import copy
import os

#解析PW或C的频点和探测深度
def parse_detect_value(str):
    list= []
    list.append(str.split(',')[0].split(':'))
    list.append(str.split(',')[1].split(':'))
    print list
    return list

#填充血流方向准确性测试_tc-53998 or tc-54009 or tc-54013
def fill_blood_direct(list, sheet, style, theory_value, flag):
    depth_index = 1
    for index, map in enumerate(list):
        if flag == 'pw':
            freq = parse_detect_value(map['PW detect depth'])[0][0]
        elif flag == 'c':
            freq = parse_detect_value(map['C detect depth'])[0][0]
        elif flag == 'p':
            freq = parse_detect_value(map['P detect depth'])[0][0]
        if freq != '' and map['Probe'] != '':
            sheet.write(depth_index, 1, map['Probe'], style)
            sheet.write(depth_index, 2, freq, style)
            sheet.write_merge(depth_index, depth_index, 3, 6, theory_value, style)
            sheet.write(depth_index, 7, None, style)
            depth_index = depth_index + 1
    sheet.write_merge(1, depth_index - 1, 0, 0, theory_value, style)

#填充灵敏度测试_tc-54001 or tc-54011 or tc-54014
def fill_sensitivity(list, sheet, style, row_max, flag):
    depth_index = 2
    width = 6
    theory_value = [['3.97', '1', '1.97'], ['3.97', '30', '59.10'], ['3.97', '60', '118.48']]
    for index, map in enumerate(list):
        sheet.write_merge(depth_index, depth_index + width - 1, 0, 0, map['Probe'], style)
        if flag == 'pw':
            freq = parse_detect_value(map['PW detect depth'])
        elif flag == 'c':
            freq = parse_detect_value(map['C detect depth'])
        elif flag == 'p':
            freq = parse_detect_value(map['P detect depth'])
        if freq != '':
            for row, pw_freq in enumerate(freq):
                for col, value in enumerate(pw_freq):
                    sheet.write_merge(depth_index + row*3, depth_index + row*3 + width/2 - 1, col + 1, col + 1, value, style)
                for row_theo, theory in enumerate(theory_value):
                    for col_theo, value_theo in enumerate(theory):
                        sheet.write(depth_index + row_theo + row*3, col_theo + 3, value_theo, style)
                    for col in range(6, row_max-1):
                        sheet.write(depth_index + row_theo + row*3, col, None, style)
                    sheet.write(depth_index + row_theo + row*3, row_max-1, Formula('0.5*2*K{0}*B{1}'.format(depth_index + row_theo + row*3 + 1, depth_index + row*3 + 1)), style)   #缺少函数
                sheet.write_merge(depth_index + row * 3, depth_index + row * 3 + width / 2 - 1, row_max, row_max, Formula('IF(AND(K{0}>=C{0},K{1}>=C{0},K{2}>=C{0}),"Pass","Fail")'
                                                                                                                          .format(depth_index + row*3 + 1, depth_index + row*3 + 2, depth_index + row*3 + 3)), style)  #缺少函数
            depth_index = depth_index + width

#填充血流速度范围测试_tc-54002 or tc-54012
def fill_blood_range(list, sheet, style):
    depth_index = 2
    for index, map in enumerate(list):
        sheet.write(depth_index + index, 0, map['Probe'], style)
        for col in range(1, 3):
            sheet.write(depth_index + index, col, None, style)
    sheet.write(depth_index + len(list), 0, "Conclusion", style)
    sheet.write_merge(depth_index + len(list), depth_index + len(list), 1, 2, None, style)

#填充取样容积位置准确性测试_tc-53999 or tc_54010
def fill_blood_deviation(list, sheet, style, flag):
    depth_index = 1
    width = 2
    theory = 'Deviation:smaller than 1/4 of the vessel diameter (0.99 mm).'
    theory_value = u'±0.99'
    for index, map in enumerate(list):
        if flag == 'pw':
            freq = parse_detect_value(map['PW detect depth'])
        elif flag == 'c':
            freq = parse_detect_value(map['C detect depth'])
        if freq != '':
            for row, freq in enumerate(freq):
                sheet.write(depth_index + row, 2, freq[0], style)
                sheet.write_merge(depth_index + row, depth_index + row, 3, 4, theory_value, style)
                sheet.write(depth_index + row, 5, None, style)
            sheet.write_merge(depth_index, depth_index + width - 1, 1, 1, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + width - 1, 6, 6, Formula('IF(AND(AND(F{0}>=-0.99,F{0}<=0.99),AND(F{1}>=-0.99,F{1}<=0.99)),"Pass","Fail")'
                                                                                  .format(depth_index + 1, depth_index + 2)), style)
            depth_index = depth_index + width
    sheet.write_merge(1, depth_index - 1, 0, 0, theory, style)


#填充PW血流方向准确性测试_tc-53998
def fill_sheet13(list, sheet, style):
    theory_value = 'The spectrum line of 3.97 mm bionic vessel appears on the top of the baseline.' \
                   'The spectrum line of 7.95mm bionic vessel appears at the bottom of the baseline.'
    flag = 'pw'
    fill_blood_direct(list, sheet, style, theory_value, flag)

#填充PW流速准确性测试_tc-54000
def fill_sheet14(list, sheet, style):
    depth_index = 3
    width = 16
    theory_value = ['10', '50', '100', '200']
    theory_angle = ['-30', '-60']
    for index, map in enumerate(list):
        if map['PW detect depth'] != '' and map['Probe'] != '':
            sheet.write_merge(depth_index, depth_index + width - 1, 0, 0, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + width - 1, 15, 15, Formula('IF(AND(O{0}<=5%,O{1}<=5%),"Pass","Fail")'.format(depth_index+1, depth_index+width/2 + 1)), style)
            pw_freq = parse_detect_value(map['PW detect depth'])
            for row_fre, value in enumerate(pw_freq):
                sheet.write_merge(depth_index + row_fre * 8 , depth_index + row_fre * 8 + width/2 - 1, 1, 1, value[0], style)
                sheet.write_merge(depth_index + row_fre * 8, depth_index + row_fre * 8 + width / 2 - 1, 14, 14, Formula('MAX(N{0}:N{1})'.format(depth_index + row_fre * 8 + 1, depth_index + row_fre * 8 + width/2)), style)
                for row, value in enumerate(theory_value):
                    sheet.write_merge(depth_index + row * 2 + row_fre * 8, depth_index + row * 2 + row_fre * 8 + 1, 2, 2, value, style)
                    for row_angle, angle in enumerate(theory_angle):
                        sheet.write(depth_index + row_angle + row*4 + row_fre*2 , 3, angle, style)
                        for col in range(4, 12):
                            sheet.write(depth_index + row_angle + row*4 + row_fre*2, col, None, style)
                        sheet.write(depth_index + row_angle + row*4 + row_fre*2, 12, Formula('L{0}*1480/1540'.format(depth_index + row_angle + row*4 + row_fre*2 + 1)), style)
                        sheet.write(depth_index + row_angle + row * 4 + row_fre * 2, 13, Formula('ABS(M{0}/C{1}-1)'.format(depth_index + row_angle + row * 4 + row_fre * 2 + 1, depth_index + row * 4 + row_fre * 2 + 1)), style)
            depth_index = depth_index + width

#填充灵敏度测试_tc-54001
def fill_sheet15(list, sheet, style):
    flag = 'pw'
    row_max = 12
    fill_sensitivity(list, sheet, style, row_max, flag)

#填充PW时间准确性测试_tc-53997
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




#填充PW血流速度范围测试_tc-54002
def fill_sheet17(list, sheet, style):
    fill_blood_range(list, sheet, style)


#填充取样容积位置准确性测试_tc-53999
def fill_sheet18(list, sheet, style):
    flag = 'pw'
    fill_blood_deviation(list, sheet, style, flag)

#填充CW血流方向准确性测试_tc-54004
def fill_sheet19(list, sheet, style):
    pass

#填充CW流速准确性测试_tc-54005
def fill_sheet20(list, sheet, style):
    pass

#填充CW灵敏度测试_tc-54006
def fill_sheet21(list, sheet, style):
    pass

#填充CW时间准确性测试_tc-54003
def fill_sheet22(list, sheet, style):
    pass

#填充CW血流速度范围测试_tc-54007
def fill_sheet23(list, sheet, style):
    pass

#填充血流方向测试_tc-54009
def fill_sheet24(list, sheet, style):
    theory_value = 'The color of 3.97 mm bionic vessel is red;' \
                   'The color of 7.95 mm bionic vessel is blue;.'
    flag = 'c'
    fill_blood_direct(list, sheet, style, theory_value, flag)

#填充灵敏度_tc-54011
def fill_sheet25(list, sheet, style):
    flag = 'c'
    row_max = 14
    fill_sensitivity(list, sheet, style, row_max, flag)

#填充Color血流速度范围测试_tc-54012
def fill_sheet26(list, sheet, style):
    fill_blood_range(list, sheet, style)

#填充血管中心位置偏差_tc-54010
def fill_sheet27(list, sheet, style):
    flag = 'c'
    fill_blood_deviation(list, sheet, style, flag)

#填充DirPower血流方向测试_tc-54013
def fill_sheet28(list, sheet, style):
    theory_value = 'The color of 3.97 mm bionic vessel is red;' \
                   'The color of 7.95 mm bionic vessel is blue;.'
    flag = 'p'
    fill_blood_direct(list, sheet, style, theory_value, flag)

#填充灵敏度_tc-54014
def fill_sheet29(list, sheet, style):
    flag = 'p'
    row_max = 14
    fill_sensitivity(list, sheet, style, row_max, flag)

