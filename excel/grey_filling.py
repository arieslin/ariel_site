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

def fill_Lateral_Longitudinal(list, sheet, style, resolution = 'Lateral Resolution'):
    depth_index = 2
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '' and map[resolution] != '':
            largest_depth = map[resolution].split(u')')[-2].split(u'≤')[-1]
            test_depths = generate_test_depth(largest_depth)
            width = len(test_depths)
            sheet.write_merge(depth_index, depth_index + (width - 1), 0, 0, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 1, 1, map['Frequency'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 2, 3, map[resolution], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 4, 4, None, style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 13, 13, None, style)
            for j in range(width):
                sheet.write(depth_index + j, 5, test_depths[j], style)
            depth_index = depth_index + width
            for x in range(2, depth_index):
                for y in range(6, 14):
                    sheet.write(x, y, None, style)

#填充tc-53974    excel函数ok
def fill_sheet0(list, sheet, style):
    for index ,map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '' and map['Blind Area'] != '':
            for i in range(0,11):
                sheet.write(2+index, i, None, style)
                sheet.write(2+index, 0, map['Probe'], style)
                sheet.write(2+index, 1, map['Frequency'], style)
                sheet.write(2+index, 2, (u'≤') + str(int(map['Blind Area'])), style)
                sheet.write(2+index, 10, Formula('IF(AND(J{0}>0,J{0}<=INT(RIGHT(C{0}))),"Pass","Fail")'.format(3 + index)), style)

#tc-53975 excel函数fail
def fill_sheet1(list, sheet, style):
    fill_Lateral_Longitudinal(list, sheet, style)

#_tc-53976 excel函数fail
def fill_sheet2(list, sheet, style):
    fill_Lateral_Longitudinal(list, sheet, style, resolution= 'Longitudinal Resolution')

#_tc-53977  excel函数ok
def fill_sheet3(list, sheet, style):
    for index ,map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '' and map['Detect Depth'] != '':
            for i in range(0,10):
                sheet.write(2+index, i, None, style)
                sheet.write(2+index, 0, map['Probe'], style)
                sheet.write(2+index, 1, map['Frequency'], style)
                sheet.write(2+index, 2, (u'≥') + str(int(map['Detect Depth'])), style)
            sheet.write(2+index, 10, Formula('IF(J{0}>=INT(RIGHT(C{0},LEN(C{0})-1)),"Pass","Fail")'.format(3 + index)), style)

#tc-53978
def fill_sheet4(list, sheet, style):
    depth_index = 2
    width = 5
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '' and map['Expected Accuracy'] != '':
            sheet.write_merge(depth_index, depth_index + (width - 1), 0, 0, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 1, 1, map['Frequency'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 2, 2, map['Expected Accuracy'], style)
            for i in range(3, 9):
                sheet.write_merge(depth_index, depth_index + (width - 1), i, i, None, style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 10, 10, Formula('MAX(ABS(J{0}/20-1),ABS(J{1}/20-1),ABS(J{2}/20-1),ABS(J{3}/20-1),ABS(J{4}/20-1))'
                                                                                      .format(depth_index+1,depth_index+2,depth_index+3,depth_index+4,depth_index+5)), style)
            for i in range(11, 17):
                sheet.write_merge(depth_index, depth_index + (width - 1), i, i, None, style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 18, 18, Formula('MAX(ABS(R{0}/20-1),ABS(R{1}/20-1),ABS(R{2}/20-1),ABS(R{3}/20-1),ABS(R{4}/20-1))'
                                                                                      .format(depth_index+1,depth_index+2,depth_index+3,depth_index+4,depth_index+5)), style)
            sheet.write_merge(depth_index,depth_index + (width - 1), 19, 19, Formula('IF(AND(K{0}<=3%,S{0}<=3%),"Pass","Fail")'.format(depth_index+1)), style)
            depth_index = depth_index + width
            for x in range(2, depth_index):
                sheet.write(x, 9, None, style)
                sheet.write(x, 17, None, style)

#tc-53980
def fill_sheet5(list, sheet, style):
    depth_index = 2
    width = 3
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '' \
            and map['Slice thickness'] != '' and map['Detect Depth'] != '':
            sheet.write_merge(depth_index, depth_index + (width - 1), 0, 0, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 1, 1, map['Frequency'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 2, 2, map['Slice thickness'], style)
            sheet.write_merge(depth_index, depth_index + (width - 1), 3, 3, ((u'≥') + str(int(map['Detect Depth']))), style)
            for i in range(width):
                sheet.write(depth_index + i, 4, '70', style)
                sheet.write(depth_index + i, 14, Formula('N{0}/TAN((70/180)*PI())'.format(depth_index + 1 + i)), style)
                if i == 0:
                    sheet.write(depth_index + i, 5, 'd/3', style)
                    sheet.write(depth_index + i, 6, Formula('RIGHT(D{0},LEN(D{0})-1)/3'.format(depth_index + 1)), style)
                elif i == 1:
                    sheet.write(depth_index + i, 5, 'd/2', style)
                    sheet.write(depth_index + i, 6, Formula('RIGHT(D{0},LEN(D{0})-1)/2'.format(depth_index + 1)), style)
                else:
                    sheet.write(depth_index + i, 5, '2d/3', style)
                    sheet.write(depth_index + i, 6, Formula('RIGHT(D{0},LEN(D{0})-1)/3*2'.format(depth_index + 1)), style)
                for j in range(7, 13):
                    sheet.write(depth_index + i, j, None, style)
                    sheet.write(depth_index + i, 13, None, style)

            sheet.write_merge(depth_index, depth_index + (width - 1), 15, 15, Formula('IF(AND(O{0}>0,O{0}<=(INT(RIGHT(C{0},LEN(C{0})-1))),'
                                                                                  'O{1}>0,O{1}<=(INT(RIGHT(C{0},LEN(C{0})-1))),'
                                                                                  'O{2}>0,O{2}<=(INT(RIGHT(C{0},LEN(C{0})-1)))),"Pass","Fail")'.format(depth_index+1, depth_index+2, depth_index+3)), style)
            depth_index = depth_index + width

#_tc-53981
def fill_sheet6(list, sheet, style):
    depth_index = 1
    width = 5
    templet_item = ['Theory value(mm)', 'Test value(mm)', 'absolute value(mm)', 'relative deviation(%)', 'conclusion']
    col_list = 'D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z'.split(',')
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Detect Depth'] != '' and map['Expected Distance'] != '':
            sheet.write_merge(depth_index, depth_index + (width - 1), 1, 1, map['Probe'], style)
            for i in range(width):
                sheet.write(depth_index + i, 2, templet_item[i], style)
            detect_list = range(10, int(map['Detect Depth']), 10)
            for j, value in enumerate(detect_list):

                for col in range(width - 1):
                    if col == 0:
                        sheet.write(depth_index + col, 3 + j, value, style)
                    elif col == 3:
                        print 'col:', col_list[j]
                        sheet.write(depth_index + col, 3 + j, Formula('{0}{1}/{0}{2}'.format(col_list[j],depth_index+3, depth_index+1)), style)
                    elif col == 2:
                        sheet.write(depth_index + col, 3 + j, Formula('ABS({0}{1}-{0}{2})'.format(col_list[j], depth_index+2, depth_index+1)), style)
                    else:
                        sheet.write(depth_index + col, 3 + j, None, style)
            sheet.write_merge(depth_index + width - 1, depth_index + width - 1, 3, 2 + len(detect_list), None, style)
            depth_index = depth_index + width
    sheet.write_merge(1, depth_index - 1, 0, 0, str(map['Expected Distance']*100) + u'%', style)

#_tc-53984
def fill_sheet7(list, sheet, style):
    depth_index = 3
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Expected Area'] != '' and map['Expected Circumference'] != '':
            if 'L' in map['Probe']:
                width = 2
                theory_value = [['4.15', '7.23']]
            else:
                width = 4
                theory_value = [['4.15', '7.23'], ['14.52', '13.51']]
            sheet.write_merge(depth_index, depth_index+width-1, 2, 2, map['Probe'], style)
            for row, row_value in enumerate(theory_value):
                for col, col_value in enumerate(row_value):
                    sheet.write_merge(depth_index+row*2, depth_index+row*2+1, col+3, col+3, col_value, style)
                for i in range(5, 12):
                    sheet.write_merge(depth_index+row*2, depth_index+row*2+1, i, i, None, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 13, 13, Formula('MAX(ABS(M{0}/D{0}-1),ABS(M{1}/D{0}-1))'.format(depth_index+row*2+1,depth_index+row*2+2)), style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 15, 15, Formula('MAX(ABS(O{0}/D{0}-1),ABS(O{1}/D{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)),style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 17, 17, Formula('MAX(ABS(Q{0}/E{0}-1),ABS(Q{1}/E{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)),style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 19, 19, Formula( 'MAX(ABS(S{0}/E{0}-1),ABS(S{1}/E{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)), style)
            for row in range(width):
                for col in range(12,19,2):
                    sheet.write(depth_index+row, col, None, style)
            sheet.write_merge(depth_index, depth_index + width - 1, 20, 20, None, style)
            depth_index = depth_index + width
    sheet.write_merge(3, depth_index-1, 0, 0, str(map['Expected Area']*100) + u'%', style)
    sheet.write_merge(3, depth_index - 1, 1, 1, str(map['Expected Circumference'] * 100) + u'%', style)

#tc_53985
def fill_sheet8(list, sheet, style):
    depth_index = 2
    width = 4
    theory_value = ['45', '90']
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Expected Angle'] != '':
            sheet.write_merge(depth_index, depth_index + width - 1, 1, 1, map['Probe'], style)
            for row, value in enumerate(theory_value):
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 2, 2, value, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 11, 11, Formula('MAX(ABS(K{0}/C{0}-1),ABS(K{1}/C{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)), style)
            for col in range(3, 10):
                sheet.write_merge(depth_index, depth_index + width - 1, col, col, None, style)
            for row in range(width):
                sheet.write(depth_index + row, 10, None, style)
            sheet.write_merge(depth_index, depth_index + width - 1, 12, 12, Formula('IF(AND(L{0}<=$A$3,L{1}<=$A$3),"Pass","Fail")'.format(depth_index+1, depth_index+3)), style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, str(map['Expected Angle'] * 100) + u'%', style)


#tc-53986
def fill_sheet9(list, sheet, style):
    depth_index = 2
    width = 4
    instruction = u"Range: the probe under 4MHz, and equal to 5 MHz and 7.5 MHz correspond with the cystic lesions of 10 mm, 6 mm and 4 mm respectively.Deviation: within +/-10％ on the traverse and the longitudinal.  "
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Frequency'] != '':
            if 'L' in map['Probe']:
                theory_value = ['6', '4']
            else:
                theory_value = ['10', '6']
            sheet.write_merge(depth_index, depth_index+width-1, 1, 1, map['Probe'], style)
            sheet.write_merge(depth_index, depth_index + width - 1, 2, 2, map['Frequency'], style)
            for row, value in enumerate(theory_value):
                sheet.write_merge(depth_index+row*2, depth_index+row*2+1, 3, 3, value, style)
                for col in range(4, 10):
                    sheet.write_merge(depth_index+row*2, depth_index+row*2+1, col, col, None, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 12, 12, Formula('MAX(ABS(K{0}/D{0}-1),ABS(K{1}/D{0}-1))'.format(depth_index+row*2+1, depth_index+row*2+2)), style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 13, 13, Formula('MAX(ABS(L{0}/D{0}-1),ABS(L{1}/D{0}-1))'.format(depth_index+row*2+1, depth_index+row*2+2)), style)
            sheet.write_merge(depth_index, depth_index+width-1, 14, 14, Formula('IF(AND(M{0}<=10%,N{0}<=10%,M{1}<=10%,N{1}<=10%),"Pass","Fail")'.format(depth_index+1,depth_index+3)), style)
            for i in range(width):
                for j in range(10,12):
                    sheet.write(depth_index+i, j, None, style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, instruction, style)

#tc-53991
def fill_sheet10(list, sheet, style):
    depth_index = 2
    width = 6
    theory_value = ['20', '30', '40']
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Expected Distance'] != '':
            sheet.write_merge(depth_index, depth_index + width - 1, 1, 1, map['Probe'], style)
            for row, value in enumerate(theory_value):
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 2, 2, value, style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 11, 11, Formula('MAX(ABS(K{0}/C{0}-1),ABS(K{1}/C{0}-1))'.format(depth_index + row * 2 + 1,depth_index + row * 2 + 2)), style)
            for col in range(3, 10):
                sheet.write_merge(depth_index, depth_index + width - 1, col, col, None, style)
            for row in range(width):
                sheet.write(depth_index + row, 10, None, style)
            sheet.write_merge(depth_index, depth_index + width - 1, 12, 12, Formula('IF(AND(L{0}<=$A$3,L{1}<=$A$3,L{2}<=$A$3),"Pass","Fail")'.format(depth_index+1, depth_index+3, depth_index+5)), style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, str(map['Expected Distance'] * 100) + u'%', style)


#_tc-53987
def fill_sheet11(list, sheet, style):
    depth_index = 2
    width = 4
    theory_value = ['Base Frequency', 'Harmonic']
    for index, map in enumerate(list):
        if map['Probe'] != '' and map['Expected Time'] != '' and map['Expected Heart Rate']:
            sheet.write_merge(depth_index, depth_index + width - 1, 2, 2, map['Probe'], style)
            for row, value in enumerate(theory_value):
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 3, 3, value, style)
                sheet.write(depth_index + row * 2, 4, '2', style)
                sheet.write(depth_index + row * 2 + 1, 4,  '3', style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 5, 5, '500', style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 6, 6, '240', style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 9, 9, Formula('MAX(ABS(H{0}/F{0}-1),ABS(H{1}/F{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)), style)
                sheet.write_merge(depth_index + row * 2, depth_index + row * 2 + 1, 10, 10, Formula( 'MAX(ABS(I{0}/G{0}-1),ABS(I{1}/G{0}-1))'.format(depth_index + row * 2 + 1, depth_index + row * 2 + 2)),style)
            sheet.write_merge(depth_index, depth_index + width - 1, 11, 11, Formula('IF(AND(J{0}<=$A$3,J{1}<=$A$3,K{0}<=$B$3,K{1}<=$B$3),"Pass","Fail")'.format(depth_index+1, depth_index+3)), style)
            for row in range(width):
                sheet.write(depth_index + row, 7, None, style)
                sheet.write(depth_index + row, 8, None, style)
            depth_index = depth_index + width
    sheet.write_merge(2, depth_index - 1, 0, 0, str(map['Expected Time'] * 100) + u'%', style)
    sheet.write_merge(2, depth_index - 1, 1, 1, str(map['Expected Heart Rate'] * 100) + u'%', style)

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

