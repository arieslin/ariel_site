# -*- coding:utf-8 -*-

import sys
import os
import xlwt
import pdb

reload(sys)
sys.setdefaultencoding('utf8')
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('linwei')
sheet.write(0,0,'A')
for i in range(10):
    pdb.set_trace()
    sheet.write(1,i,i)
workbook.save('/Users/arieslin/python/test/xlwt.xls')

def test():
    book_list_tmp = [[u'中文名', u'英文名', u'评分', u'作者', u'简介']]
    tplt = "{0:^10}\t{1:^10}\t{2:^10}\t{3:^10}\t{4:^10}"
    for info in book_list_tmp:
        str = tplt.format(info[0], info[1], info[2], info[3], info[4])
    print str

test()