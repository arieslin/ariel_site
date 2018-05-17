# -*- coding:utf-8 -*-

import bs4
from bs4 import BeautifulSoup
import xlwt
import xlrd
import sys
import chardet
import os

#如果文件已存在，就先删除
result = "html.xls"
if os.path.exists(result):
    os.remove(result)

#判断系统和文件的编码
path = "/Users/arieslin/python/htmlparse/x.xls.html"
if os.path.exists(path):
    fp = open(path, 'r')
    content = fp.read()
    content.decode('GB2312').encode('utf-8')
# print sys.getfilesystemencoding()
# print 'Html is encoding by : %',chardet.detect(content)

#解析xml，返回lists
rows = []
soup= BeautifulSoup(content,"html.parser")

for x in soup.findAll('tr'):
    row = []
    for y in x.findAll("td"):
        row.append(y.text )
    rows.append(row)

#按照模块类型进行汇总
listt = []
for row in rows:
    listt.append(row[11])

# listt = ['a','b','c','d','a','a','b','b','c']
map_module = {}
list_only = set(listt)
for i in list_only:
     map_module.update({i: listt.count(i)})


#存储lists到excel
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("result")
sheet_module = workbook.add_sheet("module")
for r_num, row in enumerate(rows):
    for c_num, col in enumerate(row):
        sheet.write(r_num ,c_num, col)

# import pdb; pdb.set_trace()
index = 1
for key, value in map_module.iteritems():
    sheet_module.write(index,0,key)
    sheet_module.write(index,1,value)
    index += 1

workbook.save(result)

