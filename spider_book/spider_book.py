# -*- coding:utf-8 -*-

"""
created by 2018-04-26
@author linwei

使用了reuqests库，随机延时来模拟模拟器操作，soup操作，excel和样式，
"""

import requests
import bs4
from bs4 import BeautifulSoup
import sys
import os
import xlwt
import time
import numpy

reload(sys)
sys.setdefaultencoding('utf8')

def get_books(url):
    try:
        r = requests.get(url, timeout = 30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        print "spide url failed!"

def parser_books(content, book_list):
    soup = BeautifulSoup(content, "html.parser")
    books = soup.find('div', attrs={'class':'indent'})
    tables = books.find_all('table')
    for table in tables:
        english_name_tmp = table.find('div', {'class':'pl2'}).find('span')
        if isinstance(english_name_tmp, bs4.element.Tag):
            english_name = english_name_tmp.string
            print  english_name
        else:
            english_name = ''

        chinese_name_tmp = table.find('div', {'class':'pl2'}).find('a')
        if isinstance(chinese_name_tmp, bs4.element.Tag):
            chinese_name = chinese_name_tmp.attrs.get('title')
            print  chinese_name
        else:
            chinese_name = ''

        score = table.find('span', attrs={'class':'rating_nums'}).string
        auther = table.find('p', attrs={'class':'pl'}).string.split('/')[0]
        sumary = table.find('span', attrs={'class':'inq'}).string
        book_list.append([score,chinese_name,english_name,auther,sumary])


def save_books(book_list_tmp, file, style):
    book_list = [[u'中文名', u'英文名', u'评分', u'作者', u'简介']]
    book_list_tmp.sort(reverse=True)
    for book in book_list_tmp:
        book_list.append([book[1],book[2],book[0],book[3],book[4]])
    if os.path.exists(file):
        os.remove(file)
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('book')
    for row, book in enumerate(book_list):
        for col, info in enumerate(book):
            sheet.write(row, col, info, style)
            sheet.col(col).width = 256 * 30
    workbook.save(file)

def set_style():
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER
    align.vert = xlwt.Alignment.VERT_CENTER
    style = xlwt.XFStyle()
    style.alignment = align
    return style

def main():
    depth = 5
    book_list_tmp = []
    file = '豆瓣图书top250.xls'
    base_url = "https://book.douban.com/top250?start="
    style = set_style()
    for i in range(depth):
        time.sleep(numpy.random.rand()*5)
        url = base_url + str(i*25)
        book_info = get_books(url)
        parser_books(book_info, book_list_tmp)
    save_books(book_list_tmp, file, style)
main()
