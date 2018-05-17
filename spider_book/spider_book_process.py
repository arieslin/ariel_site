# -*- coding:utf-8 -*-

"""
created by 2018-04-26
@author linwei

使用了requests，bs4，multiprocessing 多进程
"""

import requests
import bs4
from bs4 import BeautifulSoup
import sys
import os
import xlwt
import multiprocessing


#需要将所有的数据得到后，再存入excel或txt文件内。因为多进程下同时打开同一个文件可能会有问题。
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
        chinese_name_tmp = table.find('div', {'class':'pl2'}).find('a')
        if isinstance(chinese_name_tmp, bs4.element.Tag):
            chinese_name = chinese_name_tmp.attrs.get('title').strip()
            print  chinese_name
        else:
            chinese_name = ''

        score = table.find('span', attrs={'class':'rating_nums'}).string.strip()
        auther = table.find('p', attrs={'class':'pl'}).string.split('/')[0].strip()
        sumary = table.find('span', attrs={'class':'inq'}).string.strip()
        book_list.append([chinese_name,score,auther,sumary])

def save_book_info(list, file):
    with open(file, "a+") as f:
        for info in list:
            print info
            tplt = "{0:^20}\t{1:^20}\t{2:^20}\t{3:^20}"
            str = tplt.format(info[0],info[1],info[2],info[3])
            print str
            f.write(str + '\n')
    f.close()


def main(i):
    base_url = "https://book.douban.com/top250?start="
    url = base_url + str(i*25)
    book_info = get_books(url)
    parser_books(book_info, book_list)
    save_book_info(book_list, file_tmp)

if __name__ == '__main__':
    book_list_tmp = [[u'中文名', u'评分', u'作者', u'简介']]
    book_list = []
    depth = 5
    file_tmp = '豆瓣图书top250.txt'
    if os.path.exists(file_tmp):
        os.remove(file_tmp)
    save_book_info(book_list_tmp, file_tmp)
    for i in range(depth):
        process = multiprocessing.Process(target=main, args=(i,))
        process.start()
