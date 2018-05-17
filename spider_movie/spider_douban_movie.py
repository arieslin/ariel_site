# -*- coding:utf-8 -*-

"""

created by 2018-04-24
@author linwei

抓取的信息通过js生成，所有不能用bs4，使用的是抓包工具获取json数据的url，
使用了 requests，useragent，json，excel操作
"""

import requests
import xlwt
import json
import os
import sys

reload(sys)
sys.setdefaultencoding('utf8')

def get_movies(url, movie_list):
    headers = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) '
                                'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36'}
    reponse = requests.get(url, headers)
    reponse.encoding = reponse.apparent_encoding
    result = reponse.content
    result = json.loads(result)
    tmp_list = result['subjects']

    for movie in tmp_list:
        name = movie['title']
        score = movie['rate']
        movie_list.append([name, score])

def save_movies(list, file):
    list.sort(lambda x, y: cmp(y[1], x[1]))
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('movie')
    for row, movie in enumerate(list):
        for col, info in enumerate(movie):
            sheet.write(row, col, info)
            sheet.col(col).width = 256*20
    workbook.save(file)


def main():
    movie_list = [[u"豆瓣热门电影", u"豆瓣电影评分值"]]
    file = '豆瓣热门电影.xls'
    if os.path.exists(file):
        os.remove(file)
    base_url = 'https://movie.douban.com/j/search_subjects?' \
          'type=movie&tag=%E7%83%AD%E9%97%A8&sort=recommend&page_limit=20&page_start='
    depth = 5
    for i in range(6):
        url = base_url + str(20*i)
        print url
        get_movies(url, movie_list)
    save_movies(movie_list, file)


main()
