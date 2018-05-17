# -*- coding:utf-8 -*-

"""
created by 2018-04-19
@author linwei

使用了requests，bs4，文件操作
"""

import requests
import bs4
from bs4 import BeautifulSoup
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')

def get_html(url):
    try:
        r = requests.get(url, timeout = 30)
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        print "spide url failed!"

def parser_html(content, univ_list):
    soup = BeautifulSoup(content, "html.parser")
    trs = soup.find('tbody').children
    univ_list.append([u"序号", u"学校", u"省市", u"总分", u"指标得分"])
    for tr in trs:
        if isinstance(tr, bs4.element.Tag):
            tds = tr.find_all('td')
            univ_list.append([tds[0].string, tds[1].string, tds[2].string, tds[3].string, tds[4].string])

def save_university_info(list, file):
    if os.path.exists(file):
        os.remove(file)
    with open(file, "wa+") as f:
        for info in list:
            print info
            str = '\t'.join(info)
            print str
            f.write(str + '\n')
    f.close()

def main():
    url = "http://www.zuihaodaxue.com/zuihaodaxuepaiming2016.html"
    univ_list = []
    file = "university.txt"

    content = get_html(url)
    parser_html(content, univ_list)
    save_university_info(univ_list, file)

if __name__ == '__main__':
    main()