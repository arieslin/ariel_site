# -*- coding:utf-8 -*-

"""
created by 2018-04-29
@author linwei

爬取知乎的用户信息，session模拟登陆
使用了reuqests库，随机延时来模拟模拟器操作，soup操作，excel和样式，
"""
from selenium import webdriver
import requests
import time

def login_zhihu_sel():
    url = 'https://www.zhihu.com/signup?next=%2F'
    browser = webdriver.Chrome()
    browser.get(url)

def login_zhihu():
    s = requests.session()
    url = "http://www.segmentfault.com/user/login?_"
    #url = "https://segmentfault.com/api/user/login?_=f66d72e4946bc83de8a70718ef25de86"
    login_data = {'remember':'1', 'username':'814941978@qq.com', 'password':'lW814941978'}
    headers = {'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) '
                                'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.167 Safari/537.36'}
    rep = s.post(url, data=login_data, headers=headers)
    time.sleep(5)
    print rep.status_code
    r = s.get('https://segmentfault.com/newest')
    print r.status_code
    print r.text

login_zhihu_sel()

