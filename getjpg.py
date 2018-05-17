# -*- coding:utf-8 -*-
import urllib
import urllib2
import re
from bs4 import BeautifulSoup


class SpiderUserInfo:

    # 定义抓取的网址
    def __init__(self):
        self.url = "https://www.zhihu.com/people/"

    # 获取页面的html信息
    def get_user_info(self,userid):
        userinfo_url = self.url + userid

        userinfo_reponse = urllib2.Request(userinfo_url)
        userinfo = urllib2.urlopen(userinfo_reponse)
        return userinfo.read()

    # 获取页面里面的关注人的userid
    def get_all_userid(self, userid):
        all_userid = self.url + userid + '/followees'
        user_list_html = self.get_user_info(all_userid)

        user_list_pattern = re.compile('class="zg-link author-link"')
        user_lists = re.findall(user_list_pattern, user_list_html)
        users = []
        users.append(userid)
        for user_list in user_lists:
            users.append(user_list.split('\"')[4]).split('people/')[1]

        return users
    # 对html进行解析，找到想要的
    def resolve_user_info(self,userid):
        userhtml = self.get_user_info(userid)

        # 获取学校信息
        school_pattern = re.compile('<span class="education item" title=".*">')
        university_info = re.findall(school_pattern, userhtml)[0].split('\"')[3]
        print university_info
        # 获取居住地
        location_pattern = re.compile('<span class="location item" title=".*">')
        location_info = re.findall(location_pattern,userhtml)[0].split('\"')[3]
        print location_info
        endlist = []
        return endlist.append(userid, university_info, location_info)

    def get_allinfo(self,userid):
        lists = self.get_all_userid(userid)
        for list in lists:
            endlist = self.resolve_user_info(list)
            for info in endlist:
                print info[0], info[1], info[2]

spider1 = SpiderUserInfo()
print spider1.resolve_user_info('arieslin')