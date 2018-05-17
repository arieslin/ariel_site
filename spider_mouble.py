# -*- coding:utf-8 -*-

"""

created by 2018-04-17
@author linwei

"""

import requests
import os
#headers参数使用练习
def spider_headers():
    url = "https://www.amazon.cn/dp/B0794JYPBR?_encoding=UTF8&ref_=pc_cxrd_658399051_recTab_658399051_t_3&pf_rd_p=370ee273-6401-4eab-913a-eabcae8ccba9&pf_rd_s=merchandised-search-3&pf_rd_t=101&pf_rd_i=658399051&pf_rd_m=A1AJ19PSB66TGU&pf_rd_r=351FDRJE7W50DCY831BK&pf_rd_r=351FDRJE7W50DCY831BK&pf_rd_p=370ee273-6401-4eab-913a-eabcae8ccba9"
    try:
        kv = {'User-Agent' : 'Mozilla/5.0'}
        r = requests.get(url, headers = kv)
        print r.request.headers
        print r.status_code
        r.raise_for_status()
        r.encoding = r.apparent_encoding
        return r.text
    except:
        return "spider failed!"

#params 参数使用练习
def spider_params():
    params = {'wb':'python'}
    r = requests.get("http://www.baidu.com", params=params)
    print r.status_code
    print len(r.text)

#爬取图片练习
def spider_pic():
    url = "https://www.nationalgeographic.com/photography/proof/2018/04/sharks-groupers-photography-animals/#/sharks-22062017-DSC_3518.jpg"
    root = os.getcwd() + 'pic/'
    path = root + url.split("/")[-1]
    #path = root + "/linwei.jpg"
    try:
        if not os.path.exists(root):
            os.mkdir(root)
        if not os.path.exists(path):
            r = requests.get(url)
            print r.status_code
            with open(path, 'wb') as f:
                f.write(r.content)
                f.close()
                print "文件保存成功！"
        else:
            print '文件已存在'
    except:
        print "获取文件异常。"



spider_pic()


