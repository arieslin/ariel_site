# -*- coding:utf-8 -*-
'''
parse xml and get message list, then save to the excel
@author linwei
crreated on 2018-5-25
'''
import xml.dom.minidom as xmldom
import os
import xlwt
'''
<?xml version="1.0" encoding="utf-8"?>
<paralib>
<imgmode mode="a">
	<class name="aaaa">
		<paraName>adfds</paraName>
		<modeId>fasdflkdj</modeId>
		<messageId>dsfasldkfj</messageId>
		<dep></dep>	
	</class>	
</imgmode>
'''

def get_nodes(node, name):
    return node.getElementsByTagName(name) if node else []
def get_attribute(node, name):
    return node.getAttribute(name) if node else ''
def get_node_value(node, index = 0):
    if node:
        childNodes = node.childNodes
        return childNodes[index].nodeValue if childNodes else ''
    else:
        return ''
    return node.childNodes[index].nodeValue if node else ''
def get_nodes_obj(file):
    dom_obj = xmldom.parse(file)
    return dom_obj.documentElement
def parse_xml(file):
    elements = get_nodes_obj(file)
    imagmodes = get_nodes(elements, 'imgmode')
    for imagmode in imagmodes:
        imagemode_attri = get_attribute(imagmode, 'mode')
        categorys = get_nodes(imagmode, 'class')
        for category in categorys:
            category_name = get_attribute(category, 'name')
            paraName = get_node_value(get_nodes(category, 'paraName')[0])
            modeId = get_node_value(get_nodes(category, 'modeId')[0])
            messageId = get_node_value(get_nodes(category, 'messageId')[0])
            dep = get_node_value(get_nodes(category, 'dep')[0])
            print u'图像模式：' + imagemode_attri
            print u'POD参数：' + category_name
            print u'POD参数属性值：' + paraName + '\t' + modeId + '\t' + messageId + '\t' + \
                  dep
            para_list.append([imagemode_attri, category_name, paraName, modeId, messageId, dep])

def parse_msg_lists(xml_path):
    elements = get_nodes_obj(xml_path)
    imagmodes = get_nodes(elements, 'imgmode')
    for imagmode in imagmodes:
        imagemode_attri = get_attribute(imagmode, 'mode')
        categorys = get_nodes(imagmode, 'class')
        for category in categorys:
            category_name = get_attribute(category, 'name')
            paraName = get_node_value(get_nodes(category, 'paraName')[0]).split(',')
            modeId = get_node_value(get_nodes(category, 'modeId')[0])
            messageId = get_node_value(get_nodes(category, 'messageId')[0])
            dep = get_node_value(get_nodes(category, 'dep')[0])
            msg_tmp = []
            for msg in messageId.split(','):
                msg_tmp.append(imagemode_attri + u':' + msg)
            for para in paraName:
                para_list.append([para, ';'.join(msg_tmp)])
    return para_list

def save_excel():
    filename = "messageid_data.xls"
    if os.path.exists(filename):
        os.remove(filename)
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("messageid data")
    head = ['imagemode', 'classname', 'paraname', 'modeid', 'messageid', 'dep']
    for col, value in enumerate(head):
        ws.write(0, col, value)
    for row, message in enumerate(para_list):
        for col, para in enumerate(message):
            ws.col(col).width = 256 * 30
            ws.write(row+1, col, para)

    wb.save(filename)

if __name__ == '__main__':
    file = 'paralib.xml'
    para_list = []
    parse_msg_lists(file)
    print para_list
    save_excel()
