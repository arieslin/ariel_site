# -*- coding:utf-8 -*-
"""
crreated on 2018-6-22
@author: linwei
"""
import os
import csv
import xlrd
import xlwt


class Parser:
    def __init__(self, msgfile, funcfile, keyfile):
        self.msgfile = msgfile
        self.funcfile = funcfile
        self.keyfile = keyfile

    def parse_msg(self):
        if os.path.exists(self.msgfile):
            tmp_list = []
            mode_msg_list = []
            msg_list = []
            wb = xlrd.open_workbook(self.msgfile)
            file = wb.sheet_by_name(u'Sheet1')
            for row in range(1, file.nrows):
                row_value = file.row_values(row)
                front = row_value[1].lstrip().rstrip()
                end = row_value[2].lstrip().rstrip()
                tmp_list.append(front)
                tmp_list.append(end)

            for msg in tmp_list:
                if ';' in msg:
                    for index in msg.split(';'):
                        msg_list.append(index)
                else:
                    mode_msg_list.append(msg)
            for msg in mode_msg_list:
                if msg:
                    msg_list.append(msg)
            return msg_list

    def parse_func(self):
        if os.path.exists(self.funcfile):
            msg_func_dict = {}

            with open(self.funcfile, 'r') as file:
                for line in file:
                    if ',' in line:
                        func = line.split(',')[0].lstrip('{').rstrip()
                        msg = line.split(',')[4].lstrip('{').rstrip()
                        msg_func_dict[msg] = func
            return msg_func_dict

    def parse_key(self):
        if os.path.exists(self.keyfile):
            func_key_dict = {}
            wb = xlrd.open_workbook(self.keyfile)
            file = wb.sheet_by_name(u'Sheet1')

            for row in range(1, file.nrows):
                row_value = file.row_values(row)
                func = row_value[0].lstrip().rstrip()
                key = row_value[1].lstrip().rstrip()
                func_key_dict[func] = key
            return func_key_dict


class Saver:
    def __init__(self, filename, head, content):
        self.filename = filename
        self.head = head
        self.content = content

    def save_csv(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)

        with open(self.filename, 'wb+') as file:
            writer = csv.writer(file)
            writer.writerow(self.head)
            for pod in self.content:
                writer.writerow(pod)

    def save_xls(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('msg_key')

        for col, value in enumerate(self.head):
            ws.write(0, col, value)

        for row, message in enumerate(self.content):
            for col, para in enumerate(message):
                ws.col(col).width = 256 * 30
                ws.write(row + 1, col, para)

        wb.save(self.filename)

def parse_config(file):

    with open(file, 'r') as file:
        for line in file:
            if 'msgfile' in line:
                msgfile = line.split('=')[1].lstrip().rstrip()
            elif 'funcfile' in line:
                funcfile = line.split('=')[1].lstrip().rstrip()
            elif 'keyfile' in line:
                keyfile = line.split('=')[1].lstrip().rstrip()

    return msgfile, funcfile, keyfile

if __name__ == '__main__':
    msgfile, funcfile, keyfile = parse_config('config')
    result_list = []

    parser = Parser(msgfile, funcfile, keyfile)
    msg_list = parser.parse_msg()
    msg_func_dict = parser.parse_func()
    func_key_dict = parser.parse_key()

    for msg in msg_list:
        if ':' in msg:
            msg_tmp = msg.split(':')[1].lstrip().rstrip()
            if msg_func_dict.has_key(msg_tmp):
                func = msg_func_dict[msg_tmp]
                if func_key_dict.has_key(func):
                    key = func_key_dict[func]
                    result_list.append([msg, func, key])
                else:
                    result_list.append([msg, func, ''])
            else:
                result_list.append([msg, '', ''])

    head = ['msg', 'func', 'key']
    saver = Saver('msg_to_key.xls', head, result_list)
    saver.save_xls()