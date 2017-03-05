#! python3
import json
import os
import keyboard
import xlrd
import random
from time import time
from hashlib import sha1
from urllib.parse import urlencode
from httplib2 import Http

global queue
global mobiles
global http

queue = []
mobiles = []
http = Http()

os.chdir(os.path.split(os.path.realpath(__file__))[0])
f = open('config.json', encoding="utf-8")
config = json.load(f)
f.close()


# 输出说明信息
def printInfo() :
    print(
    'SMS System for BBT - By Tamce in 2017\n'
    'Copyright by Tamce, All rights reserved.\n'
    '----------------------------------------\n'
    '注意：程序并不会联网查询实际的模板内容而是直接使用 config.json 中的配置，如果修改模板请同步修改 config.json\n'
    )
    pass


# 读 Excel 文件
def read() :
    f = input('请输入要处理的 Excel 文档 > ')
    # 读取表格信息
    book = xlrd.open_workbook(f)
    table = book.sheet_by_index(0)
    for i in range(table.nrows) :
        # 第一列留作电话号码列，从第二列开始作为变量列读入
        mobiles.append(table.row_values(i)[0])
        queue.append(table.row_values(i)[1 : 1 + config['var-count']])
    # 读取完毕，所有信息读入到 queue 中
    book.release_resources()
    pass


# 显示已经读取的记录
def check() :
    print('共 %s 条记录，使用的模板：' % len(queue))
    print(config['template']['content'])
    print('------------------------------')
    for i in range(len(queue)) :
        print(str(mobiles[i]) + '\t' + '\t'.join(queue[i]))
    pass


# 获取调用接口所需要的请求头
def getHeader() :
    random.seed()
    curTime = str(int(time()))
    nonce = ''.join(random.choices('0123456789', k = 32))
    checkSum = sha1((config['appsecret'] + nonce + curTime).encode()).hexdigest()
    return {
        'Content-Type': 'application/x-www-form-urlencoded;charset=utf-8',
        'AppKey': config['appkey'],
        'Nonce': nonce,
        'CurTime': curTime,
        'CheckSum': checkSum
    }

# 发送模板信息的接口调用
def send() :
    # https://api.netease.im/sms/sendtemplate.action
    resp, body = http.request('http://127.0.0.1/',
        headers = getHeader(),
        body = urlencode({
            'templateid': config['template']['id'],
            'mobiles': json.dumps(mobiles),
            "params": json.dumps(queue)
        })
    )
    print('\n------DEBUG------')
    print(resp)
    print('----------------')
    print(body)
    print('----------------')
    pass


# 清空读入的信息
def clear() :
    queue.clear()
    mobiles.clear()
    pass


# 功能选择
def action() :
    print(
    '-----------\n'
    'a > 读入文件\n'
    'b > 查看列表\n'
    'c > 发出短信\n'
    'd > 清空列表\n'
    'q > 退出程序\n'
    '-----------'
    )
    choice = input('请选择 > ')
    try:
        {
            'a': read,
            'b': check,
            'c': send,
            'd': clear,
            'q': exit,
        }[choice]()
    except KeyError:
        print('\n输入了错误的选择！\n')
    pass


while True:
    action()
