#! python3
import json
import sys
import os
import traceback
import keyboard
import xlrd
import random
import time
from hashlib import sha1
from urllib.parse import urlencode
from httplib2 import Http

# 毫无用处的全局变量声明
global queue
global mobiles
global http
global logFile

queue = []
mobiles = []
http = Http()

# 切换目录到脚本运行目录，并读入配置
os.chdir(os.path.split(os.path.realpath(__file__))[0])
f = open('config.json', encoding='utf-8')
config = json.load(f)
f.close()
logFile = open(config['log-file'], mode='a', encoding='utf-8')


def log(content = None, msgType = 'Message', pure = None) :
    global logFile
    if pure is None :
        logFile.write(time.asctime() + '\t' + msgType + ':\t' + content)
    else :
        logFile.write(pure)
    logFile.write('\n')
    logFile.flush()
    return True

# 输出说明信息
def printInfo() :
    print(
    'SMS System for BBT - By Tamce in 2017\n'
    'Copyright by Tamce, All rights reserved.\n'
    '----------------------------------------\n'
    '注意：程序并不会联网查询实际的模板内容而是直接使用 config.json 中的配置，如果修改模板请同步修改 config.json\n'
    )
    


# 读 Excel 文件
def read(f = None) :
    global queue
    global mobiles
    if f is None :
        f = input('请输入要处理的 Excel 文档 > ')
    log('Reading file: ' + f + '...')
    # 读取表格信息
    book = xlrd.open_workbook(f)
    table = book.sheet_by_index(0)
    log('File opened.')
    for i in range(table.nrows) :
        # 第一列留作电话号码列，从第二列开始作为变量列读入
        mobiles.append(table.row_values(i)[0])
        queue += table.row_values(i)[1 : 1 + config['var-count']]
    # 读取完毕，所有信息读入到 queue 中
    log('Table analysis complete.')
    book.release_resources()


# 显示已经读取的记录
def check() :
    print('共 %s 条记录，使用的模板：' % len(mobiles))
    print(config['template']['content'])
    print('------------------------------')
    for i in range(len(mobiles)) :
        print(str(mobiles[i]) + '\t' + '\t'.join(queue[i * config['var-count']:(i + 1) * config['var-count']])) 


# 获取调用接口所需要的请求头
def getHeader() :
    log('Getting HTTP Headers...')
    random.seed()
    curTime = str(int(time.time()))
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
    log('Ready to send request to https://api.netease.im/sms/sendtemplate.action...')
    header = getHeader()
    body = urlencode({
        'templateid': config['template']['id'],
        'mobiles': json.dumps(mobiles),
        'params': json.dumps(queue)
    })
    log(pure = '\n' + str(header) + '\n' + body + '\n--------')
    # https://api.netease.im/sms/sendtemplate.action
    resp, body = http.request('https://api.netease.im/sms/sendtemplate.action',
        method='POST',
        headers = header,
        body = body
    )
    log('Request complete, response:')
    print(resp)
    print(body)
    log(pure = '\n' + str(resp) + '\n' + str(body) + '\n--------')


# 清空读入的信息
def clear() :
    log('Clearing readed data..')
    queue.clear()
    mobiles.clear()


# 功能选择
def action() :
    print(
    '-----------\n'
#    'a > 读入文件\n'
    'b > 查看列表\n'
    'c > 发出短信\n'
    'd > 清空列表\n'
    'q > 退出程序\n'
    '-----------'
    )
    choice = input('请选择 > ')
    try:
        {
#            'a': read,
            'b': check,
            'c': send,
            'd': clear,
            'q': lambda : log('Program exit.') and exit(),
        }[choice]()
    except KeyError:
        print('\n输入了错误的选择！\n')


log('Program bootstrap complete... Running...')
try:
    read(config['excel-file'])
    while True:
        action()
except BaseException:
    type_, value_, traceback_ = sys.exc_info()
    if (type_ is SystemExit) :
        exit()
    err = ''.join(traceback.format_exception(type_, value_, traceback_))
    log(pure = time.asctime() + '\tError:\tAn Exception has been thrown when processing...\n' + err + '\n---------')
    print(err)

