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

# 毫无用处的全局变量声明，只是为了标识哪些变量在局部作用域使用了
global rowParams
global mobiles
global http
global logFile
global result

rowParams = []
mobiles = []
http = Http()
result = []

# 切换目录到脚本运行目录，并读入配置
os.chdir(os.path.split(os.path.realpath(__file__))[0])
f = open('config.json', encoding='utf-8')
config = json.load(f)
f.close()
logFile = open(config['log-file'], mode='a', encoding='utf-8')


# 函数无论如何均返回 True
def log(content = None, msgType = 'Message', pure = None, level = 1) :
    global logFile
    # 仅当事件 level 小于或等于 config['log']['level'] 的记录时才进行记录
    if level > config['log']['level'] :
        return True

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
    global rowParams
    global mobiles
    if f is None :
        f = input('请输入要处理的 Excel 文档 > ')
    log('Reading file: ' + f + '...', level = 1)
    # 读取表格信息
    book = xlrd.open_workbook(f)
    table = book.sheet_by_index(0)
    log('File opened.', level = 2)
    for i in range(table.nrows) :
        # 第一列留作电话号码列，从第二列开始作为变量列读入
        mobiles.append(table.row_values(i)[0])
        rowParams.append(table.row_values(i)[1 : 1 + config['var-count']])
    # 读取完毕，所有信息读入到 rowParams 中
    log('Table analysis complete.', level = 2)
    book.release_resources()


# 显示已经读取的记录
def check() :
    print('共 %s 条记录，使用的模板：' % len(mobiles))
    print(config['template']['content'])
    print('------------------------------')
    for i in range(len(mobiles)) :
        print(str(mobiles[i]) + '\t' + '\t'.join(rowParams[i])) 


# 获取调用接口所需要的请求头
def getHeader() :
    log('Getting HTTP Headers...', level = 5)
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
# mobile -> string 电话号码
# params -> list 参数
def sendSingle(mobile, params) :
    log('Sending message to: %s with params %s' % (mobile, ','.join(params)), level = 3)
    header = getHeader()
    body = urlencode({
        'templateid': config['template']['id'],
        'mobiles': json.dumps([mobile]),
        'params': json.dumps(params)
    })
    log(pure = str(header) + '\n' + body + '\n--------', level = 4)
    log(pure = body, level = 3)
    # https://api.netease.im/sms/sendtemplate.action
    resp, body = http.request('https://api.netease.im/sms/sendtemplate.action',
        method='POST',
        headers = header,
        body = body
    )
    result.append([{"mobile": mobile, "params": params}, json.loads(body)])
    log('Request complete, response:', level = 3)
    log(pure = str(resp) + '\n' + str(body) + '\n--------', level = 4)
    log(pure = body.decode(), level = 3)
    print('%s - %s \t-> %s' % (mobile, ','.join(params), body.decode()))


def send() :
    log('Start sending message for each one...', level = 1)
    print('Sending Message for each one...')
    for i in range(len(mobiles)) :
        sendSingle(mobiles[i], rowParams[i])


# 清空读入的信息
def clear() :
    log('Clearing readed data..', level = 2)
    rowParams.clear()
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
            'q': lambda : log('Program exit.', level = 1) and exit(),
        }[choice]()
    except KeyError:
        print('\n输入了错误的选择！\n')


log('Program bootstrap complete... Running...', level = 1)
try:
    read(config['excel-file'])
    while True:
        action()
except BaseException:
    type_, value_, traceback_ = sys.exc_info()
    if (type_ is SystemExit) :
        exit()
    err = ''.join(traceback.format_exception(type_, value_, traceback_))
    log(pure = time.asctime() + '\tError:\tAn Exception has been thrown when processing...\n' + err + '\n---------', level = 1)
    print(err)

