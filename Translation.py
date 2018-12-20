#!/usr/bin

### excel表格翻譯

import os
import hashlib
import random
import json
import urllib
import time,threading
from urllib import request,parse
from openpyxl import *

"""-----------------------------------------------------------------"""
### 百度翻译接口，如果翻译不成功，返回字符串 'THRANS_ERROR'
### 翻译成功，返回翻译字段
def BaiDuTranslate(appid, secretKey, fromLang, toLang, src_trans):
    print('Python test baidu fanyi api')
    
    s_url = 'http://api.fanyi.baidu.com/api/trans/vip/translate'  #提交地址
    salt = random.randint(32768, 65536)  #随机数
    sign = appid+src_trans+str(salt)+secretKey
    
    #计算MD5校验码
    o_md5 = hashlib.md5()
    o_md5.update(sign.encode('utf-8'))
    sign = o_md5.hexdigest()
    #end
    
    #格式化http有效数据
    #q = parse.quote_plus(q)
    data = {
        'appid': appid,
        'q': src_trans,
        'from': fromLang,
        'to': toLang,
        'salt': str(salt),
        'sign': sign
    }
    #end
    
    #编码数据并且POST提交，成功返回json数据，并进行解析
    data = parse.urlencode(data)
    data = data.encode('utf8')
    req = request.Request(s_url, data=data, method='POST')
    res = request.urlopen(req)

    #print(res.read().decode('utf8'))

    if res.status==200:
        json_s = res.read().decode('utf8')
        json_d = json.loads(json_s)
        if 'trans_result' in json_d.keys():
            result = json_d['trans_result']
            dst = result[0]
            dst = dst['dst']
            return str(dst)
        else:
            return 'THRANS_ERROR'
### 百度翻译接口函数定义结束
"""-----------------------------------------------------------------"""
            
"""-------------------------------------------------------------------"""
### 输入待处理的表格文件名称，不带格式
### 代码功能：
###     1、实现表格复制，并新保存表格
###     2、表格不存在，会输出提醒用户信息
def create_xlsx(src_filename):
    temp_filename = src_filename +'.xlsx'
    if not(os.path.isfile(temp_filename) and os.path.exists(temp_filename)):
        print('source file no exists...')
        return
    dest_filename = src_filename+'_temp.xlsx'
    src_wb = load_workbook(filename=temp_filename)
    dest_wb = Workbook()

    #复制整个表格至新表中
    for sheet in src_wb:
        print(sheet.title)
        dest_wb.create_sheet(sheet.title)
        src_ws = src_wb[sheet.title]
        dest_ws = dest_wb[sheet.title]
        row_len = len(list(src_ws.rows))
        col_len = len(list(src_ws.columns))
        
        for row in range(1,row_len+1):
            for col in range(1,col_len+1):
                cell = src_ws.cell(row=row, column=col) #获取表格中单元格
                dest_ws.cell(row=row, column=col, value=cell.value)
                col += 1
            row += 1
    dest_wb.save(dest_filename)
"""-------------------------------------------------------------------"""

"""-------------------------------------------------------------------"""
### 读取配置文件，配置文件不存在，则新建一个默认配置文件
def read_config_json():
    return_dic = {}
    json_data = {
        'appid':'modify appid',
        'key':'modify key',
        'fromLang':'en',
        'toLang':'zh',
        'select':'all',          #选择是否全部翻译，all-全部 select-只翻译没有翻译的内容
        'thread': 'on' ,         #是否开启线程翻译，on-开，off-关
        'thread_count': '500'    #线程一次处理的数据个数
    }

    json_filename = 'config'  #json配置文件名，不带后缀文件名格式
    json_filename += '.json'

    #判断文件是否存在，如果不存在，新建文件并且加入默认json数据
    if not(os.path.isfile(json_filename) and os.path.exists(json_filename)):
        print('file is no exists '+ json_filename)
        file = open(json_filename,'w')
        json.dump(json_data, file)  #将数据编码成json数据，并写入文件中
        file.close()
        print('file '+ json_filename +' created!')
        
    file = open(json_filename,'r')
    json_data = file.read() #读取文件中的数据
    json_dic = json.loads(json_data)  #将json数据解析成字典数据类型
        
    #检查数据是否有效
    if 7==len(json_dic.keys()) and ('appid' in json_dic.keys()):
        return_dic = json_dic
        return return_dic   #返回json数据，以字典数据类型返回
    else:
        print('数据配置错误，请检查文件')
    return {}
"""-------------------------------------------------------------------"""

"""-------------------------------------------------------------------"""
### 线程函数执行接口
def thread_loop(api_class, count):
    time.sleep(100)
    print(api_class)

"""-------------------------------------------------------------------"""


### 主程序代码
def user_main():
    config_data_dic = read_config_json() #读取配置数据
    print(config_data_dic)
    
    while True:
        filename = input("请输入表格的名称:  ")
        if not(os.path.isfile(filename+'.xlsx') and os.path.exists(filename+'.xlsx')):
            continue
        print("\t"+filename+'.xlsx'+"存在")
        break
    while True:    
        enCell = input("待翻译列：(如：A列)  ")
        if not enCell.isalpha():
            continue
        print("\t"+"输入>>" + enCell)
        break
    while True:
        zhCell = input("翻译保存列:  ")
        if not zhCell.isalpha():
            continue
        if zhCell==enCell:
            continue
        print("\t"+"输入>>" + zhCell)
        break
    print("输入数据完成")

    #多线程处理
    

    


user_main()
"""
### 测试代码
create_xlsx('tempxxx')
dd = BaiDuTranslate('20181219000250232', 'H_wPtwXs6KEDPtLra2ol' , 'en', 'zh', 'good good study')
print(dd)
### 测试结束
"""


