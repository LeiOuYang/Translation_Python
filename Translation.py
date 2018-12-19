#!/usr/bin

### excel表格翻譯

import os
import hashlib
import random
import json
import urllib
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

def user_main():
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

    


user_main()
"""
### 测试代码
create_xlsx('tempxxx')
dd = BaiDuTranslate('20181219000250232', 'H_wPtwXs6KEDPtLra2ol' , 'en', 'zh', 'good good study')
print(dd)
### 测试结束
"""


