#!/usr/bin

### excel表格翻譯

import os
import json
import urllib
import time,threading
from urllib import request,parse
from openpyxl import *

"""-----------------------------------------------------------------"""
### 功能说明
def information():
    print(
    """
            <<< 表格自动翻译工具 >>>

    使用说明:

        1、将需要翻译的表格文件防止脚本同级目录下，文件格式为 xlsx
        2、windows用户可运行目录下的run.bat文件
        3、输入文件名称，不需要输入文件格式
        4、输入需要翻译的表格单元格列名，例如： A，B
        5、输入翻译保存表格单元格
        6、直到运行完成即可(提示' 翻译完成 '字段信息)

    """
        )
"""-----------------------------------------------------------------"""

"""-----------------------------------------------------------------"""
### 百度翻译接口，如果翻译不成功，返回字符串 'THRANS_ERROR'
### 翻译成功，返回翻译字段
def BaiDuTranslate(appid, secretKey, fromLang, toLang, src_trans):
    #print('Python test baidu fanyi api')
    
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
    del dest_wb['Sheet']
    dest_wb.save(dest_filename)
    return (row_len, col_len)
"""-------------------------------------------------------------------"""

"""-------------------------------------------------------------------"""
### 线程函数执行接口
### src=待翻译列 dest=翻译保存列  start_count-开始行数  end_count-结束行数 
def baidu_thread_loop(src, dest, start_count, end_count):
    while True:
        time.sleep(0.01)
        

"""-------------------------------------------------------------------"""


### 主程序代码
def user_main():

    ### 各平台翻译接口网址定义
    #http://translate.google.cn/translate_a/single?client=gtx&dt=t&dj=1&ie=UTF-8&sl=auto&tl=zh_TW&q=calculate
    google_trans_url = 'http://translate.google.cn/translate_a/single?client=gtx'
    google_trans_data = 'good'
    google_trans_data_struct = { 'dt':'t',  'dj':'1', 'ie':'UTF-8', 'sl':'en', 'tl':'zh_CN', 'q':google_trans_data}

    #http://api.microsofttranslator.com/v2/Http.svc/Translate?appId=AFC76A66CF4F434ED080D245C30CF1E71C22959C&from=&to=en&text=考勤计算
    bing_trans_url = 'http://api.microsofttranslator.com/v2/Http.svc/Translate?'

    #http://fanyi.youdao.com/translate?&doctype=json&type=AUTO&i=计算
    youdao_trans_url = 'http://fanyi.youdao.com/translate?'

    #http://fanyi.baidu.com/transapi?from=auto&to=cht&query=Calculation
    baidu_trans_url = 'http://fanyi.baidu.com/transapi?'

    headers = {
        #'Accept':'text/html,application/xhtml+xml,aplication/xml;q=0.9,*/*;q=0.8',
        #'Accept-Language':'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        #'Connection':'keep-alive',
        'Host':'translate.google.cn',
        'Upgrade-Insecure-Requests':'1',
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0;WOW64;rv:64.0) Gecko/20100101 Firefox/64.0'
    }
    ### 结束各平台翻译网址定义
    
    data = parse.urlencode(google_trans_data_struct)
    data = data.encode('utf8')
    req = request.Request(url=google_trans_url, data=data, headers=headers, method="GET")
    print(req.data)
    request.urlopen(req)
    
    #全局数据
    wb_max_rows = 0  #最大行 
    wb_max_col = 0   #最大列
    wb_src_filename = '' #待表格翻译文件名
    enCell = None
    zhCell = None

    information()    #显示提示信息

    

"""

    while True:
        filename = input("\n\n请输入表格的名称:  ")
        if not(os.path.isfile(filename+'.xlsx') and os.path.exists(filename+'.xlsx')):
            continue
        print("\t"+filename+'.xlsx'+"存在")
        wb_src_filename = filename+'.xlsx'
        t_len = create_xlsx(filename)  #创建表格文件，并复制该表格
        wb_max_rows = t_len[0]
        wb_max_col = t_len[1]
        if baidu_count<=wb_max_rows:
            baidu_count = wb_max_rows
        break
    while True:    
        fromCell = input("待翻译列：(如：A列)  ")
        if not fromCell.isalpha():
            continue
        fromCell = fromCell.strip().upper()
        print("\t"+"输入>>" + fromCell)
        break
    while True:
        toCell = input("翻译保存列:  ")
        toCell = toCell.strip().upper()
        if not toCell.isalpha():
            continue
        if toCell==fromCell:
            continue
        print("\t"+"输入>>" + toCell)
        break
    print("输入数据完成")

    wb_src_filename = filename+'_temp.xlsx'
    wb_src = load_workbook(wb_src_filename)  #加载工作表
    ws_sheet = wb_src[wb_src.sheetnames[0]]  #取工作表中第一张表
    print('rows:'+str(wb_max_rows) + '  col: '+str(wb_max_col) + ' sheet: '+ws_sheet.title)

    row = baidu_start_pos   #开始翻译行开始数
    for row in range(baidu_start_pos,baidu_count+1):
        error_count = 0
        result = ''
        fromCell_t = str(fromCell)+str(row)
        toCell_t = str(toCell)+str(row)
        cell = ws_sheet[fromCell_t]
        srcs = cell.value
        cell = ws_sheet[toCell_t]
        dests = cell.value
        if (srcs is None) or (''==srcs.strip()) or ((dests!=None) and (''!=dests.strip())): 
            row += 1
            #print('none')
            continue
        while True:
            result = BaiDuTranslate(baidu_appid, baidu_key, baidu_fronLang, baidu_toLang, srcs)
            print(result)
            if 'THRANS_ERROR'==result:
                error_count += 1
                continue
                if error_count>10:
                    print(fromCell_t + ' translate error...')
                    break
            else:
                break
        cell.value = ''
        cell.value += result
        #print('fanyi ok')
        row += 1
    wb_src.save(wb_src_filename)
    print("<<< 翻译工作完成 >>>")
    """
        
        
    #多线程处理
    
    #baidu_thread = threading.Thread(target=baidu_thread_loop, name="BAIDU_API", args=('baidu',12 ))
    #baidu_thread.start()
    


user_main()
"""
### 测试代码
create_xlsx('tempxxx')
dd = BaiDuTranslate('20181219000250232', 'H_wPtwXs6KEDPtLra2ol' , 'en', 'zh', 'good good study')
print(dd)
### 测试结束
"""


