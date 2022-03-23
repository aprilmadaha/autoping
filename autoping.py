import re
import pandas as pd
import os
import time

def get_ping_list(filename, sheet): #从表里读取objec列的所有数值
    table_device = pd.read_excel(filename,sheet)
    ping_list = table_device['V2X-IP']
    return  ping_list

def write_to_excel(filename, sheet, ncols, write_result):#写入excel并保存
    wb = pd.read_excel(filename,sheet)
    wb.insert(loc=ncols, column=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),value=write_result, allow_duplicates=True)#新增一列并插入相应的数值
    wb.to_excel("testnew.xlsx")

def get_result(result): #得到结果
    if result:
        return 0 #不存活返回0
    else:
        return 1 #存活返回1

def get_tcpresult(result): #得到结果,tcp的结果转换下（和os相反）
    if result:
        return 0 #
    else:
        return 1 #

excel_name = 'text.xls'
excel_sheet = 'Sheet1'
ping_list = get_ping_list(excel_name, excel_sheet)
result_list = []
for ip in ping_list:
    result = os.system('fping -c 1 -t 260 %s' % ip)
    if result != 0: #如果ping不通则进行tcp检测
        resultcmd = os.popen('masscan -p 22 %s --rate=100000 --wait 1'%ip)
        resulttcp = resultcmd.read()
        result=get_tcpresult(len(resulttcp))
    result_list.append(get_result(result))
write_to_excel(excel_name, excel_sheet, 15, result_list)
