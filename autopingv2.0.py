import pandas as pd
import time
import operator
import subprocess


def get_ping_list(filename, sheet,tables): #从excel表里读取objec列的所有数值
    table_device = pd.read_excel(filename,sheet)
    ping_list = table_device[tables]
    return  ping_list

def get_change_ping_list(ping_list):#序列化exce表的ip,让程序能能够处理
    change_ping_list =[]
    for ip in ping_list:
        change_ping_list.append(ip)
    fping_ping_list = " ".join(change_ping_list) #数组增加空格符合fping格式
    return fping_ping_list

def masscan_run(fping_ping_list): #masscan扫描并返回发现的ip列表
    m = subprocess.getoutput('masscan -p 22 %s --rate=100000'%fping_ping_list) 
    m_list_ip= []
    m_list = m.split('\n')
    for m_arr in m_list:
        if 'Discovered' in m_arr:
            m_arr_list = m_arr.split(" ")
            m_list_ip.append(m_arr_list[5])#将IP放入
    return m_list_ip

def fping_run(fping_ping_list,m_list_ip):#fping检测，并返回存活值
    p = subprocess.getoutput('fping -a -c 3 %s'%fping_ping_list) 
    device_live_list=[]
    fping_list = p.split('\n') #进行拆分结果
    for fping_arr in  fping_list:
        fping_ip_init = fping_arr.split(":")
        fping_ip = fping_ip_init[0].strip()#ip两边空格去除
        device_live_init = False
        for m_ip in m_list_ip:
            if fping_ip == m_ip:
                device_live_init = True
        device_live = operator.contains(fping_arr,"min") or device_live_init #fping检测或者masscan扫描有其中一项即可
        device_live_list.append(change_result(device_live)) #将True转换为1
    return device_live_list


def write_to_excel(filename, write_filename,sheet, ncols, write_result):#写入excel并保存
    wb = pd.read_excel(filename,sheet)
    wb.insert(loc=ncols, column=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),value=write_result, allow_duplicates=True)#新增一列并插入相应的数值
    wb.to_excel(write_filename)

def change_result(result): #将结果转换为1或者0,看你喜欢
    if result:
        return 1 
    else:
        return 0 

excel_name = 'test.xlsx' #源表
excel_sheet = 'Sheet1'
excel_table = 'IP' #表的索引
execl_ncols = 15 #插入表的第几列
execl_writename ='test-v1.0.xlsx'#保存表的名字

ping_list = get_ping_list(excel_name, excel_sheet,excel_table) #得到excel的某一列的数值
fping_ping_list = get_change_ping_list(ping_list)   #序列化能让fping执行的数组
m_list_ip = masscan_run(fping_ping_list)            #得到mascan扫描的结果IP组
device_live_list=fping_run(fping_ping_list,m_list_ip)   #通过fping结合masscan的结果得到存活率

write_to_excel(excel_name,execl_writename, excel_sheet, execl_ncols, device_live_list) #写入excel
