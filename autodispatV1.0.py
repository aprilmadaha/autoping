from datetime import datetime
import openpyxl
import pandas as pd
import time

def get_work_value(excel_name):  #得到索引和工作界面值
    df =pd.read_excel(excel_name)#读取excel
    listindex = [] #索引数组
    listpole = [] #杆号
    listallresult = []
    for index,row in df.iterrows():                   #循环
        listindex.append(row[0])                      #获取index
        listpole.append(row[8])                      #获取杆子号
        listindex_sort =list(set(listindex))        #进行排序从小到大并去重
        listpole_sort = list(set(listpole))         #进行排序从小到大并去重
    for index_loop in listindex_sort:   #序号遍历
        for pole_loop in listpole_sort: #杆子号遍历
            pole_result = df[(df.序号==index_loop)& (df.百度杆位==pole_loop)] #先遍历路口,然后遍历每个路口的杆子
            if not pole_result.empty: #过滤掉路口没有这个杆号
                pole_result_in = pole_result.time.isin([1]).any() #杆号存活率里面筛选出包含1的数值，进行判断
                pole_result_index = pole_result.index #返回某个路口，某一杆子的索引

                if pole_result_in: #如果包含1
                    for  b_index in pole_result_index:
                        listallresult.append(str(b_index) + ' '+'百度')
                else:   #如果全部为0，那么为0
                    for w_index in pole_result_index:
                        listallresult.append(str(w_index)+' '+ '网联')
    return listallresult

def write_excel(excel_name,excel_sheet,list_workvalue,execl_ncols):
    wb= openpyxl.load_workbook(excel_name)
    ws = wb[excel_sheet]
    for workvalue in list_workvalue:
        ws.cell(row=int(workvalue.split(' ')[0])+2,column=execl_ncols).value=workvalue.split(' ')[1]
    saveexcel = excel_name.strip('.xlsx')+str(round(time.time()))+'.xlsx'#加入时间戳
    wb.save(saveexcel)

excel_name = 'test.xlsx' #源表
excel_sheet = 'Sheet1'
#excel_table = 'time'
execl_ncols = 15 #插入表的第几列

list_allresult = get_work_value(excel_name)
write_excel(excel_name,excel_sheet,list_allresult,execl_ncols)
