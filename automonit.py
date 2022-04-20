import pandas as pd
import time
import operator
import subprocess
import openpyxl
import os
  
def get_monit_value(excel_name):  #筛选工作界面，得到索引和工作界面值
    df =pd.read_excel(excel_name)#读取excel
    for index,row in df.iterrows():                   #循环
        indexNum = str(row[0])+'-'+str(index)
        ipNum = row[2]
        str1 = '++'+' '+indexNum+'\n'
        str2 = 'menu = '+indexNum+'\n'
        str3 = 'title = '+indexNum+' '+ipNum+'\n'
        str4 = 'host'+'='+ipNum+'\n'
        str5 = str1+str2+str3+str4+'\n'
        # listall.append(listall)
        fm = open('a.txt',mode='a+')
        fm.write(str5)
 
excel_name = 'monit.xlsx' #源表
get_monit_value(excel_name)
