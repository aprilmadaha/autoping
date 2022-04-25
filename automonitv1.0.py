import pandas as pd
import time
import operator
import subprocess
import openpyxl
import os
  
def get_monit_value(excel_name):  #筛选工作界面，得到索引和工作界面值
    df =pd.read_excel(excel_name)#读取excel
    listindex = [] #索引数组
    listindex_sort = []
    for index,row in df.iterrows():                   #循环
        listindex.append(row[0])  
        listindex_sort =list(set(listindex))
    for indexlist in listindex_sort:
        indexall = df[(df.序号== indexlist)]
        str11 = '+'+' '+str(indexlist)+'\n'
        str12 = 'menu = '+str(indexlist)+'road'+'\n'
        str13 = 'title = '+str(indexlist)+'road'+'\n'
        str14 = str11+str12+str13+'\n'
        # print("indexall",indexall)
        fm = open('liy.txt',mode='a+')
        fm.write(str14)

        for indexl,rowl in indexall.iterrows():
            indexNum = str(rowl[0])+'-'+str(indexl)
            ipNum = row[4]
            str1 = '++'+' '+indexNum+'\n'
            str2 = 'menu = '+indexNum+'\n'
            str3 = 'title = '+indexNum+' '+ipNum+'\n'
            str4 = 'host'+'='+ipNum+'\n'
            str5 = str1+str2+str3+str4+'\n'
            # print(str5)
            fm = open('liy.txt',mode='a+')
            fm.write(str5)
 
excel_name = 'yzjk.xlsx' #源表
get_monit_value(excel_name)
