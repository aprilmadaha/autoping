import pandas as pd
import time
import operator
import subprocess
import openpyxl
import os
import numpy as np

cameraCount = 0
cameraLossCount =0 
cameraLossCheck =0
cameraLossNoCheck= 0
cameraListLossAll = []

ccuCount = 0
ccuLossCount =0 
ccuLossCheck =0
ccuLossNoCheck= 0
ccuListLossAll = []

rsuCount = 0
rsuLossCount =0 
rsuLossCheck =0
rsuLossNoCheck= 0
rsuListLossAll = []

radarCount = 0
radarLossCount =0 
radarLossCheck =0
radarLossNoCheck= 0
radarListLossAll = []

idcCount = 0
idcLossCount =0 
idcLossCheck =0
idcLossNoCheck= 0
idcListLossAll = []

loop1 = [118,286,119,120,121,138,137,144,151,150,288,153,152,149,143,136,135,127,129,128] #网信大环号和百度路口号的对应关系
loop2 = [125,126,134,133,141,142,148,147,146,145,287,132,131,140,139,130,122,123,124]
loop3 = [114,107,99,100,108,109,101,102,110,117,116,115]
loop4 = [95,262,263,264,282,283,284,98,97,104,105,106,113,285,112,111,103,96]
loop5 = [91,92,261,260,94,93,259,258,88,83,84,89,90,85]
loop6 = [86,81,80,79,78,74,75,10008,76,77,82,87]
loop7 = [20,15,10,5,6,2,3,8,7,11,12,13,18,23,22,17,16,21]
loop8 = [55,44,45,41,46,56,255,256,64,63,60,59,254]
loop9 = [62,66,67,68,72,71,70,69,65,304,61,58,253]
loop10 = [38,43,252,257,57,54,53,52,35,31,51,50,49,30,48,47,27,24,37]
loop11 = [40,42,39,36,34,32,28,25,26,29,33,9,4,1,14,2]
loop12 = [162,163,174,173,182,183,184,193,192,191,181,172]
loop13 = [200,201,202,203,204,213,294,212,223,222,211]
loop14 = [207,214,218,303,302,301,206,205,194,195,186,185,175,164,165,176,187,196]
loop15 = [189,190,199,198,209,210,217,216,220,221,290,219,215,208,197,188]
loop16 = [224,228,229,230,237,236,235,296,239,295,240,247,246,245,244,238,234]
loop17 = [227,243,250,251,248,249,299,242,241,297,298,233,232,231,225,226]
loop18 = [289,156,154,155,292,293,157,161,171,160,159,158,291]
loop19 = [169,170,180,179,178,177,166,300,167,168]
loop20 = [73,276,265,266,267,268,269,275,281,280,274,273,279,278,272,271,270,277]

loopAll = np.array([loop1,loop2,loop3,loop4,loop5,loop6,loop7,loop8,loop9,loop10,loop11,loop12,loop13,loop14,loop15,loop16,loop17,loop18,loop19,loop20])
roadList = [2,9,55,60,62,64,74,75,76,80,81,83,93,97,98,100,101,102,104,107,119,122,129,135,136,141,142,148,157,182,184,188,189,192,205,209,214,215,216,217,238,240,242,244,249,267,268,282,286,295,299]
xinSiRoadlist=[162,172,181,200,211,221,222,224,225,226,227,231,232,238,239,240,241,243,244,245,246,247,248,249,250,251,292,293,295,296,297,298,59,60,62,63,64,66,67,68,71,72]

dateTime = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime())
# dateTimeS = dateTime.split(' ')[0]
dateM = dateTime.split(' ')[0].split('-')[1]
dateD = dateTime.split(' ')[0].split('-')[2]

def add_sheet(excel_name):
    wb = openpyxl.load_workbook(excel_name)
    wb.create_sheet(title= 'Sheet1',index=1)
    wb.save(excel_name)

def get_loss_value(excel_name,checkValue):  #筛选工作界面，得到索引和工作界面值
    # global countNum,lossCountNum,lossCheckNum,lossNoCheckNum
    df =pd.read_excel(excel_name)#读取excel
    listAll = []
    listLossAll = []
    listOfflinePowerdown = []
    n = 1
    lossCount = 0
    lossCheck = 0
    offLineCount = 0 #离线计数
    powerDownCount = 0 #掉电计数
    norMalCount = 0

    clos =[i for i in df.columns if i not in ['Time']]
    df2 = df[clos]
    for column in df2.columns:
        n= n+1 
        listallresult = df2[column].tolist()
        a = np.array(df2[column].tolist())
        vu = a[(np.where((a>0) & (a<1000)))]
        # vu = a[~np.isnan(a)]
        # print("column",column)
        # print("a",a)
        # print("vu",vu)
        # print(np.where(a == 1000))
        # print(np.size(np.where(a == 1000)))
        num1000 = np.size(np.where(a == 1000)) 
        numNan= np.size(np.where(np.isnan(a))) 
        num0 = np.size(np.where(a == 0)) 
        numoffline = numNan + num1000
        # if numNan > 0:
        #     print("column",column)
        #     print("numNan",numNan)
        #     print("num1000",num1000)
        #     print("numoffline",numoffline)
        # if num1000==288:
        #     # print("离线",column)
        #     offLineCount = offLineCount +1
        #     listOfflinePowerdown.append("离线")
        # elif num1000 >0 and num0 >0:
        #     powerDownCount = powerDownCount +1
        #     listOfflinePowerdown.append("掉过电")
        # else:
        #     norMalCount = norMalCount+1
        #     listOfflinePowerdown.append("没掉电过")

        # b,s ,t,w= np.unique(a,return_counts=True,return_index=True,return_inverse=True)
        # print("column",column)
        # print("b",b)
        # print("b.size",b.size)
        # print("b.sum",b.sum())
        # print("s",s)
        # print("t",t)
        # print("w",w)
        # b2,s2 ,t2,w2= np.unique(vu,return_counts=True,return_index=True,return_inverse=True)
        # print("column",column)
        # print("b2",b2)
        # print("s2",s2)
        # print("t2",t2)
        # print("w2",w2)

        b1,s1 = np.unique(vu,return_counts=True)
        # b1,s1 = np.unique(a,return_counts=True)

        # print("column",column)
        # print("b1",b1)
        # print("b1.sum",b1.sum())
        # print("s1",s1)
   
        # print("s1.sum",s1.sum())
        # print("s1.size",s1.size)
        
        # print("column",column)
        #b,s ,t,w= np.unique(a,return_counts=True,return_index=True,return_inverse=True)
        sumNum = int(np.nansum(a)) #求和把nan去掉
        # print("n",n,"sumNum",sumNum)
        # print("vu",vu)
        # print("vu.sum",vu.sum())
        # print("\n")
        if sumNum > checkValue: #是否丢包超过阀值
            lossCount = lossCount+1
            #columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum)
            if s1.size >= 5:
                lossCheck = lossCheck+1
                if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                    offLineCount = offLineCount +1
                    columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                    listAll.append(columnV)
                    listLossAll.append(column)
                elif 30>num1000 >0 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                    listAll.append(columnV)
                    listLossAll.append(column)
                elif 260>num1000 >=30 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                    listAll.append(columnV)
                    listLossAll.append(column)
                elif 290>num1000 >=260 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                    listAll.append(columnV)
                    listLossAll.append(column)
                else:
                    norMalCount = norMalCount+1
                    columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                    listAll.append(columnV)
                    listLossAll.append(column)
            elif s1.size == 2 and s1.sum() >= 26 :
                    lossCheck = lossCheck+1
                    if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                        offLineCount = offLineCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 30>num1000 >0 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 260>num1000 >=30 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 290>num1000 >=260 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    else:
                        norMalCount = norMalCount+1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                        listAll.append(columnV)
                        listLossAll.append(column)
            elif s1.size == 3 and s1.sum() >= 14:
                    lossCheck = lossCheck+1
                    if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                        offLineCount = offLineCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 30>num1000 >0 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 260>num1000 >=30 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 290>num1000 >=260 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    else:
                        norMalCount = norMalCount+1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                        listAll.append(columnV)
                        listLossAll.append(column)
            elif s1.size == 4 and s1.sum() >= 4:
                    lossCheck = lossCheck+1
                    if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                        offLineCount = offLineCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 30>num1000 >0 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 260>num1000 >=30 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    elif 290>num1000 >=260 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                        listAll.append(columnV)
                        listLossAll.append(column)
                    else:
                        norMalCount = norMalCount+1
                        columnV = str(n) + ' ' + str(1) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                        listAll.append(columnV)
                        listLossAll.append(column)
            elif  s1.size ==1 and s1.sum()==1 and b1.sum()>checkValue:
                    if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                        offLineCount = offLineCount +1
                        columnV = str(n) + ' ' + str(3) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                        listAll.append(columnV)

                    elif 30>num1000 >0 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(3) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                        listAll.append(columnV)
                    
                    elif 260>num1000 >=30 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(3) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                        listAll.append(columnV)
                    
                    elif 290>num1000 >=260 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(3) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                        listAll.append(columnV)
                    
                    else:
                        norMalCount = norMalCount+1
                        columnV = str(n) + ' ' + str(3) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                        listAll.append(columnV)
            else:#丢包但不排查的
                    if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                        offLineCount = offLineCount +1
                        columnV = str(n) + ' ' + str(2) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                        listAll.append(columnV)
                       
                    elif 30>num1000 >0 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(2) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                        listAll.append(columnV)
                        
                    elif 260>num1000 >=30 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(2) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                        listAll.append(columnV)
                       
                    elif 290>num1000 >=260 and num0 >0:
                        powerDownCount = powerDownCount +1
                        columnV = str(n) + ' ' + str(2) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                        listAll.append(columnV)
                        # listLossAll.append(column)
                    else:
                        norMalCount = norMalCount+1
                        columnV = str(n) + ' ' + str(2) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                        listAll.append(columnV)
                        # listLossAll.append(column)
        else: #没问题的
                if numoffline==289 or num1000 == 289  or num0 <1  : #状态判断 离线的
                    offLineCount = offLineCount +1
                    columnV = str(n) + ' ' + str(0) + ' ' + column + ' ' +str(sumNum) + ' ' +'离线'
                    listAll.append(columnV)

                elif 30>num1000 >0 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(0) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率小于10%'
                    listAll.append(columnV)
                
                elif 260>num1000 >=30 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(0) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率10%-90%'
                    listAll.append(columnV)
                
                elif 290>num1000 >=260 and num0 >0:
                    powerDownCount = powerDownCount +1
                    columnV = str(n) + ' ' + str(0) + ' ' + column + ' ' +str(sumNum) + ' ' +'掉线率大于90%'
                    listAll.append(columnV)
                
                else:
                    norMalCount = norMalCount+1
                    columnV = str(n) + ' ' + str(0) + ' ' + column + ' ' +str(sumNum) + ' ' +'未掉过线'
                    listAll.append(columnV)
              
    # print("n",n)
    # print(listAll)
    countNum = n -1
    lossCountNum = lossCount
    lossCheckNum = lossCheck
    lossNoCheckNum=  lossCountNum - lossCheckNum
    # print("countNum",countNum)
    # print("lossCountNum",lossCountNum)
    # print("lossCheckNum",lossCheckNum)
    # print("lossNoCheckNum",lossNoCheckNum)
    # print("listLossAll",listLossAll)
    # print("len(listLossAll)",len(listLossAll))
    # print("len(listAll)",len(listAll))
    # print("listAll",listAll)
    return listAll,countNum,lossCountNum,lossCheckNum,lossNoCheckNum,listLossAll,offLineCount,powerDownCount,norMalCount

def rscu_now_status(rscuStatus):
    if rscuStatus == -1:
        rscu_status = '断电'
    elif rscuStatus == 0:
        rscu_status = '没问题'
    elif rscuStatus == 1:
        rscu_status = '4G不通'
    elif rscuStatus == 2:
        rscu_status = '到机房不通'
    elif rscuStatus == 3:
        rscu_status = '关机'
    else:
        rscu_status ='未知'
    return rscu_status

def rscu_history_status(rscuHistoryStatus):
    historyStatus = []
    for hiIndex in rscuHistoryStatus:
        if hiIndex == -1:
            historyStatus.append('断电')
        elif hiIndex == 0:
            historyStatus.append('没问题')
        elif hiIndex == 1:
            historyStatus.append('4G不通')
        elif hiIndex == 2:
            historyStatus.append('到机房不通')
        elif hiIndex == 3:
            historyStatus.append('关机')
        else:
            historyStatus.append('未知')
    return historyStatus
 

def get_rscu_status(excel_name):  #筛选工作界面，得到索引和工作界面值
 # global countNum,lossCountNum,lossCheckNum,lossNoCheckNum
    df =pd.read_excel(excel_name)#读取excel
    listAll = []
    listLossAll = []
    listOfflinePowerdown = []
    n = 1
    lossCount = 0
    lossCheck = 0
    offLineCount = 0 #离线计数
    powerDownCount = 0 #掉电计数
    norMalCount = 0

    clos =[i for i in df.columns if i not in ['Time']]
    df2 = df[clos]
    for column in df2.columns:
        # print("column",column)
        # n= n+1 
        listallresult = df2[column].tolist()
        nowStatus = listallresult[288]
        # print("listallresult",listallresult[288])
        # for ii in listallresult:
            # print("ii",ii)
        a = np.array(df2[column].tolist())

        b,w= np.unique(a,return_counts=True)
        # print("column",column)
        # print("nowStatus",nowStatus)
        # print("judge_rscu_now_status",rscu_now_status(int(nowStatus)))
        rscuNowStatus = rscu_now_status(int(nowStatus))

        # print("b",b)
        # print("s",s)
        # print("t",t)
        # print("w",w)
        # print("rscu_history_status",rscu_history_status(b))
        rscuHistoryStatus1 = rscu_history_status(b)
        rscuHistoryStatus = "|".join(rscuHistoryStatus1)
        # print("str(rscu_history_status)",yy)
        columnV= str(column)+' ' +str(rscuNowStatus)  +' ' +str(rscuHistoryStatus)
        listAll.append(columnV)
        # print('columnV',columnV)
    # print("listAll",listAll)
    return listAll
       
        

def impact_road(listLossAll):
    #roadList = [2,9,55,60,62,64,74,75,76,80,81,83,93,97,98,100,101,102,104,107,119,122,129,135,136,141,142,148,157,182,184,188,189,192,205,209,214,215,216,217,238,240,242,244,249,267,268,282,286,295,299]
    impactRoadNum=0
    xinSiRoadNum=0
    roadLossList = []
    impactRoadList=[]
    xinSiRoadList = []
    for i,valueR in enumerate(listLossAll):
        roadNum = int(valueR.split('-')[1])
        roadLossList.append(roadNum)
        
    roadLossListSort = list(set(roadLossList)) 

    for roadLossNum in roadLossListSort:
        if roadLossNum in roadList:
            impactRoadNum = impactRoadNum+1
            impactRoadList.append(roadLossNum)
    for roadLossNum1 in roadLossListSort:
        if roadLossNum1 in xinSiRoadlist:
            xinSiRoadNum = xinSiRoadNum+1
            xinSiRoadList.append(roadLossNum1)
    return impactRoadNum,xinSiRoadNum,impactRoadList,xinSiRoadList

def impact_road_idc(listLossAll):
    #roadList = [2,9,55,60,62,64,74,75,76,80,81,83,93,97,98,100,101,102,104,107,119,122,129,135,136,141,142,148,157,182,184,188,189,192,205,209,214,215,216,217,238,240,242,244,249,267,268,282,286,295,299]
    impactRoadNum=0
    xinSiRoadNum=0
    roadLossList = []
    impactRoadList=[]
    xinSiRoadList = []
    for i,valueR in enumerate(listLossAll):
        roadNum = int(valueR[4:7])
        #print("valueR",valueR)
    #    print("valueR[4:7]",valueR[4:7])
       # print("roadNum",roadNum)
        roadLossList.append(roadNum)
        
    roadLossListSort = list(set(roadLossList)) 
    # print("roadLossListSort",roadLossListSort)
    for roadLossNum in roadLossListSort:
        if roadLossNum in roadList:
            impactRoadNum = impactRoadNum+1
            impactRoadList.append(roadLossNum)
    # print ("impactRoadNum",impactRoadNum)
    for roadLossNum1 in roadLossListSort:
        if roadLossNum1 in xinSiRoadlist:
            xinSiRoadNum = xinSiRoadNum+1
            xinSiRoadList.append(roadLossNum1)
    return impactRoadNum,xinSiRoadNum,impactRoadList,xinSiRoadList

def write_excel(excel_name,excel_sheet,excel_save,list_value):
    wb= openpyxl.load_workbook(excel_name)
    ws = wb[excel_sheet]
    #ws.delete_cols(1)
    for lossValue  in list_value:
        ws.cell(row=291,column=int(lossValue.split(' ')[0])).value=int(lossValue.split(' ')[1])
  
    wb.save(excel_save)

def write_excel1(excel_name,excel_sheet,excel_save,list_value):
    wb= openpyxl.load_workbook(excel_name)
    ws1 = wb[excel_sheet]
    #ws.delete_cols(1)
    # for cameravalue  in list_value:
    #     ws.cell(row=291,column=int(cameravalue.split(' ')[0])).value=int(cameravalue.split(' ')[1])
  
    ws1.cell(row=1,column=1).value='路口号'
    ws1.cell(row=1,column=2).value='设备号'
    ws1.cell(row=1,column=3).value='IP'
    ws1.cell(row=1,column=4).value='丢包数'
    ws1.cell(row=1,column=5).value='是否排查'
    ws1.cell(row=1,column=6).value='状态'
    for i,valueR in enumerate(list_value):
        # print(valueR)
        roleNume = valueR.split('-')[1]
        deviceIP = valueR.split('-')[2].split('_')[1].split(' ')[0]
        deviceName = valueR.split(' ')[2]
        checkReult = valueR.split(' ')[1]
        pingLoss =valueR.split(' ')[3]
        deviceState = valueR.split(' ')[4]

        ws1.cell(row=i+2,column=1).value=int(roleNume)
        ws1.cell(row=i+2,column=2).value=deviceName
        ws1.cell(row=i+2,column=3).value=deviceIP
        ws1.cell(row=i+2,column=4).value=int(pingLoss)
        ws1.cell(row=i+2,column=5).value=int(checkReult)
        ws1.cell(row=i+2,column=6).value=deviceState
    # for i1,valueR1 in enumerate(list_Offline_Powerdown):
    #      ws1.cell(row=i1+2,column=6).value=valueR1

    wb.save(excel_save)

def write_excel_idc(excel_name,excel_sheet,excel_save,list_value):
    wb= openpyxl.load_workbook(excel_name)
    ws1 = wb[excel_sheet]
    #ws.delete_cols(1)
    # for cameravalue  in list_value:
    #     ws.cell(row=291,column=int(cameravalue.split(' ')[0])).value=int(cameravalue.split(' ')[1])
    # bigLoopList = []
    
    ws1.cell(row=1,column=1).value='路口号'
    ws1.cell(row=1,column=2).value='设备号'
    #ws1.cell(row=1,column=3).value='IP'
    ws1.cell(row=1,column=3).value='丢包数'
    ws1.cell(row=1,column=4).value='是否排查'
    ws1.cell(row=1,column=5).value='涉及的大环'
    ws1.cell(row=1,column=6).value='状态'
    for i,valueR in enumerate(list_value):
        # print(valueR)
        roleNume = valueR.split(' ')[2]
        # print("roleNume",roleNume)
        # print("valueR",valueR)
        deviceName = valueR.split(' ')[2]
        checkReult = valueR.split(' ')[1]
        pingLoss =valueR.split(' ')[3]
        deviceState =valueR.split(' ')[4]
        ws1.cell(row=i+2,column=1).value=int(roleNume[4:7])
        ws1.cell(row=i+2,column=2).value=deviceName
        # ws1.cell(row=i+2,column=3).value=deviceIP
        ws1.cell(row=i+2,column=3).value=int(pingLoss)
        ws1.cell(row=i+2,column=4).value=int(checkReult)
        ws1.cell(row=i+2,column=6).value=deviceState

        # for roadNumIndex in roadNumList:   # RSCU到机房路口分析大环问题,通过丢包路口号返回大环号
        bigLoopNum = 0 
        for loopAllIndex in loopAll:
            bigLoopNum = bigLoopNum+1
            if int(roleNume[4:7]) in loopAllIndex:
                # bigLoopList.append(n)
               # print("save-bigLoopList",n,"int(roleNume[4:7])",int(roleNume[4:7]))
                ws1.cell(row=i+2,column=5).value=int(bigLoopNum)
    wb.save(excel_save)

def write_excel_rscu(excel_name,excel_sheet,excel_save,list_value):
    wb= openpyxl.load_workbook(excel_name)
    ws1 = wb[excel_sheet]
    #ws.delete_cols(1)
    # for cameravalue  in list_value:
    #     ws.cell(row=291,column=int(cameravalue.split(' ')[0])).value=int(cameravalue.split(' ')[1])
    # bigLoopList = []
    
    ws1.cell(row=1,column=1).value='路口号'
    ws1.cell(row=1,column=2).value='设备号'
    #ws1.cell(row=1,column=3).value='IP'
    ws1.cell(row=1,column=3).value='当前状态'
    ws1.cell(row=1,column=4).value='历史发生过的状态'
    ws1.cell(row=1,column=5).value='涉及的大环'
    # ws1.cell(row=1,column=6).value='状态'
    for i,valueR in enumerate(list_value):
        # print(valueR)
        roleNume = valueR.split(' ')[0].split('_')[1]
        # print("roleNume",roleNume)
        # print("valueR",valueR)
        deviceName = valueR.split(' ')[0]
        nowStatus = valueR.split(' ')[1]
        historyStatus =valueR.split(' ')[2]
        # deviceState =valueR.split(' ')[4]
        ws1.cell(row=i+2,column=1).value=int(roleNume[4:7])
        ws1.cell(row=i+2,column=2).value=deviceName
        # ws1.cell(row=i+2,column=3).value=deviceIP
        ws1.cell(row=i+2,column=3).value=nowStatus
        ws1.cell(row=i+2,column=4).value=historyStatus
        # ws1.cell(row=i+2,column=6).value=deviceState

        # for roadNumIndex in roadNumList:   # RSCU到机房路口分析大环问题,通过丢包路口号返回大环号
        bigLoopNum = 0 
        for loopAllIndex in loopAll:
            bigLoopNum = bigLoopNum+1
            if int(roleNume[4:7]) in loopAllIndex:
                # bigLoopList.append(n)
               # print("save-bigLoopList",n,"int(roleNume[4:7])",int(roleNume[4:7]))
                ws1.cell(row=i+2,column=5).value=int(bigLoopNum)
    wb.save(excel_save)


def get_roadNum(deviceList): #得到各种类型的路口号
    roadNumList = []
    offlineNumList = []
    powerdownNumList = []
    for i,valueR in enumerate(deviceList):
        roadIndex = int(valueR.split(' ')[1])
        stateIndex = valueR.split(' ')[4]
       # print("roadIndex",roadIndex)
        if roadIndex==1:
            roadNume = int(valueR.split('-')[1])
            roadNumList.append(roadNume)
        if stateIndex == '离线':
            offlineNume = int(valueR.split('-')[1])
            offlineNumList.append(offlineNume)
        elif stateIndex == '未掉电过':
            powerdownNume = int(valueR.split('-')[1])
            powerdownNumList.append(powerdownNume)
    
    uniqueRoadNum,uniqueRoadNumCount= np.unique(roadNumList,return_counts=True)
    uniqueofflineNum,uniqueofflineNumCount= np.unique(offlineNumList,return_counts=True)
    uniquepowerdownNum,uniquepowerdownNumCount= np.unique(powerdownNumList,return_counts=True)
    # set_roadNumList =set(roadNumList)
    # print(len(set_roadNumList))
    # print("roadNumList",roadNumList)
    # print("sorted(roadNumList)",sorted(roadNumList))
    # print("set_roadNumList",set_roadNumList)
    # print("uniqueofflineNum",uniqueofflineNum)
    # print("uniqueofflineNumCount",uniqueofflineNumCount)
    # print("uniquepowerdownNum",uniquepowerdownNum)
    # print("uniquepowerdownNumCount",uniquepowerdownNumCount)
    # print("b",b)
    # print("s",s)
    # print("t",t)
    # print("w",w)
    # print("\n")
    return uniqueRoadNum,uniqueRoadNumCount,uniqueofflineNum,uniqueofflineNumCount,uniquepowerdownNum,uniquepowerdownNumCount

def get_idcroadNum(deviceList): #得到每个表的路口号
    roadNumList = []
    bigLoopList = []
    for i,valueR in enumerate(deviceList):
        roadIndex = int(valueR.split(' ')[1])
        # print("roadIndex",roadIndex)
        if roadIndex == 1:
            roadNume = valueR.split(' ')[2]
            # print("i",i)
            # print("roadNume",roadNume,int(roadNume[4:7]))
            roadNumList.append(int(roadNume[4:7]))
    # b,s ,t,w= np.unique(roadNumList,return_counts=True,return_index=True,return_inverse=True)
    uniqueRoadNum,uniqueRoadNumCount= np.unique(roadNumList,return_counts=True)

    for roadNumIndex in roadNumList:   # RSCU到机房路口分析大环问题,通过丢包路口号返回大环号
        n = 0
        for loopAllIndex in loopAll:
            n = n+1
            if roadNumIndex in loopAllIndex:
                bigLoopList.append(n)
    # print("bigLoopList",bigLoopList)
    uniqueBigLoop,uniqueBigLoopCount =np.unique(bigLoopList,return_counts=True)
    # print("uniqueBigLoop",uniqueBigLoop)
    # print("uniqueBigLoopCount",uniqueBigLoopCount)
    # dictValue = dict(zip(uniqueBigLoop,uniqueBigLoopCount))
    # print("dictIdcValue",dictValue)
    return uniqueRoadNum,bigLoopList,uniqueBigLoop,uniqueBigLoopCount

def compare_allRoad(camerList,ccuList,rsuList,radarList,idcList):
    cameraLossRoad,cameraLossRoadCount,cameraofflineRoad,cameraofflineRoadCount,camerapowerdownRoad,camerapowerdownRoadCount = get_roadNum(camerList)
    ccuLossRoad,ccuLossRoadCount,ccuofflineRoad,ccuofflineRoadCount,ccuapowerdownRoad,ccuapowerdownRoadCount = get_roadNum(ccuList)
    rsuLossRoad,rsuLossRoadCount,rsuofflineRoad,rsuofflineRoadCount,rsupowerdownRoad,rsupowerdownRoadCount = get_roadNum(rsuList)
    radarLossRoad,radarLossRoadCount,radarofflineRoad,radarofflineRoadCount,radarpowerdownRoad,radarpowerdownRoadCount = get_roadNum(radarList)
    idcLossRoad,idcBigLoopList,idcUniqueBigLoop,idcUniqueBigLoopCount = get_idcroadNum(idcList)
    idcBigLoopDictValue = dict(zip(idcUniqueBigLoop,idcUniqueBigLoopCount))#字典大环
   
    # print("np.hstack(cameraLossRoad,cameraLossRoadCount)",np.hstack((cameraLossRoad,cameraLossRoadCount)))
    # print("np.dstack(cameraLossRoad,cameraLossRoadCount)",np.dstack((cameraLossRoad,cameraLossRoadCount)))
    # print("np.vstack(cameraLossRoad,cameraLossRoadCount)",np.vstack((cameraLossRoad,cameraLossRoadCount)))

    # yy = [1,2,3,4,5]
    # bb = [2,5]

    # print("cameraRoad",cameraLossRoad)
    # print("ccuRoad",ccuLossRoad)
    # print("rsuRoad",rsuLossRoad)
    # print("radarRoad",radarLossRoad)
    # print("idcRoad",idcLossRoad)

    # print("test",list(set(yy).intersection(set(bb))))
    # print("test-1",list(set(bb).intersection(set(yy))))
    
    compare1 = list(set(idcLossRoad).intersection(set(cameraLossRoad)))
    compare2 = list(set(compare1).intersection(set(ccuLossRoad)))
    compare3 = list(set(compare2).intersection(set(rsuLossRoad)))
    compare_all = list(set(compare3).intersection(set(radarLossRoad)))

    # print("compare1",compare1)
    # print("compare2",compare2)
    # print("compare3",compare3)    
    # print("compare4",compare4)

    compare_idcLossRoad_ccuLossRoad = list(set(idcLossRoad).intersection(set(ccuLossRoad)))
    compare_idcLossRoad_rsuLossRoad = list(set(idcLossRoad).intersection(set(rsuLossRoad)))
    compare_idcLossRoad_radarLossRoad = list(set(idcLossRoad).intersection(set(radarLossRoad)))


    compare_camera_ccu = list(set(cameraLossRoad).intersection(set(ccuLossRoad)))
    compare_camera_ccu_rsu = list(set(compare_camera_ccu).intersection(set(rsuLossRoad)))
    compare_camera_ccu_rsu_radar = list(set(compare_camera_ccu_rsu).intersection(set(radarLossRoad)))
    # print("compare_idcLossRoad_cameraLossRoad",compare1)
    # print("compare_idcLossRoad_ccuLossRoad",compare_idcLossRoad_ccuLossRoad)
    # print("compare_idcLossRoad_rsuLossRoad",compare_idcLossRoad_rsuLossRoad)
    # print("compare_idcLossRoad_radarLossRoad",compare_idcLossRoad_radarLossRoad)
    # print("compare_all",compare_all)
    
    filetxt = open("丢包路口分析结果"+dateM+dateD+'.txt','w')
    filetxt.write('相机丢包路口号:'+str(len(cameraLossRoad))+'个'+str(sorted(cameraLossRoad)))
    filetxt.write('\n')
    filetxt.write('串口丢包路口号:'+str(len(ccuLossRoad))+'个'+str(sorted(ccuLossRoad)))
    filetxt.write('\n')
    filetxt.write('RSU丢包路口号:'+str(len(rsuLossRoad))+'个'+str(sorted(rsuLossRoad)))
    filetxt.write('\n')
    filetxt.write('雷达丢包路口号:'+str(len(radarLossRoad))+'个'+str(sorted(radarLossRoad)))
    filetxt.write('\n')
    filetxt.write('机房丢包路口号:'+str(len(idcLossRoad))+'个'+str(sorted(idcLossRoad)))
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('相机掉线路口号:'+str(len(cameraofflineRoad))+'个'+str(sorted(cameraofflineRoad)))
    filetxt.write('\n')
    filetxt.write('串口掉线路口号:'+str(len(ccuofflineRoad))+'个'+str(sorted(ccuofflineRoad)))
    filetxt.write('\n')
    filetxt.write('RSU掉线路口号:'+str(len(rsuofflineRoad))+'个'+str(sorted(rsuofflineRoad)))
    filetxt.write('\n')
    filetxt.write('雷达掉线路口号:'+str(len(radarofflineRoad))+'个'+str(sorted(radarofflineRoad)))
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('相机/串口/RSU/雷达交集:'+str(len(compare_camera_ccu_rsu_radar))+'个'+str(sorted(compare_camera_ccu_rsu_radar)))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('机房和相机路口号交集:'+str(len(compare1))+'个'+str(sorted(compare1)))
    filetxt.write('\n')
    filetxt.write('机房和串口路口号交集:'+str(len(compare_idcLossRoad_ccuLossRoad))+'个'+str(sorted(compare_idcLossRoad_ccuLossRoad)))
    filetxt.write('\n')
    filetxt.write('机房和RSU路口号交集:'+str(len(compare_idcLossRoad_rsuLossRoad))+'个'+str(sorted(compare_idcLossRoad_rsuLossRoad)))
    filetxt.write('\n')
    filetxt.write('机房和雷达路口号交集:'+str(len(compare_idcLossRoad_radarLossRoad))+'个'+str(sorted(compare_idcLossRoad_radarLossRoad)))
    filetxt.write('\n')
    filetxt.write('机房和所有设备路口号交集:'+str(len(compare_all))+'个'+str(sorted(compare_all)))
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('RSCU到机房丢包涉及的大环信息')
    filetxt.write('\n')
    for bigLoopIndex in idcBigLoopDictValue:
        filetxt.write('大环号:'+str(bigLoopIndex)+'     '+'不同RSCU丢包共计:'+str(idcBigLoopDictValue[bigLoopIndex]))
        filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')
        # print('大环号:'+str(bigLoopIndex)+'     '+'丢的路口次数:'+str(idcBigLoopDictValue[bigLoopIndex]))

    filetxt.write('相机涉及重点路口号:'+str(len(cameraImpactRoadList))+'个'+str(sorted(cameraImpactRoadList)))
    filetxt.write('\n')
    filetxt.write('串口涉及重点路口号:'+str(len(ccuImpactRoadList))+'个'+str(sorted(ccuImpactRoadList)))
    filetxt.write('\n')
    filetxt.write('RSU涉及重点路口号:'+str(len(rsuImpactRoadList))+'个'+str(sorted(rsuImpactRoadList)))
    filetxt.write('\n')
    filetxt.write('雷达涉及重点路口号:'+str(len(radarImpactRoadList))+'个'+str(sorted(radarImpactRoadList)))
    filetxt.write('\n')
    filetxt.write('机房涉及重点路口号:'+str(len(idcImpactRoadList))+'个'+str(sorted(idcImpactRoadList)))
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('相机涉及新四跨路口号:'+str(len(cameraXinSiRoadList))+'个'+str(sorted(cameraXinSiRoadList)))
    filetxt.write('\n')
    filetxt.write('串口涉及新四跨路口号:'+str(len(ccuXinSiRoadList))+'个'+str(sorted(ccuXinSiRoadList)))
    filetxt.write('\n')
    filetxt.write('RSU涉及新四跨路口号:'+str(len(rsuXinSiRoadList))+'个'+str(sorted(rsuXinSiRoadList)))
    filetxt.write('\n')
    filetxt.write('雷达涉及新四跨路口号:'+str(len(radarXinSiRoadList))+'个'+str(sorted(radarXinSiRoadList)))
    filetxt.write('\n')
    filetxt.write('机房涉及新四跨路口号:'+str(len(idcXinSiRoadList))+'个'+str(sorted(idcXinSiRoadList)))
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.close()

def funnel_check(devcieList): #漏斗检测
    listAll=[]
    funnelList = []
    funnelNoList = []
    funnelYesList = []
    # funnelListString = []
    stringAllList = []
    for i,valueR in enumerate(devcieList):
        # print(valueR)
        roleNume = valueR.split(' ')[2]
        # print("roleNume",roleNume[4:7])
        # deviceIP = valueR.split('-')[2].split('_')[1].split(' ')[0]
        # deviceName = valueR.split(' ')[2]
        # checkReult = valueR.split(' ')[1]
        # pingLoss =valueR.split(' ')[3]
        # deviceState = valueR.split(' ')[4] 
        listAll.append(int(roleNume[4:7]))
    funnelAllRoad,s ,t,w= np.unique(listAll,return_counts=True,return_index=True,return_inverse=True)
    # print("column",column)
    # print("b",b)
    # print("s",s)
    # print("t",t)
    # print("w",w)
    # print("len-listall",len(b))
    for findex in funnelAllRoad:
        funnelListString = []
        for i1,valueR1 in enumerate(devcieList):
            
            roleNume1 = valueR1.split(' ')[2]
            roleNume2 = int(roleNume1[4:7])
            checkNum = int(valueR1.split(' ')[1])
            checkString = valueR1.split(' ')[4]
            # print("roleNume1",roleNume1)
            # print("valueR1",valueR1)
            # print("findex",findex)
            # n=0
            # print("funnelList",funnelList)
            if int(findex) == roleNume2:
                # print("roleNume2",roleNume2)
                # print("valueR1.split(' ')[1]",valueR1.split(' ')[1])
                # print("valueR1.split(' ')[4]",valueR1.split(' ')[4])
                # funnelString = str(roleNume2)+" " + str(valueR1.split(' ')[4])
                funnelListString.append(valueR1.split(' ')[4]) #获取路口的掉线分析
                # funnelListString.append(funnelString)
                if checkNum!=0 or checkString != '未掉过线': #选取未通过的路口
                    # print(findex)
                     funnelNoList.append(roleNume2)
                # print('valueR1',valueR1)
                    # funnelList.append(n)
        b1,s1,t1,w1= np.unique(funnelListString,return_counts=True,return_index=True,return_inverse=True)
        deviceHistoryStatus = "|".join(b1) #把字符串用｜分开
        funnelString = str(findex)+" " + str(deviceHistoryStatus)#组合下
    
        stringAllList.append(funnelString) #获取路口字符串
        # print("funnelString",funnelString)
        # print("funnelListString",b1)
    # print("stringAllList",stringAllList)
    funnelNoRoad,s1 ,t1,w1= np.unique(funnelNoList,return_counts=True,return_index=True,return_inverse=True)#b1为未通过漏斗测试的
        
    # print("funnelAllRoad",funnelAllRoad)
    # print("len-funnelAllRoad",len(funnelAllRoad))
    # print("funnelNoRoad",funnelNoRoad)  
    # print("len-funnelNoRoad",len(funnelNoRoad))

    # print("len-funnelList",len(funnelList))
    # return funnelList
    # compare1 = list(b1.intersection(b))
    funnelYesRoad = np.setdiff1d(funnelAllRoad,funnelNoRoad) #通过漏斗测试的路口

    # print("funnelYesRoad",funnelYesRoad)
    # print("len-funnelYesRoad",len(funnelYesRoad))
    return funnelAllRoad,funnelYesRoad,funnelNoRoad,stringAllList
    
def funnel_check_test(funnnelNum,funnelNoList,funnelAllList):
    # funnel_test = [9,15,16,17,22,27,34,37,38,49,54,68,74,76,77,87,103,105,116,119,125,141,146,161,208,216,227,236,238,242,245,263,264,267,268,282,286,298,303]
    if funnnelNum in funnelNoList:
        return 0
    elif funnnelNum in funnelAllList:
        return 1
    else:
        return 2 #监控缺失数据

def funnel_check_string(funnnelNum,funnelString):
    for funnnelIndex in funnelString:
        # print("funnnelIndex",funnnelIndex)
        # print("funnnelIndex[1]",funnnelIndex[1])
        # print("funnnelIndex.split(' ')[0]",funnnelIndex.split(' ')[0])
        # print("funnnelIndex.split(' ')[1]",funnnelIndex.split(' ')[1])
        # print("funnnelIndex.split(' ')[3]",funnnelIndex.split(' ')[3])
        if funnnelNum ==  int(funnnelIndex.split(' ')[0]):
            return funnnelIndex.split(' ')[1]

def write_excel_funnel(list_value):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    ws.cell(row=1,column=1).value='路口号'
    ws.cell(row=1,column=2).value='整体是否通过'
    ws.cell(row=1,column=3).value='相机是否通过'
    ws.cell(row=1,column=4).value='串口是否通过'
    ws.cell(row=1,column=5).value='RSU是否通过'
    ws.cell(row=1,column=6).value='雷达是否通过'
    ws.cell(row=1,column=7).value='机房是否通过'
    ws.cell(row=1,column=8).value='NTP是否通过'
    ws.cell(row=1,column=9).value='GPS是否通过'

    ws.cell(row=1,column=10).value='相机原因总结'
    ws.cell(row=1,column=11).value='串口原因总结'
    ws.cell(row=1,column=12).value='RSU原因总结'
    ws.cell(row=1,column=13).value='雷达原因总结'
    ws.cell(row=1,column=14).value='机房原因总结'

    ws.cell(row=1,column=15).value='0:不通过、1:通过、2:缺失监控数据'

    for i,valueR in enumerate(list_value):
        # print(valueR)
        roleNume = int(valueR.split(' ')[0])
        cameraCheck = int(valueR.split(' ')[1])
        ccuCheck = int(valueR.split(' ')[2])
        rsuCheck = int(valueR.split(' ')[3])
        radarCheck = int(valueR.split(' ')[4])
        idcCheck = int(valueR.split(' ')[5])
        ntpCheck = int(valueR.split(' ')[6])
        gpsCheck = int(valueR.split(' ')[7])
        allCheck = cameraCheck & ccuCheck & rsuCheck & idcCheck & ntpCheck & gpsCheck

        cameraCheckString = valueR.split(' ')[8]
        ccuCheckString = valueR.split(' ')[9]
        rsuCheckString = valueR.split(' ')[10]
        radarCheckString = valueR.split(' ')[11]
        idcCheckString = valueR.split(' ')[12]

        ws.cell(row=i+2,column=1).value=roleNume
        ws.cell(row=i+2,column=2).value=allCheck
        ws.cell(row=i+2,column=3).value=cameraCheck
        ws.cell(row=i+2,column=4).value=ccuCheck
        ws.cell(row=i+2,column=5).value=rsuCheck
        ws.cell(row=i+2,column=6).value=radarCheck
        ws.cell(row=i+2,column=7).value=idcCheck
        ws.cell(row=i+2,column=8).value=ntpCheck
        ws.cell(row=i+2,column=9).value=gpsCheck
        ws.cell(row=i+2,column=10).value=cameraCheckString
        ws.cell(row=i+2,column=11).value=ccuCheckString
        ws.cell(row=i+2,column=12).value=rsuCheckString
        ws.cell(row=i+2,column=13).value=radarCheckString
        ws.cell(row=i+2,column=14).value=idcCheckString
       
        

    wb.save('亦庄304个路口全量分析结果'+dateM+dateD+'.xlsx')



def save_funnel(cameraAll,cameraYes,cameraNo,cameraString,ccuAll,ccuYes,ccuNo,ccuString,rsuAll,rsuYes,rsuNo,rsuString,radarAll,radarYes,radarNo,radarString,idcAll,idcYes,idcNo,idcString,ntpAll,ntpYes,ntpNo,gpsAll,gpsYes,gpsNo):
    c1 = np.intersect1d(cameraYes,ccuYes)
    c2 = np.intersect1d(c1,rsuYes)
    c3 = np.intersect1d(c2,ntpYes)
    c4 = np.intersect1d(c3,gpsYes)
    c5 = np.intersect1d(c4,idcYes)
    funnelTestList = []
    a1 = np.append(cameraAll,ccuAll)
 
    a3 = np.setdiff1d(ccuAll,cameraAll)
    a4 = np.setdiff1d(rsuAll,cameraAll)
    a5 = np.setdiff1d(radarAll,cameraAll)
    a6 = np.setdiff1d(idcAll,cameraAll)
    a7 = np.setdiff1d(ntpAll,cameraAll)
    a8 = np.setdiff1d(gpsAll,cameraAll)

    y1 = np.intersect1d(cameraAll,ccuAll)
    y2 = np.intersect1d(y1,rsuAll)
    y3 = np.intersect1d(y2,radarAll)
    y4 = np.intersect1d(y3,idcAll)
    y5 = np.intersect1d(y4,ntpAll)
    y6 = np.intersect1d(y5,gpsAll)
    # print("cameraAll+ccuAll",a1)
    # print("cameraAll==ccuAll",cameraAll==ccuAll)
    # print("cameraAll-intersert-ccuAll",y2)
    # print("len-cameraAll-intersert-ccuAll",len(y2))
    # print("All-intersert-all",y6)
    # print("len-All-intersert-all",len(y6))
    # print("len(a2)",len(a2))
    # print("len(cameraAll)",len(cameraAll))
    # print("len(ccuAll)",len(ccuAll))
    # print("cameraAll-diff-ccuAll",a3)
    # print("cameraAll-diff-rsuAll",a4)
    # print("cameraAll-diff-radarAll",a5)
    # print("cameraAll-diff-idcAll",a6)
    # print("cameraAll-diff-ntpAll",a7)
    # print("cameraAll-diff-gpsAll",a8)
    # funnel_test=[9,11,15,16,17,21,22,27,34,37,38,44,49,54,68,74,76,77,87,103,105,110,116,119,125,126,141,146,150,157,161,208,216,227,236,238,242,245,246,248,263,264,267,268,282,286,287,295,298,303]
    funnel_test= [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
    41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79,
    80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114,
    115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 
    146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176,
    177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207,
    208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 
    239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 
    270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285, 286, 287, 288, 289, 290, 291, 292, 293, 294, 295, 296, 297, 298, 299, 300, 
    301, 302, 303, 304]
    for findex in funnel_test:
        findexV  = str(findex) +' '+ str(funnel_check_test(findex,cameraNo,cameraAll)) +' '+ str(funnel_check_test(findex,ccuNo,ccuAll)) +' '+ str(funnel_check_test(findex,rsuNo,rsuAll))+' '+ str(funnel_check_test(findex,radarNo,radarAll))+' '+ str(funnel_check_test(findex,idcNo,idcAll))+' '+str(funnel_check_test(findex,ntpNo,ntpAll))+' '+str(funnel_check_test(findex,gpsNo,gpsAll)) + ' ' + str(funnel_check_string(findex,cameraString))+ ' ' + str(funnel_check_string(findex,ccuString))+ ' ' + str(funnel_check_string(findex,rsuString))+ ' ' + str(funnel_check_string(findex,radarString))+ ' ' + str(funnel_check_string(findex,idcString))
        funnelTestList.append(findexV)
        # print("findexV",findexV)
    write_excel_funnel(funnelTestList)

    funnel_check_string(1,cameraString)

    filetxt = open("24h稳定通过分析结果"+dateM+dateD+'.txt','w')
    filetxt.write('24h稳定通过路口号（相机+串口+RSU+NTP校时+GPS校时+机房）:'+str(len(c5))+'个'+str(c5))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('24h稳定通过路口号（不包含大环、校时）:'+str(len(c2))+'个'+str(c2))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')
    
    filetxt.write('相机通过率:'+str(round((len(cameraYes)/len(cameraAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('串口通过率:'+str(round((len(ccuYes)/len(ccuAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('RSU通过率:'+str(round((len(rsuYes)/len(rsuAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('雷达通过率:'+str(round((len(radarYes)/len(radarAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('机房通过率:'+str(round((len(idcYes)/len(idcAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('NTP通过率:'+str(round((len(ntpYes)/len(ntpAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('GPS通过率:'+str(round((len(gpsYes)/len(gpsAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')


    filetxt.write('相机路口总数:'+str(len(cameraAll))+'个'+str(cameraAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('相机通过路口号:'+str(len(cameraYes))+'个'+str(cameraYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('相机不通过路口号:'+str(len(cameraNo))+'个'+str(cameraNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('相机通过率:'+str(round((len(cameraYes)/len(cameraAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('串口路口总数:'+str(len(ccuAll))+'个'+str(ccuAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('串口通过路口号:'+str(len(ccuYes))+'个'+str(ccuYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('串口不通过路口号:'+str(len(ccuNo))+'个'+str(ccuNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('串口通过率:'+str(round((len(ccuYes)/len(ccuAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('RSU路口总数:'+str(len(rsuAll))+'个'+str(rsuAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('RSU通过路口号:'+str(len(rsuYes))+'个'+str(rsuYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('RSU不通过路口号:'+str(len(rsuNo))+'个'+str(rsuNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('RSU通过率:'+str(round((len(rsuYes)/len(rsuAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('雷达路口总数:'+str(len(radarAll))+'个'+str(radarAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('雷达通过路口号:'+str(len(radarYes))+'个'+str(radarYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('雷达不通过路口号:'+str(len(radarNo))+'个'+str(radarNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('雷达通过率:'+str(round((len(radarYes)/len(radarAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('机房路口总数:'+str(len(idcAll))+'个'+str(idcAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('机房通过路口号:'+str(len(idcYes))+'个'+str(idcYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('机房不通过路口号:'+str(len(idcNo))+'个'+str(idcNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('机房通过率:'+str(round((len(idcYes)/len(idcAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('NTP路口总数:'+str(len(ntpAll))+'个'+str(ntpAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('NTP通过路口号:'+str(len(ntpYes))+'个'+str(ntpYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('NTP不通过路口号:'+str(len(ntpNo))+'个'+str(ntpNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('NTP通过率:'+str(round((len(ntpYes)/len(ntpAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

    filetxt.write('GPS路口总数:'+str(len(gpsAll))+'个'+str(gpsAll))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('GPS通过路口号:'+str(len(gpsYes))+'个'+str(gpsYes))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('GPS不通过路口号:'+str(len(gpsNo))+'个'+str(gpsNo))
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('GPS通过率:'+str(round((len(gpsYes)/len(gpsAll)*100),2))+'%')
    filetxt.write('\n')
    filetxt.write('\n')
    filetxt.write('\n')

def  rscu_time_diff(excel_name):
     df =pd.read_excel(excel_name)#读取excel
     timeList = []
    # listLossAll = []
    # listOfflinePowerdown = []
    # n = 1
    # lossCount = 0
    # lossCheck = 0
    # offLineCount = 0 #离线计数
    # powerDownCount = 0 #掉电计数
    # norMalCount = 0

     clos =[i for i in df.columns if i not in ['Time']]
     df2 = df[clos]
     for column in df2.columns:
        # n= n+1 
        # listallresult = df2[column].tolist()
        a = np.array(df2[column].tolist())
        # vu = a[(np.where((a>0) & (a<1000)))]
        # print(np.where(a == 1000))
        # print(np.size(np.where(a == 1000)))
        timeMin = np.nanmin(a)
        timeMax= np.nanmax(a)
        if timeMin < -10 or timeMax >10:
            columnV=column + ' ' + '1'+' ' + str(timeMin)+' '+str(timeMax)
            timeList.append(columnV)
        else:
            columnV=column + ' ' + '0'+' ' + str(timeMin)+' '+str(timeMax)
            timeList.append(columnV)
        # print("timeMin",timeMin)
        # print("timeMax",timeMax)
        # print("column",column)
        # timeMin = np.size(np.where(a == 0)) 
        # numoffline = numNan + num1000
    #  print("timeList",timeList)
     return timeList


def funnel_time_check(devcieList): #漏斗检测
    listAll=[]
    funnelList = []
    funnelNoList = []
    funnelYesList = []
    for i,valueR in enumerate(devcieList):
        # print(valueR)
        roleNume = valueR.split(' ')[0]
        # print("roleNume",roleNume[4:7])
        # deviceIP = valueR.split('-')[2].split('_')[1].split(' ')[0]
        # deviceName = valueR.split(' ')[2]
        # checkReult = valueR.split(' ')[1]
        # pingLoss =valueR.split(' ')[3]
        # deviceState = valueR.split(' ')[4] 
        listAll.append(int(roleNume[4:7]))
    funnelAllRoad,s ,t,w= np.unique(listAll,return_counts=True,return_index=True,return_inverse=True)
    # print("column",column)
    # print("b",b)
    # print("s",s)
    # print("t",t)
    # print("w",w)
    # print("len-listall",len(b))
    for findex in funnelAllRoad:
        for i1,valueR1 in enumerate(devcieList):
            # funnelList = []
            roleNume1 = valueR1.split(' ')[0]
            roleNume2 = int(roleNume1[4:7])
            checkNum = int(valueR1.split(' ')[1])
            # checkString = valueR1.split(' ')[4]
            # print("roleNume1",roleNume1)
            # print("valueR1",valueR1)
            # print("findex",findex)
            # n=0
            # print("funnelList",funnelList)
            if int(findex) == roleNume2:
                # print("valueR1.split(' ')[1]",valueR1.split(' ')[1])
                # print("valueR1.split(' ')[4]",valueR1.split(' ')[4])
                if checkNum==1:
                    # print(findex)
                     funnelNoList.append(roleNume2)
                # print('valueR1',valueR1)
                    # funnelList.append(n)
    # print("funnelList",funnelList)
    funnelNoRoad,s1 ,t1,w1= np.unique(funnelNoList,return_counts=True,return_index=True,return_inverse=True)#b1为未通过漏斗测试的
        
    # print("funnelAllRoad",funnelAllRoad)
    # print("len-funnelAllRoad",len(funnelAllRoad))
    # print("funnelNoRoad",funnelNoRoad)  
    # print("len-funnelNoRoad",len(funnelNoRoad))

    # print("len-funnelList",len(funnelList))
    # return funnelList
    # compare1 = list(b1.intersection(b))
    funnelYesRoad = np.setdiff1d(funnelAllRoad,funnelNoRoad) #通过漏斗测试的路口

    # print("funnelYesRoad",funnelYesRoad)
    # print("len-funnelYesRoad",len(funnelYesRoad))
    return funnelAllRoad,funnelYesRoad,funnelNoRoad

def write_excel_time(excel_name,excel_sheet,excel_save,list_value):
    wb= openpyxl.load_workbook(excel_name)
    ws1 = wb[excel_sheet]
    #ws.delete_cols(1)
    # for cameravalue  in list_value:
    #     ws.cell(row=291,column=int(cameravalue.split(' ')[0])).value=int(cameravalue.split(' ')[1])
    # bigLoopList = []
    
    ws1.cell(row=1,column=1).value='路口号'
    ws1.cell(row=1,column=2).value='设备号'
    #ws1.cell(row=1,column=3).value='IP'
    ws1.cell(row=1,column=3).value='校时误差最小值'
    ws1.cell(row=1,column=4).value='校时误差最大值'
    ws1.cell(row=1,column=5).value='是否异常'
    # ws1.cell(row=1,column=6).value='状态'
    for i,valueR in enumerate(list_value):
        # print(valueR)
        roleNume = valueR.split(' ')[0]
        # print("roleNume",roleNume)
        # print("valueR",valueR)
        deviceName = valueR.split(' ')[0]
        checkStatus = valueR.split(' ')[1]
        timeMin =valueR.split(' ')[2]
        timeMax =valueR.split(' ')[3]
        # print("timeMin",timeMin)
        # deviceState =valueR.split(' ')[4]
        ws1.cell(row=i+2,column=1).value=int(roleNume[4:7])
        ws1.cell(row=i+2,column=2).value=deviceName
        # ws1.cell(row=i+2,column=3).value=deviceIP
        ws1.cell(row=i+2,column=3).value=timeMin
        ws1.cell(row=i+2,column=4).value=timeMax
        ws1.cell(row=i+2,column=5).value=int(checkStatus)

    wb.save(excel_save)



cameraCheckValue = 90 #相机参数检测指标
normalCheckValue = 30
time_24 = 24

excel_camera_name = 'camera.xlsx' #源表
excel_camera_sheet = 'RSCU到感知相机通信-data-as-seriestocol'
excel_camera_save = 'cameracheck1.xlsx'
excel_camera_sheet1 = 'Sheet1'
excel_camera_save1 = '相机'+dateM+dateD+'.xlsx'

excel_ccu_name = 'ccu.xlsx' #源表
excel_ccu_sheet = 'RSCU到CCU丢包率-data-as-seriestocol'
excel_ccu_save = 'ccucheck1.xlsx'
excel_ccu_sheet1 = 'Sheet1'
excel_ccu_save1 = '串口'+dateM+dateD+'.xlsx'

excel_rsu_name = 'rsu.xlsx' #源表
excel_rsu_sheet = 'RSCU到RSU丢包率-data-as-seriestocol'
excel_rsu_save = 'rsucheck1.xlsx'
excel_rsu_sheet1 = 'Sheet1'
excel_rsu_save1 = 'RSU'+dateM+dateD+'.xlsx'

excel_radar_name = 'radar.xlsx' #源表
excel_radar_sheet = 'RSCU到感知雷达通信-data-as-seriestocol'
excel_radar_save = 'radarcheck1.xlsx'
excel_radar_sheet1 = 'Sheet1'
excel_radar_save1 = '雷达'+dateM+dateD+'.xlsx'

excel_idc_name = 'idc.xlsx' #源表
excel_idc_sheet = 'RSCU到机房丢包率-data-as-seriestocolu'
excel_idc_save = 'idccheck1.xlsx'
excel_idc_sheet1 = 'Sheet1'
excel_idc_save1 = '机房'+dateM+dateD+'.xlsx'

excel_rscu_name = 'rscu.xlsx' #源表
excel_rscu_sheet = 'RSCU通电通网状态（二期）-data-as-seriesto'
excel_rscu_save = 'rscuheck1.xlsx'
excel_rscu_sheet1 = 'Sheet1'
excel_rscu_save1 = 'RSCU状态'+dateM+dateD+'.xlsx'

excel_ntp_name = 'ntp.xlsx' #源表
excel_ntp_sheet = 'rscu_time_ntp_diff（与中心机房时差）-dat'
# excel_rscu_save = 'rscuheck1.xlsx'
excel_ntp_sheet1 = 'Sheet1'
excel_ntp_save1 = 'NTP状态'+dateM+dateD+'.xlsx'

excel_gps_name = 'gps.xlsx' #源表
excel_gps_sheet = 'rscu_time_gps_diff（与GPS时差）-data'
# excel_rscu_save = 'rscuheck1.xlsx'
excel_gps_sheet1 = 'Sheet1'
excel_gps_save1 = 'GPS状态'+dateM+dateD+'.xlsx'

excel_switch_name = 'switch.xlsx' #源表
excel_switch_sheet = 'RSCU到交换机通信-data-as-seriestocolu'
excel_switch_save = 'switchcheck1.xlsx'
excel_switch_sheet1 = 'Sheet1'
excel_switch_save1 = '交换机'+dateM+dateD+'.xlsx'


# add_sheet(excel_camera_name) #增加sheet1
# add_sheet(excel_ccu_name)
# add_sheet(excel_rsu_name)
# add_sheet(excel_radar_name)
# add_sheet(excel_idc_name)
# add_sheet(excel_rscu_name)
# add_sheet(excel_ntp_name)
# add_sheet(excel_gps_name)

ntpList = rscu_time_diff(excel_ntp_name)
ntpFunnelAllRoad,ntpFunnelYesRoad,ntpFunnelNoRoad= funnel_time_check(ntpList)
write_excel_time(excel_ntp_name,excel_ntp_sheet1,excel_ntp_save1,ntpList)

gpsList = rscu_time_diff(excel_gps_name) #校时检查
gpsFunnelAllRoad,gpsFunnelYesRoad,gpsFunnelNoRoad= funnel_time_check(gpsList)
write_excel_time(excel_gps_name,excel_gps_sheet1,excel_gps_save1,gpsList)

rscuList = get_rscu_status(excel_rscu_name) #RSCU状态检测
write_excel_rscu(excel_rscu_name,excel_rscu_sheet1,excel_rscu_save1,rscuList)

cameraList,cameraCount,cameraLossCount,cameraLossCheck,cameraLossNoCheck,cameraListLossAll,cameraOffLineCount,cameraPowerDownCount,cameraNorMalCount = get_loss_value(excel_camera_name,cameraCheckValue)
cameraImpactRoad,cameraXinSiRoad,cameraImpactRoadList,cameraXinSiRoadList= impact_road(cameraListLossAll)
cameraFunnelAllRoad,cameraFunnelYesRoad,cameraFunnelNoRoad,cameraFunnelString = funnel_check(cameraList) #漏斗检测
write_excel(excel_camera_name,excel_camera_sheet,excel_camera_save,cameraList)
write_excel1(excel_camera_save,excel_camera_sheet1,excel_camera_save1,cameraList)


ccuList,ccuCount,ccuLossCount,ccuLossCheck,ccuLossNoCheck,ccuListLossAll,ccuOffLineCount,ccuPowerDownCount,ccuNorMalCount = get_loss_value(excel_ccu_name,normalCheckValue)
ccuImpactRoad,ccuXinSiRoad,ccuImpactRoadList,ccuXinSiRoadList= impact_road(ccuListLossAll)
ccuFunnelAllRoad,ccuFunnelYesRoad,ccuFunnelNoRoad,ccuFunnelString = funnel_check(ccuList) #漏斗检测
write_excel(excel_ccu_name,excel_ccu_sheet,excel_ccu_save,ccuList)
write_excel1(excel_ccu_save,excel_ccu_sheet1,excel_ccu_save1,ccuList)

rsuList,rsuCount,rsuLossCount,rsuLossCheck,rsuLossNoCheck,rsuListLossAll,rsuOffLineCount,rsuPowerDownCount,rsuNorMalCount = get_loss_value(excel_rsu_name,normalCheckValue)
rsuImpactRoad,rsuXinSiRoad,rsuImpactRoadList,rsuXinSiRoadList= impact_road(rsuListLossAll)
rsuFunnelAllRoad,rsuFunnelYesRoad,rsuFunnelNoRoad,rsuFunnelString = funnel_check(rsuList)
write_excel(excel_rsu_name,excel_rsu_sheet,excel_rsu_save,rsuList)
write_excel1(excel_rsu_save,excel_rsu_sheet1,excel_rsu_save1,rsuList)

radarList,radarCount,radarLossCount,radarLossCheck,radarLossNoCheck,radarListLossAll,radarOffLineCount,radarPowerDownCount,radarNorMalCount = get_loss_value(excel_radar_name,normalCheckValue)
radarImpactRoad,radarXinSiRoad,radarImpactRoadList,radarXinSiRoadList= impact_road(radarListLossAll)
radarFunnelAllRoad,radarFunnelYesRoad,radarFunnelNoRoad,radarFunnelString = funnel_check(radarList)
write_excel(excel_radar_name,excel_radar_sheet,excel_radar_save,radarList)
write_excel1(excel_radar_save,excel_radar_sheet1,excel_radar_save1,radarList)


idcList,idcCount,idcLossCount,idcLossCheck,idcLossNoCheck,idcListLossAll,idcOffLineCount,idcPowerDownCount,idcNorMalCount = get_loss_value(excel_idc_name,normalCheckValue)
idcImpactRoad,idcXinSiRoad,idcImpactRoadList,idcXinSiRoadList= impact_road_idc(idcListLossAll)
idcFunnelAllRoad,idcFunnelYesRoad,idcFunnelNoRoad,idcFunnelString = funnel_check(idcList)
write_excel(excel_idc_name,excel_idc_sheet,excel_idc_save,idcList)
write_excel_idc(excel_idc_save,excel_idc_sheet1,excel_idc_save1,idcList)

switchList,switchCount,switchLossCount,switchLossCheck,switchLossNoCheck,switchListLossAll,switchOffLineCount,switchPowerDownCount,switchNorMalCount = get_loss_value(excel_switch_name,normalCheckValue)
switchImpactRoad,switchXinSiRoad,switchImpactRoadList,switchXinSiRoadList= impact_road_idc(switchListLossAll)
switchFunnelAllRoad,switchFunnelYesRoad,switchFunnelNoRoad,switchFunnelString = funnel_check(switchList)
write_excel(excel_switch_name,excel_switch_sheet,excel_switch_save,switchList)
write_excel_idc(excel_switch_save,excel_switch_sheet1,excel_switch_save1,switchList)

countSum = cameraCount + ccuCount + rsuCount + radarCount + idcCount + switchCount
lossCountSum = cameraLossCount + ccuLossCount + rsuLossCount + radarLossCount + idcLossCount +switchLossCount
lossNoCheckSum = cameraLossNoCheck + ccuLossNoCheck + rsuLossNoCheck + radarLossNoCheck + idcLossNoCheck + switchLossNoCheck
lossCheckSum = cameraLossCheck + ccuLossCheck + rsuLossCheck + radarLossCheck + idcLossCheck + switchLossCheck
impactRoadSum = cameraImpactRoad + ccuImpactRoad + rsuImpactRoad + radarImpactRoad + idcImpactRoad + switchImpactRoad
xinSiRoadSum = cameraXinSiRoad + ccuXinSiRoad + rsuXinSiRoad + radarXinSiRoad + idcXinSiRoad + switchXinSiRoad
offlineSum = cameraOffLineCount + ccuOffLineCount + rsuOffLineCount + radarOffLineCount + idcOffLineCount + switchOffLineCount
powerdownSum = cameraPowerDownCount + ccuPowerDownCount + rsuPowerDownCount + radarPowerDownCount + idcPowerDownCount + switchPowerDownCount

allRustlList = pd.DataFrame({'监控总条数':[cameraCount,ccuCount,rsuCount,radarCount,idcCount,switchCount,countSum],
'丢包设备总条数':[cameraLossCount,ccuLossCount,rsuLossCount,radarLossCount,idcLossCount,switchLossCount,lossCountSum],
'丢包不排查':[cameraLossNoCheck,ccuLossNoCheck,rsuLossNoCheck,radarLossNoCheck,idcLossNoCheck,switchLossNoCheck,lossNoCheckSum],
'丢包排查':[cameraLossCheck,ccuLossCheck,rsuLossCheck,radarLossCheck,idcLossCheck,switchLossCheck,lossCheckSum],
'涉及重点路口数量':[cameraImpactRoad,ccuImpactRoad,rsuImpactRoad,radarImpactRoad,idcImpactRoad,switchImpactRoad,impactRoadSum],
'涉及新四跨路口数量':[cameraXinSiRoad,ccuXinSiRoad,rsuXinSiRoad,radarXinSiRoad,idcXinSiRoad,switchXinSiRoad,xinSiRoadSum],
'离线设备条数':[cameraOffLineCount,ccuOffLineCount,rsuOffLineCount,radarOffLineCount,idcOffLineCount,switchOffLineCount,offlineSum],
'掉过线设备条数':[cameraPowerDownCount,ccuPowerDownCount,rsuPowerDownCount,radarPowerDownCount,idcPowerDownCount,switchPowerDownCount,powerdownSum]},index = ["相机","串口设备","RSU","雷达","机房","交换机","共计"])
# print(allRustlList)
allRustlList.to_excel('丢包设备分析结果'+dateM+dateD+'.xlsx')

compare_allRoad(cameraList,ccuList,rsuList,radarList,idcList)

save_funnel(cameraFunnelAllRoad,cameraFunnelYesRoad,cameraFunnelNoRoad,cameraFunnelString,ccuFunnelAllRoad,ccuFunnelYesRoad,ccuFunnelNoRoad,ccuFunnelString,rsuFunnelAllRoad,rsuFunnelYesRoad,rsuFunnelNoRoad,rsuFunnelString,radarFunnelAllRoad,radarFunnelYesRoad,radarFunnelNoRoad,radarFunnelString,idcFunnelAllRoad,idcFunnelYesRoad,idcFunnelNoRoad,idcFunnelString,ntpFunnelAllRoad,ntpFunnelYesRoad,ntpFunnelNoRoad,gpsFunnelAllRoad,gpsFunnelYesRoad,gpsFunnelNoRoad)

os.remove(excel_camera_save)
os.remove(excel_ccu_save)
os.remove(excel_rsu_save)
os.remove(excel_radar_save)
os.remove(excel_idc_save)
os.remove(excel_switch_save)
