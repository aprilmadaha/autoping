import pandas as pd
import time
import operator
import subprocess
import openpyxl
import os
import numpy as np

def getDeviceList(filename): #从excel表里读取objec列的所有数值 按照每个杆子
        df =pd.read_excel(excel_name,keep_default_na=False) #nan转换空

        allList=[]
        # print("df.路口编号",int(df.路口编号))
        roadNumList = list(set(df.路口编号)) #获取路口序号并去重
        roadNumList = [i for i in roadNumList if i != ''] #数组去空

        # goldRoadNume = roadNumList
        # print("roadNumList",roadNumList)

        for roadNum in roadNumList:
                # print("df[df.路口编号==roadNum]",df[df.路口编号==roadNum])
                roadDF = df[df.路口编号==roadNum]       #筛选每个路口比如1，2，3
                poleNumList = roadDF.点位杆号      #获取每个路口,每个杆子号
                # print("poleNumList",poleNumList)
                for poleNum in poleNumList:
                        # print(roadDF[roadDF.点位杆号==poleNum])
                        poleDF=roadDF[roadDF.点位杆号 ==poleNum]     #按照杆子筛选（索引）

                        # roadNum = int(roadNum)
                        roadName = poleDF.路口名称.values[0] #会返回dtype object，使用values，然后在[0]数组取得相应的值
                        # print("roadName",roadName)
                        poleNum = poleNum[0]
                        # print("poleNum",poleNum)
                        pointIn = compositePole(poleNum)        #根据杆子编号得到方向
                        # print(pointIn)

                        cameraNum = checkNan(poleDF.感知枪机.values[0]) #先将nan转为0,在转int
                        # print("poleDF.感知枪机",checkNan(cameraNum))
                        fisheyeNum = checkNan(poleDF.鱼眼相机.values[0])
                        radarNum = checkNan(poleDF.Radar.values[0])
                        # # print("lidarNum",poleDF.LiDAR)
                        lidarNum = checkNan(poleDF.LiDAR.values[0])
                        
                        rsuNum = checkNan(poleDF.RSU.values[0])
                        switchNum = checkNan(poleDF.交换机.values[0])
                        rscuNum = checkNan(poleDF.高配RSCU.values[0])
                        ccuNum = checkNan(poleDF.采集器.values[0])
                     
                        
                        poleDeviceNum = [roadNum,roadName,poleNum,pointIn,cameraNum,fisheyeNum,radarNum,lidarNum,rsuNum,switchNum,rscuNum,ccuNum]
                        allList.append(poleDeviceNum)

        # print(allList)
        return allList,roadNumList
   
def checkNan(numberNan):

        if numberNan== 'nan' or numberNan =='':
                numberNan =0    
        return numberNan

def compositePole(pole):        #根据杆子编号生成方向

        if pole == 'a' or pole == 'b':
                pointInfo = '北侧杆子'
        elif pole == 'c' or pole == 'd':
                pointInfo = '东侧杆子'
        elif pole == 'e' or pole == 'f':
                pointInfo = '南侧杆子'
        elif pole == 'g' or pole == 'h':
                pointInfo = '西侧杆子'

        return pointInfo   
       
def outDeviceList(allDicList,roadIndexList): #根据Bom输出IP表，但是不包含IP
        
        camera='感知枪机'
        fisheye='鱼眼相机'
        lidar='LiDAR'
        radar='Radar'
        rsu='RSU'
        switch='交换机'
        rscu='RSCU'
        ccu='采集器'
        deviceList = []

        for poleDicList in allDicList:
                for roadindex in roadIndexList:
                        if poleDicList[0] ==roadindex: #循环每个路口
                                # print(poleDicList)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[4],camera)#循环增加对应的设备
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[5],fisheye)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[6],radar)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[7],lidar)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[8],rsu)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[9],switch)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[10],rscu)
                                loopDevice(deviceList,poleDicList[0],poleDicList[1],poleDicList[2],poleDicList[3],poleDicList[11],ccu)

        return deviceList
        # print(deviceList)
 
def loopDevice(iplist,roadNum,roadName,poleNum,pointIn,loopNum,deviceName):#根据表里设备的数量循环添加设备
        # a=[]
        for i in range(loopNum):
                deviceInfo = [roadNum,roadName,poleNum,pointIn,deviceName]
                # print(deviceInfo)
                iplist.append(deviceInfo)
       
def assignIP(deviceList,roadIndexList):               #分配地址
        arrayIP =[172,21]               #初始地址分配
        deviceInfoList = []

        for roadindex in roadIndexList:         #为了每个路口遍历后重新读区IP开始地址
                cameraIP = 101                  #摄像头启使地址
                fisheyeIP= 131
                radarIP= 161
                lidarIP= 151
                rsuIP= 11
                switchIP= 21
                rscuIP= 6
                ccuIP= 31

                for device in deviceList:       #遍历每行设备列表
                        if device[0] == roadindex:
                                # print(device)
                                if device[4] == '感知枪机':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],cameraIP)
                                        cameraID = 'camera-'+str(device[0])+'-'+str(cameraIP)      #摄像头ID
                                        cameraIP = cameraIP + 1 #序号自增1
                                        device.append(cameraID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == '鱼眼相机':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],fisheyeIP)
                                        fisheyeID = 'camera-'+str(device[0])+'-'+str(fisheyeIP) 
                                        fisheyeIP = fisheyeIP + 1
                                        device.append(fisheyeID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)#添加网关
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == 'Radar':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],radarIP)
                                        radarID = 'radar-'+str(device[0])+'-'+str(radarIP) 
                                        radarIP = radarIP + 1
                                        device.append(radarID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == 'LiDAR':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],lidarIP)
                                        lidarID = 'lidar-'+str(device[0])+'-'+str(lidarIP) 
                                        lidarIP = lidarIP + 1
                                        device.append(lidarID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == 'RSU':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],rsuIP)
                                        rsuID = 'rsu-'+str(device[0])+'-'+str(rsuIP) 
                                        rsuIP = rsuIP + 1
                                        device.append(rsuID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == '交换机':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],switchIP)
                                        switchID = 'sw-'+str(device[0])+'-'+str(switchIP) 
                                        switchIP = switchIP + 1
                                        device.append(switchID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == 'RSCU':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],rscuIP)
                                        rscuID = 'rscu-'+str(device[0])+'-'+str(rscuIP) 
                                        rscuIP = rscuIP + 1
                                        device.append(rscuID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                elif device[4] == '采集器':
                                        deviceIP,deviceNetMask,deviceGateway = compositeIP(arrayIP[0],arrayIP[1],device[0],ccuIP)
                                        ccuID = 'ccu-'+str(device[0])+'-'+str(ccuIP) 
                                        ccuIP = ccuIP + 1
                                        device.append(ccuID)
                                        device.append(deviceIP)
                                        device.append(deviceNetMask)
                                        device.append(deviceGateway)
                                        # print(device)
                                deviceInfoList.append(device)
        # print(deviceInfoList)
        return deviceInfoList

def compositeIP(ip1,ip2,ip3,ip4):       #合成IP地址，ip分成四段输入

        deviceIP = str(ip1)+'.'+str(ip2)+'.'+str(ip3)+'.'+str(ip4)
        deviceNetMask='255.255.255.0'
        deviceGateway = str(ip1)+'.'+str(ip2)+'.'+str(ip3)+'.'+str(254)

        return deviceIP,deviceNetMask,deviceGateway


excel_name = './whbomliy.xlsx' #源表

dfColumns =['路口编号','路口名称','点位杆号','点位信息','设备类型','设备编号','设备IP地址','子网掩码','网关']

# getDeviceList(excel_name)
allBomList,roadIndexList= getDeviceList(excel_name)

deviceList = outDeviceList(allBomList,roadIndexList)
deviceIpList = assignIP(deviceList,roadIndexList)

# print(deviceIpList)
numpyDeviceIpList = pd.DataFrame(deviceIpList)
numpyDeviceIpList.columns=dfColumns
# print(nuDeviceIpList)
numpyDeviceIpList.to_excel('设备IP规划表.xlsx',index=False)
