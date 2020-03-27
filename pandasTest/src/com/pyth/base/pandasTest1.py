#!/usr/bin/python
# -*- coding: UTF-8 -*-
'''
Created on 2020-1-7

@author: Administrator
'''
import pandas as pd
from datetime import datetime 
##获取Excel文件
class StatisticsUnitTest:
    '''
            统计单元测试Excel文件的数据
    '''
    excelfile='E:/个人/python/base/ZLJX_测试用例_单元.xls'
    excelsheet='测试案例（功能）'
    excelDF = pd.DataFrame(pd.read_excel(excelfile,excelsheet))
    ##构造方法
    def __init__(self):
        print('处理文件：')
        
    def readExcFun(self):
        return StatisticsUnitTest.excelDF
    ##获取文件头
    def getFileHead(self): 
        return StatisticsUnitTest.excelDF.columns
    ##获取总行数
    def getAllDataLen(self): 
        return StatisticsUnitTest.excelDF.__len__()
    def getAllData(self): 
        return StatisticsUnitTest.excelDF.values
    ##获取日期列
    def getTestDate(self):
        return StatisticsUnitTest.excelDF[['执行日期','执行人']]
    ##根据月份统计测试用例数
    def StatisticsByMonth(self): 
        datef_df = StatisticsUnitTest.getTestDate(self) 
        print(datef_df.groupby('执行月份').count())  
        return datef_df 
    ##根据人员统计用例总数
    def StatisticsByUser(self):
        datef_df = StatisticsUnitTest.getTestDate(self) 
        print(format(datef_df.groupby(['执行人']).count()))
    ##根据人员统计每个人每个月提交的测试用例数
    def StatisticsByMonthAndUser(self):
        datef_df = StatisticsUnitTest.getTestDate(self) 
        datef_f = datef_df['执行日期'].apply(lambda x: datetime.strftime(x,"%Y-%m"))
        datef_df.insert(loc=0,column='执行月份',value=datef_f,allow_duplicates=True) 
        print(format(datef_df.groupby(['执行月份','执行人']).count()))
        print('----------------------')
        print(format(datef_df.groupby(['执行人','执行月份']).count()))

sut = StatisticsUnitTest()
sut.StatisticsByMonthAndUser()

