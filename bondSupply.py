# -*- coding: utf-8 -*-
"""
Created on Thu Sep 22 16:07:35 2016

@author: yiran.zhou
"""


import pandas as pd
import numpy as np
import sys
sys.path.append('..')
import taifook.taifook as tf

# 逐行处理函数
def processLine(strDt, strEvt, idxRow):
    
#    if idxRow == 2144:
#        a = 0
    
    # 带‘Sell’字符的是拍卖，有些写Sell，有些写Sells
    if strEvt.find('Sell') == -1:
        return 0

    # 处理日期
    dt = pd.Timestamp(strDt)
    yr = dt.year
    
    # 处理拍卖量,找关键字‘USD’、‘GBP’、‘JPY’、‘JY’、‘EUR’、‘EU’后面的数字
    strAmt = ''
    cnyList = ['USD', 'GBP', 'JPY', 'JY', 'EUR', 'EU']
    for cny in cnyList:
        idx = strEvt.find(cny)
        if idx != -1: #找到了货币代码
            i = idx + len(cny)
            while (strEvt[i] != ' ' and strEvt[i] != 'B' and strEvt[i] != 'T' and strEvt[i] != 'M'):
                strAmt += strEvt[i]
                i += 1
            break
    if strAmt == '':
        print('line' + str(idxRow) + ' Didn\'t find amount')
        return 0
    try: # str转换成数字   
        amt = float(strAmt)
    except Exception as e:
        print('line' + str(idxRow) + ' amount ' + str(e))
        return 0    
    # 找货币单位， Mln、Bln、Tln,默认是Bln
    amt = amt / 1000 if strEvt.find('Mln') != -1 else amt
    amt = amt / 1000 if strEvt.find('Million') != -1 else amt
    amt = amt * 1000 if strEvt.find('Tln') != -1 else amt
    amt = amt * 1000 if strEvt.find('Trillion') != -1 else amt   
    
    # 找duration
    strDur = ''        
    if strEvt.find('-Yr') != -1: # 10-Yr类似 格式
        idx = strEvt.find('-Yr')
        strDur = strEvt[:idx].split(' ')[-1]
    elif strEvt.find('-Year') != -1: # 10-Year类似格式
        idx = strEvt.find('-Year')
        strDur = strEvt[:idx].split(' ')[-1]
    elif strEvt.find(' 20') != -1: #如果没有-Year之类格式，直接找20XX年, 年前有空格
        idx = strEvt.find(' 20')
        if idx + 5 == len(strEvt): #以年结尾
            strDur = strEvt[idx + 1:idx + 5]
        elif idx + 5 < len(strEvt): #不是以年结尾，年后必有空格
            if strEvt[idx + 5] == ' ':
                strDur = strEvt[idx + 1:idx + 5]
    if strDur == '': # 可能是1年以下的bill，不需要
        return 0
    try: # str转换成数字
        dur = float(strDur)
    except Exception as e:
        print('line' + str(idxRow) + ' duration ' + str(e))
        return 0
    if (dur > 50 and dur < 2000):
        print('line' + str(idxRow) + ' ' + str(dur), ' duration may be not right')
        return 0 # 可能是1年以下的bill，不需要
    elif (dur > 2000 and dur < 2070):
        dur = dur - yr
    elif dur > 2070:
        print('line' + str(idxRow) + ' ' + str(dur), ' duration may be not right')
        
    # 处理完成
    res = [dt, amt, dur]
    return res
        
        
def weekAddup(dfS, dfT):
    idxS = 0
    idxT = 0
    while idxT < len(dfT.index) - 1:        
        if idxS < len(dfS.index):
            while dfS.index[idxS] < dfT.index[idxT + 1]: #下周一之前
                if dfS.index[idxS] >= dfT.index[idxT]: #这周一之后
                    dfT.ix[idxT, '10YR EQ'] += dfS.ix[idxS, '10YR EQ']
                idxS += 1 
                if idxS == len(dfS.index):
                    break
            idxT += 1 #这周处理完，移到下一周       
        else: #source所有行处理完毕
            break
        
        

if __name__ == '__main__':
    
    # 读入数据
    US_raw = pd.read_excel('Bond Supply.xls', sheetname = 'US')
    JP_raw = pd.read_excel('Bond Supply.xls', sheetname = 'JP')
    UK_raw = pd.read_excel('Bond Supply.xls', sheetname = 'UK')
    GE_raw = pd.read_excel('Bond Supply.xls', sheetname = 'GE')
    rawList = [US_raw, JP_raw, UK_raw, GE_raw]
    dtCol = 'Date Time' # excel里面的列名称
    evCol = 'Event' 
    
    # 设定列和处理完的df
    cols = ['DateTime', 'Amount', 'Duration']
    US = pd.DataFrame(np.random.randn(1,len(cols)), columns = cols) #空的df不能用ix赋值，好像是bug
    JP = pd.DataFrame(np.random.randn(1,len(cols)), columns = cols)
    UK = pd.DataFrame(np.random.randn(1,len(cols)), columns = cols)
    GE = pd.DataFrame(np.random.randn(1,len(cols)), columns = cols)
    procList = [US, JP, UK, GE]    
    
    # 对每个国家逐行处理
    k = 0
#    while k < 4:
    while k < len(rawList):
        i = 0
        j = 0
        while i < len(rawList[k]):
            res = processLine(rawList[k].ix[i, dtCol], rawList[k].ix[i, evCol], i)
            if res == 0:
                i += 1
            else:
                procList[k].ix[j, :] = res
                i += 1
                j += 1
        k += 1
                   
    # 计算10yr equivalent amount
    k = 0
    while k < len(rawList):
        procList[k]['10YR EQ'] = procList[k]['Duration'] / 10 * procList[k]['Amount']
        procList[k].set_index('DateTime', inplace = True)
        k += 1
            
    # 数据按周加总
    dr = pd.date_range('2005-9-1', '2016-10-30', freq = 'W-MON')
    USweek = pd.DataFrame(index = dr, columns = ['10YR EQ'])
    JPweek = pd.DataFrame(index = dr, columns = ['10YR EQ'])
    UKweek = pd.DataFrame(index = dr, columns = ['10YR EQ'])
    GEweek = pd.DataFrame(index = dr, columns = ['10YR EQ'])
    weekList = [USweek, JPweek, UKweek, GEweek]
    k = 0
    while k < len(weekList):    
        weekList[k] = weekList[k].fillna(0)
        weekAddup(procList[k], weekList[k])
        weekList[k].index = weekList[k].index.shift(4, 'D')
        k += 1
        
        
    writer = pd.ExcelWriter('output.xlsx')
    weekList[0].to_excel(writer, 'USW')
    weekList[1].to_excel(writer, 'JPW')
    weekList[2].to_excel(writer, 'UKW')
    weekList[3].to_excel(writer, 'GEW')
    procList[0].to_excel(writer, 'US')
    procList[1].to_excel(writer, 'JP')
    procList[2].to_excel(writer, 'UK')
    procList[3].to_excel(writer, 'GE')
    writer.save()
        
        