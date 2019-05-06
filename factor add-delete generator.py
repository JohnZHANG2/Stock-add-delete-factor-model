# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 16:45:32 2019

@author: asus
"""
from pylab import *
mpl.rcParams['font.sans-serif'] = ['SimHei']
import pandas as pd
import os
from pandas_datareader import data as web
from datetime import date, timedelta
import datetime
import matplotlib.pyplot as plt
import numpy as np
start = datetime.datetime(2016,1,1)
end = datetime.datetime(2019,3,12)
window = 21
os.chdir('E:\Desktop\pe market cap\change lists')
data_ss2hk = pd.read_excel('Change_of_SSE_Securities_Lists_c.xls')
data_sz2hk = pd.read_excel('Change_of_SZSE_Securities_Lists_c.xls')
data_hscei = pd.read_excel('hist_hscei.xlsx')
data_hsi = pd.read_excel('hist_hsi.xlsx')
data_hk2ss = pd.read_excel('hk2ss-list.xlsx')
data_hk2sz = pd.read_excel('hk2sz-list.xlsx')
datecn = pd.read_excel('sh2hk chg stock cap.xlsx')['Date']
datehk = pd.read_excel('hk2ss chg stock pe.xlsx')['Date']
data_ss2hk = pd.concat([data_ss2hk['生效日期'],data_ss2hk['上交所股份編號']],axis=1)

results = pd.DataFrame()
lostss = []
lostsz = []
losthsi = []
losthscei = []
rtnss = []
namess = []
namehsi = []
namehscei = []
namesz = []
rtnhsi = []
rtnhscei = []
rtnsz = []
rhsiss = []
rhsisz = []
rtnhsiei = []
r = []
co1 = []
nameco1 = []
co2 = []
nameco2 = []
co3 = []
nameco3 = []
co4 = []
nameco4 = []
#%%
results = pd.DataFrame()
hitopx = pd.read_excel('hk2ss chg stock close.xlsx')
for i in data_hk2ss.index:
    tickers = str(data_hk2ss.loc[i,'港股代码'])
    tickers = str(tickers)
    if len(tickers)==2:
        tickers = '00'+tickers
    if len(str(tickers))==3:
        tickers = '0'+tickers
    if len(str(tickers))==1:
        tickers = '000'+tickers
    pindate = data_hk2ss.loc[i,'生效日期']
    print(tickers)
    try:
        sta = hitopx.loc[(hitopx['Date']>(pindate-timedelta(window)))&(hitopx['Date']<(pindate+timedelta(window)))][tickers+'.HK']
        sdate = hitopx.loc[(hitopx['Date']>(pindate-timedelta(window)))&(hitopx['Date']<(pindate+timedelta(window)))]['Date']
        if len(sta)>0:
            s = pd.concat([pd.DataFrame(sta),pd.DataFrame(sdate)],axis=1)
            s.set_index('Date',inplace=True)
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2ss.loc[i,'调整内容']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2ss'+tickers+'.png')
            plt.show()
            
        else:
            webdata = web.DataReader(tickers+'.HK','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            s = webdata['Close']

            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2ss.loc[i,'调整内容']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2ss'+tickers+'.png')
            plt.show()
        namess.append(tickers+'HK'+data_hk2ss.loc[i,'调整内容'])
        rtnss.append(float(s.iloc[-1]/s.iloc[1]-1))
        rhsiss.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
        co1.append(np.corrcoef(s,hsi)[0,1])
        nameco1.append(tickers+'HK'+data_hk2ss.loc[i,'调整内容'])
    except:
        try:
            webdata = web.DataReader(tickers+'.HK','yahoo',pindate-timedelta(window),pindate+timedelta(window))
            
            s = webdata['Close']
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2ss.loc[i,'调整内容']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2ss'+tickers+'.png')
            plt.show()
            namess.append(tickers+'HK'+data_hk2ss.loc[i,'调整内容'])
            rtnss.append(float(s.iloc[-1]/s.iloc[1]-1))
            rhsiss.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
            co1.append(np.corrcoef(s,hsi)[0,1])
            nameco1.append(tickers+'HK'+data_hk2ss.loc[i,'调整内容'])
        except:
            
            print(tickers+'failed on'+ str(pindate))
            lostss.append(tickers+'.HK')
            pass
    hkdatelist = list(datehk)
    for d in range(0,len(hkdatelist)):
        
        if hkdatelist[d] < pindate:
            hkdatelist[d] = 0
        else:
            
            hkdatelist[d] = 1
        
    results = pd.concat([results,pd.DataFrame(hkdatelist,columns = [tickers+'.HK'])],axis=1)
HKSS_RTN = pd.concat([pd.DataFrame(namess),pd.DataFrame(rtnss)],axis=1)
HKSS_RTN.to_excel('HK-SS成分调整前后收益.xlsx') 
HKSS_COR = pd.concat([pd.DataFrame(nameco1),pd.DataFrame(co1)],axis=1)
HKSS_COR.to_excel('HK-SS corr.xlsx') 


#%%
results = pd.DataFrame()
hitopx = pd.read_excel('hk2sz chg stock close.xlsx')
for i in data_hk2sz.index:
    tickers = str(data_hk2sz.loc[i,'港股代码'])
    tickers = str(tickers)
    if len(tickers)==2:
        
        tickers = '00'+tickers
    if len(str(tickers))==3:
        tickers = '0'+tickers
    if len(str(tickers))==1:
        tickers = '000'+tickers
    pindate = data_hk2sz.loc[i,'Date']
    print(tickers)
    try:
        sta = hitopx.loc[(hitopx['Date']>(pindate-timedelta(window)))&(hitopx['Date']<(pindate+timedelta(window)))][tickers+'.HK']
        sdate = hitopx.loc[(hitopx['Date']>(pindate-timedelta(window)))&(hitopx['Date']<(pindate+timedelta(window)))]['Date']
    
        if len(sta)>0:
            s = pd.concat([pd.DataFrame(sta),pd.DataFrame(sdate)],axis=1)
            s.set_index('Date',inplace=True)
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2sz.loc[i,'STATUS']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2sz'+tickers+'.png')
            plt.show()
        else:
            webdata = web.DataReader(tickers+'.HK','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            s = webdata['Close']
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2sz.loc[i,'STATUS']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2sz'+tickers+'.png')
            plt.show()
        namesz.append(tickers+'HK'+data_hk2sz.loc[i,'STATUS'])
        rtnsz.append(float(s.iloc[-1]/s.iloc[1]-1))
        rhsisz.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
        co2.append(np.corrcoef(s,hsi)[0,1])
        nameco2.append(tickers+'HK'+data_hk2sz.loc[i,'STATUS'])
    except:
         try:
            webdata = web.DataReader(tickers+'.HK','yahoo',pindate-timedelta(window),pindate+timedelta(window))
            s.index = pd.to_datetime(s.index)
            s = webdata['Close']
            hsi = web.DataReader('^HSI','yahoo',pindate-timedelta(window),pindate+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hk2sz.loc[i,'STATUS']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hk2sz'+tickers+'.png')
            plt.show()
            namesz.append(tickers+'HK'+data_hk2sz.loc[i,'STATUS'])
            rtnsz.append(float(s.iloc[-1]/s.iloc[1]-1))
            rhsisz.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
            co2.append(np.corrcoef(s,hsi)[0,1])
            nameco2.append(tickers+'HK'+data_hk2sz.loc[i,'STATUS'])
         except:
            print(tickers+'failed on'+ str(pindate))
            lostsz.append(tickers+'.HK')
            pass
    hkdatelist = list(datehk)
    for d in range(0,len(hkdatelist)):
        
        if hkdatelist[d] < pindate:
            hkdatelist[d] = 0
        else:
            
            hkdatelist[d] = 1
        
    results = pd.concat([results,pd.DataFrame(hkdatelist,columns = [tickers+'.HK'])],axis=1)
HKSZ_RTN = pd.concat([pd.DataFrame(namesz),pd.DataFrame(rtnsz)],axis=1)
HKSZ_RTN.to_excel('HK-SZ成分调整前后收益.xlsx')
HKSZ_COR = pd.concat([pd.DataFrame(nameco2),pd.DataFrame(co2)],axis=1)
HKSZ_COR.to_excel('HK-SZcorr with hsi.xlsx')
#%%
results = pd.DataFrame()
hitopx = pd.read_excel('HSCEI chg stock close.xlsx')
for i in data_hscei.index:
    tickers = str(data_hscei.loc[i,'StockCode'])
    tickers = str(tickers)
    if len(tickers)==2:
        tickers = '00'+tickers
    if len(str(tickers))==3:
        tickers = '0'+tickers
    if len(str(tickers))==1:
        tickers = '000'+tickers
    pindate = data_hscei.loc[i,'Effective Date 生效日期']
    print(tickers)
    try:
        sta = hitopx.loc[(hitopx['Date']>(datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window)))&(hitopx['Date']<(datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window)))][tickers+'.HK']
        sdate = hitopx.loc[(hitopx['Date']>(datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window)))&(hitopx['Date']<(datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window)))]['Date']

        if len(sta)>0:
            s = pd.concat([pd.DataFrame(sta),pd.DataFrame(sdate)],axis=1)
            s.set_index('Date',inplace=True)
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hscei.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hscei'+tickers+'.png')
            plt.show()
        else:
            
            webdata = web.DataReader(tickers+'.HK','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))
            s = webdata['Close']
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hscei.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hscei'+tickers+'.png')
            plt.show()
        namehscei.append(tickers+'HK'+data_hscei.loc[i,'Change 變動'])
        rtnhscei.append(float(s.iloc[-1]/s.iloc[1]-1))      
        rtnhsiei.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
        co3.append(np.corrcoef(s,hsi)[0,1])
        nameco3.append(tickers+'HK'+data_hscei.loc[i,'Change 變動'])
    except:
        try:
            webdata = web.DataReader(tickers+'.HK','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))
            
            s = webdata['Close']
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hscei.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hscei'+tickers+'.png')
            plt.show()
            namehscei.append(tickers+'HK'+data_hscei.loc[i,'Change 變動'])
            rtnhscei.append(float(s.iloc[-1]/s.iloc[1]-1))
            rtnhsiei.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
            co3.append(np.corrcoef(s,hsi)[0,1])
            nameco3.append(tickers+'HK'+data_hscei.loc[i,'Change 變動'])
        except:
            print(tickers+'failed on'+ str(pindate))
            
            losthscei.append(tickers+'.HK')
            pass
    hkdatelist = list(datehk)
    for d in range(0,len(hkdatelist)):
        
        if hkdatelist[d] < datetime.datetime.strptime(pindate,"%Y-%m-%d"):
            hkdatelist[d] = 0
        else:
            
            hkdatelist[d] = 1
        
    results = pd.concat([results,pd.DataFrame(hkdatelist,columns = [tickers+'.HK'])],axis=1)
HSCEI_RTN = pd.concat([pd.DataFrame(namehscei),pd.DataFrame(rtnhscei)],axis=1)
HSCEI_RTN.to_excel('HSCEI成分调整前后收益.xlsx')
HSCEI_COR = pd.concat([pd.DataFrame(nameco3),pd.DataFrame(co3)],axis=1)
HSCEI_COR.to_excel('hsceicorr with hsi.xlsx')
#%%
results = pd.DataFrame()
hitopx = pd.read_excel('HSI CHANGE stocks close.xlsx')
for i in data_hsi.index:
    tickers = str(data_hsi.loc[i,'StockCode'])
    tickers = str(tickers)
    if len(tickers)==2:
        tickers = '00'+tickers
    if len(str(tickers))==3:
        tickers = '0'+tickers
    if len(str(tickers))==1:
        tickers = '000'+tickers
    pindate = data_hsi.loc[i,'Effective Date 生效日期']
    
    print(tickers)
    try:
        sta = hitopx.loc[(hitopx['Date']>(datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window)))&(hitopx['Date']<(datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window)))][tickers+'.HK']
        sdate = hitopx.loc[(hitopx['Date']>(datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window)))&(hitopx['Date']<(datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window)))]['Date']
        if len(sta)>0:
            s = pd.concat([pd.DataFrame(sta),pd.DataFrame(sdate)],axis=1)
            s.set_index('Date',inplace=True)
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hsi.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hsi'+tickers+'.png')
            plt.show()
            
        else:
            
            webdata = web.DataReader(tickers+'.HK','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))
        
            s = webdata['Close']
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hsi.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hsi'+tickers+'.png')
            plt.show()
                
        namehsi.append(tickers+'.HK'+data_hsi.loc[i,'Change 變動'])
        rtnhsi.append(float(s.iloc[-1]/s.iloc[1]-1))
        r.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
        co4.append(np.corrcoef(s,hsi)[0,1])
        nameco4.append(tickers+'HK'+data_hsi.loc[i,'Change 變動'])
    except:
          try:
            webdata = web.DataReader(tickers+'.HK','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))
            
            s = webdata['Close']
            s.index = pd.to_datetime(s.index)
            hsi = web.DataReader('^HSI','yahoo',datetime.datetime.strptime(pindate,"%Y-%m-%d")-timedelta(window),datetime.datetime.strptime(pindate,"%Y-%m-%d")+timedelta(window))            
            hsi = hsi['Close']
           
            
            fig, ax1 = plt.subplots(figsize=(10,10))            
            color = 'tab:red'
            ax1.set_xlabel('date (s)')
            ax1.set_ylabel('stock price', color=color)
            ax1.plot(s, color=color)
            ax1.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            ax1.set_title(tickers+'.HK'+data_hsi.loc[i,'Change 變動']+str(pindate))
            ax2 = ax1.twinx()
            color = 'tab:blue'
            ax2.set_ylabel('hsi', color=color)  # we already handled the x-label with ax1
            ax2.plot(hsi, color=color)
            ax2.tick_params(axis='y', labelcolor=color,labelsize=15,labelrotation=45)
            fig.tight_layout()
            plt.savefig('hsi'+tickers+'.png')
            plt.show()
            namehsi.append(tickers+'HK'+data_hsi.loc[i,'Change 變動'])
            rtnhsi.append(float(s.iloc[-1]/s.iloc[1]-1))
            r.append(float(hsi.iloc[-1]/hsi.iloc[1]-1))
            co4.append(np.corrcoef(s,hsi)[0,1])
            nameco4.append(tickers+'HK'+data_hsi.loc[i,'Change 變動'])
          except:
            print(tickers+'failed on'+ str(pindate))
            
            losthsi.append(tickers+'.HK')
            pass
    hkdatelist = list(datehk)
    for d in range(0,len(hkdatelist)):
        if hkdatelist[d] < datetime.datetime.strptime(str(pindate),"%Y-%m-%d"):
            hkdatelist[d] = 0
        else:
            
            hkdatelist[d] = 1
        
    results = pd.concat([results,pd.DataFrame(hkdatelist,columns = [tickers+'.HK'])],axis=1)
HSI_RTN = pd.concat([pd.DataFrame(namehsi),pd.DataFrame(rtnhsi)],axis=1)
HSI_RTN.to_excel('HSI成分调整前后收益.xlsx')
HSI_COR = pd.concat([pd.DataFrame(nameco4),pd.DataFrame(co4)],axis=1)
HSI_COR.to_excel('HSIcorr hsi.xlsx')

