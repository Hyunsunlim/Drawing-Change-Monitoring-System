#%%
#Writer: May Lim
#Project: Drawing Mangament System
#Language: Python
#Date: 2022/08/25

from cmath import isnan
from pickle import TRUE
from pandas.core.frame import DataFrame
import pymysql
import pandas as pd
import math
import numpy as np
import matplotlib.pyplot as plt
import re
import datetime
now = datetime.datetime.now()
df_all = pd.read_excel('C:/Users/user/Desktop/DB_sample2.xlsx', sheet_name = 'List', usecols = [2,3,4], names= ['Drawing_num', 'Data File', 'Version'], skiprows=[0])
df_l = pd.read_excel('C:/Users/user/Desktop/DB_sample2.xlsx', sheet_name = 'Letter')


def Reshape(data):
    data = data.dropna(axis=0)
    data = data.astype({'Version': 'int'})
    data.drop_duplicates(['Drawing_num'])
    return data

Reshape(df_all)

def casedate(data):
    cases = data.iloc[1].fillna('-').values.tolist()
    return cases

g = casedate(df_l)


#Identify Table from Letter Sheet

a = []
b = []
issue = []

Statistics = pd.DataFrame(index=range(0,4), columns = ['First'])
Statistics.index = ['Drawing_num', 'Added', 'Revised', 'Deleted']
## Table Query
all_index = []
for i in range(len(g)):
    if g[i] != '-':
        locals()['df_' + g[i]] = df_l.iloc[3:,g.index(g[i])+1:g.index(g[i])+4].reset_index(drop=True)
        locals()['df_' + g[i]].index=locals()['df_' + g[i]].index + 1
        locals()['df_' + g[i]].columns=['Letter_num','Receipt','Date']
        locals()['df_' + g[i]]['Date'] = pd.to_datetime(locals()['df_' + g[i]]['Date'])
        locals()['df_' + g[i]] = locals()['df_' + g[i]].assign(Time_Lapse = now - locals()['df_' + g[i]]['Date'])
        locals()['L_' + g[i]+'_l'] = locals()['df_' + g[i]].loc[locals()['df_' + g[i]]['Time_Lapse'].idxmin(axis=0)]
        locals()[ g[i] + '_d'] = pd.read_excel('C:/Users/user/Desktop/DB_sample2.xlsx', sheet_name = g[i])
        locals()[ g[i] + 'l_d'] = locals()[ g[i] + '_d'].iloc[5:,locals()[ g[i] + '_d'].iloc[0].values.tolist().index(locals()['df_' + g[i]]['Time_Lapse'].idxmin(axis=0))+1:locals()[ g[i] + '_d'].iloc[0].values.tolist().index(locals()['df_' + g[i]]['Time_Lapse'].idxmin(axis=0))+5].reset_index(drop=True)
        locals()[ g[i] + 'l_d'].index = locals()[ g[i] + 'l_d'].index + 1 
        locals()[ g[i] + 'l_d'].columns = ['Drawing_num','Added','Revised','Deleted']
        locals()[ g[i] + 'l_d'] = locals()[ g[i] + 'l_d'].dropna(subset=["Drawing_num"])
        
        locals()[ g[i] + '_sta'] = locals()[ g[i] + 'l_d'].count().to_frame(name = g[i])
        Statistics = pd.concat([Statistics,locals()[ g[i] + '_sta']], axis=1)
        
        all_index.append([item for item in locals()['df_' + g[i]].index[locals()['df_' + g[i]]['Receipt'] == 'SRJV']])

        a.append(locals()['L_' + g[i]+'_l'][1])
        b.append(locals()['L_' + g[i]+'_l'][3])
        issue.append(g[i])

l = pd.DataFrame({ 'Issue':issue, 'Ellapsed Time':b ,'Receipant':a })

d_list = []

#print(searching)

for i in range(len(l)):
    a = pd.read_excel('C:/Users/user/Desktop/DB_sample2.xlsx', sheet_name = l['Issue'][i]) ##a:Drawing list
    for j in range(len(g)):
      if l['Issue'][i] == g[j]:
        c =  df_l.iloc[3:,g.index(g[j])+1:g.index(g[j])+4].reset_index(drop=True)
        c.index = c.index + 1
        c.columns = ['Letter_num','Receipt','Date']
        c = c.dropna(subset=["Letter_num"])
        c2 = c.assign(Time_Lapse = now - c['Date'])
        b = a.iloc[5:,a.iloc[0].values.tolist().index(c2['Time_Lapse'].idxmin(axis=0))+1:a.iloc[0].values.tolist().index(c2['Time_Lapse'].idxmin(axis=0))+5].reset_index(drop=True)
        b.reset_index()
        b.columns = ['Drawing_num','Added','Revised','Deleted']
        b = b.dropna(subset=["Drawing_num"])
        f= b['Drawing_num'].values.tolist()
        d_list.append(f)

l = l.assign(drawing_list = d_list)
l['Quantity'] = l.apply(lambda x: len(x['drawing_list']), axis = 1)
l['State'] = l.apply(lambda x: 'Closed' if x['Receipant'] == 'SRJV' else 'Opened', axis = 1 )

al_d = df_all['Drawing_num'].values.tolist()


df_all["Process In"] = ""
df_all["Current State"] = ""
df_all["Elapsed Time"] = ""
df_all["Previous State"] = "" 

srjv_index = []
dc_index = []
gc_index = []
add=[]
for i in range(len(l)):
    for x in l['drawing_list'][i]:
        if l['Receipant'][i] == 'SRJV':
            if (df_all['Drawing_num']==x).any():
                index=df_all[df_all['Drawing_num'] == x].index[0]
                srjv_index.append(index)
                if int(df_all['Version'].loc[index]) >= 5001:
                    df_all['Version'].loc[index] = int(df_all['Version'].loc[index]) + 1
                elif int(df_all['Version'].loc[index]) < 5001:
                    df_all['Version'].loc[index] = 5001
                df_all["Previous State"].loc[index] = l['Issue'][i]
            else:
                added_list = [x, '■', '5001', '', '', '', l['Issue'][i]]
                add.append(added_list)
                df_all = df_all.append(pd.Series(added_list, index=df_all.columns), ignore_index=True)
                df_all["Previous State"].loc[index] = l['Issue'][i]
        if l['Receipant'][i] == 'DC':
            if (df_all['Drawing_num']==x).any():
                index=df_all[df_all['Drawing_num'] == x].index[0]
                dc_index.append(index) # l['Receipant'][i] 
                df_all['Process In'].loc[index] = l['Receipant'][i]
                df_all['Current State'].loc[index] = l['Issue'][i]
                df_all['Elapsed Time'].loc[index] = l['Ellapsed Time'][i]
            else:
                added_list = [x, '■', 5000,l['Receipant'][i], l['Issue'][i], l['Ellapsed Time'][i], '']
                df_all = df_all.append(pd.Series(added_list, index=df_all.columns), ignore_index=True)
        if l['Receipant'][i] == 'GC':
            if (df_all['Drawing_num']==x).any():
                index=df_all[df_all['Drawing_num'] == x].index[0]
                gc_index.append(index)    
                df_all['Process In'].loc[index] = l['Receipant'][i]
                df_all['Current State'].loc[index] = l['Issue'][i]
                df_all['Elapsed Time'].loc[index] = l['Ellapsed Time'][i]
            else:
                added_list = [x, '■', 5000,l['Receipant'][i], l['Issue'][i], l['Ellapsed Time'][i], '']
                df_all = df_all.append(pd.Series(added_list, index=df_all.columns), ignore_index=True)
        if l['Receipant'][i] == 'TIAC':
            if (df_all['Drawing_num']==x).any():
                index=df_all[df_all['Drawing_num'] == x].index[0]
                gc_index.append(index)    
                df_all['Process In'].loc[index] = l['Receipant'][i]
                df_all['Current State'].loc[index] = l['Issue'][i]
                df_all['Elapsed Time'].loc[index] = l['Ellapsed Time'][i]
            else:
                added_list = [x, '■', 5000,l['Receipant'][i], l['Issue'][i], l['Ellapsed Time'][i], '']
                df_all = df_all.append(pd.Series(added_list, index=df_all.columns), ignore_index=True)
        
            
df_all = df_all[['Drawing_num','Version','Previous State','Current State', 'Process In','Elapsed Time']]
df_all.columns = ['Drawing_num','Version','Previous State','Current State', 'Process In','Elapsed Time(day)']
df_all['Version'] = pd.to_numeric(df_all['Version'])

df_all.to_excel('result_drawinglist.xlsx')

def search(name):
    df = pd.DataFrame(columns = {'Drawing_num','Version','Current State','Process In','Elapsed Time(day)'})
    for i in name:
        bb = df_all[df_all['Drawing_num'].str.contains(i)]
        df=pd.concat([df, bb], ignore_index = True, axis = 0)
    return df


a = input('圖號: ').split()
pd.options.display.max_rows = 90
search(a)



# %%
