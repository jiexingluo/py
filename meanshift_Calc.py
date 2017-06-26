import xlwings as xw
import pandas as pd
import numpy as np
import os
import sys
from tkinter.filedialog import askdirectory

#select meanshift folder
#dir = askdirectory()

#dir=os.path.abspath(os.curdir)+'\meanshift'
dir = r'C:\Users\luo\Desktop\HS6601 Correlation Report\meanshift'

def importsheet2dataframe(bookname, sheetname):
 wb = xw.Book(dir + bookname)
 sht = wb.sheets[sheetname]


 #use VBA function to acquire the max rows and columns

 columnmax = sht.api.UsedRange.Columns.count
 rowmax = sht.api.UsedRange.Rows.count

 raw_data_all = pd.DataFrame(sht[1:rowmax,23:columnmax].value, index =sht[1:rowmax,22].value, columns=sht[0,23:columnmax].value)


 return raw_data_all

rawdata_1 = importsheet2dataframe( r'\\site0.xlsx', r'_00_STDF01') #get data start from site_num and used as index
rawdata_1 = rawdata_1.drop(['hbin_num','sbin_num','NUM_TEST','TEST_T(ms)','site_num'], axis = 1) #all data include site, time, result
site1_mean = rawdata_1.mean()

rawdata_2 = importsheet2dataframe( r'\\site1.xlsx', r'_01_STDF01') #get data start from site_num and used as index
rawdata_2 = rawdata_2.drop(['hbin_num','sbin_num','NUM_TEST','TEST_T(ms)','site_num'], axis = 1) #all data include site, time, result
site2_mean = rawdata_2.mean()

rawdata = importsheet2dataframe( r'\\site2.xlsx', r'_02_STDF01') #get data start from site_num and used as index
rawdata = rawdata.drop(['hbin_num','sbin_num','NUM_TEST','TEST_T(ms)','site_num'], axis = 1) #all data include site, time, result
site3_mean = rawdata.mean()

rawdata = importsheet2dataframe( r'\\site3.xlsx', r'_03_STDF01') #get data start from site_num and used as index
rawdata = rawdata.drop(['hbin_num','sbin_num','NUM_TEST','TEST_T(ms)','site_num'], axis = 1) #all data include site, time, result
site4_mean = rawdata.mean()

def meanshift(site1_mean,site2_mean):
 msarr = []
 for row in range(site1_mean.size):
     ms = abs((site1_mean[row] - site2_mean[row]) / site1_mean[row] * 100)

     msarr.append(ms)  #direct append and return the value
 msseries = pd.DataFrame(msarr, index=list(site1_mean.index))  # list to Series
 print(msseries)
 return msseries

meanshift_1to2 = meanshift(site1_mean, site2_mean)
meanshift_1to3 = meanshift(site1_mean, site3_mean)
meanshift_1to4 = meanshift(site1_mean, site4_mean)

sites_mean = pd.concat([site1_mean,site2_mean,site3_mean,site4_mean,meanshift_1to2,meanshift_1to3,meanshift_1to4], axis=1)
sites_mean.columns = ['Site 1 Mean','Site 2 Mean','Site 3 Mean','Site 4 Mean', 'Site 1 to Site 2 Meanshift (%)', 'Site 1 to Site 3 Meanshift (%)', 'Site 1 to Site 4 Meanshift (%)',]

#add book and sheet and print
#app=xw.App(visible=True,add_book=False)
#wb = app.books.add()
#wb.sheets.add(name= "meanshift result")
#wb.sheets[0].range('A1').value = sites_mean

#wb.save(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Quad 30 Loops')
#wb=xw.Book(os.path.abspath('.')+'\Correlation Calc.xlsm')
wb=xw.Book(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Correlation Calc.xlsm')
sht=wb.sheets['Meanshift_Data']
sht.range('A1').value=sites_mean

