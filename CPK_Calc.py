import xlwings as xw
import pandas as pd
import numpy as np
import os
from tkinter.filedialog import askdirectory

#select quad 30 loops folder
#dir = askdirectory()

#dir = os.path.abspath('.')+'\Quad 30 Loops'

dir = r'C:\Users\luo\Desktop\HS6601 Correlation Report\Quad 30 Loops'
path = r'\test.xlsx'
#the excel must be converted and opened successfully without select
wb = xw.Book(dir + path)


def importsheet2dataframe(sheetname):
 sht = wb.sheets[sheetname]
 #use VBA function to acquire the max rows and columns

 columnmax = sht.api.UsedRange.Columns.count
 rowmax = sht.api.UsedRange.Rows.count

 raw_data_all = pd.DataFrame(sht[1:rowmax,23:columnmax].value, index =sht[1:rowmax,22].value, columns=sht[0,23:columnmax].value)
 return raw_data_all

def importlimit2dataframe(sheetname):
 sht = wb.sheets[sheetname]
 #use VBA function to acquire the max rows and columns

 columnmax = sht.api.UsedRange.Columns.count
 rowmax = sht.api.UsedRange.Rows.count

 raw_data_all = pd.DataFrame(sht[1:rowmax,1:columnmax].value, index =sht[1:rowmax,0].value, columns=sht[0,1:columnmax].value)
 return raw_data_all

def Cp(mylist, usl, lsl):
    arr = np.array(mylist)
    arr = arr.ravel()
    sigma = np.std(arr)
    Cp = float(usl - lsl) / (6*sigma)
    return Cp

def Cpk(mylist, usl, lsl):
    arr = np.array(mylist)
    arr = arr.ravel()
    sigma = np.std(arr)
    m = np.mean(arr)

    Cpu = float(usl - m) / (3*sigma)
    Cpl = float(m - lsl) / (3*sigma)
    Cpk = np.min([Cpu, Cpl])
    return Cpk

rawdata = importsheet2dataframe(r'_00_STDF01') #get data start from site_num and used as index
rawdata = rawdata.drop(['hbin_num','sbin_num','NUM_TEST'], axis = 1) #all data include site, time, result

rawlimit = importlimit2dataframe('_00_Test_Limits')
rawlimit = rawlimit.drop(['Units','SpecLo','SpecHi'],axis=1)


def site_cpcpk(site, num):
   site1 = site.drop(['site_num','TEST_T(ms)'], axis =1) #only calculation needed data
   cpkarr = []
   cparr = []
   #calculate the cpk and add to each site
   for c in range (len(site1.columns)):
    series = site1.ix[:, c]
    ndarray = series.values

    lsl = rawlimit.iloc[c, 1]
    usl = rawlimit.iloc[c, 2]

    cp = Cp(ndarray,usl,lsl )
    cpk = Cpk(ndarray,usl,lsl)

   #generate cpk list for all test item by for loop
    cpkarr.append(cpk)
    cparr.append(cp)
   site1_cp = pd.DataFrame(cparr, index=list(site1.columns), columns=['site'+num+'_cp']) #list to Series
   site1_cpk = pd.DataFrame(cpkarr, index=list(site1.columns), columns=['site'+num+'_cpk'])

   return (site1_cpk, site1_cp)

site1 = rawdata.query('site_num == 0')
site1_cpcpk  = site_cpcpk(site1, '1')

site2 = rawdata.query('site_num == 1')
site2_cpcpk = site_cpcpk(site2, '2')
site3 = rawdata.query('site_num == 2')
site3_cpcpk = site_cpcpk(site3, '3')
site4 = rawdata.query('site_num == 3')
site4_cpcpk = site_cpcpk(site4, '4')

sites_cpcpk = site1_cpcpk[0].join(site1_cpcpk[1])
sites_cpcpk = sites_cpcpk.join(site2_cpcpk)
sites_cpcpk = sites_cpcpk.join(site3_cpcpk)
sites_cpcpk = sites_cpcpk.join(site4_cpcpk)

print(sites_cpcpk)

#add book and sheet and print
#app=xw.App(visible=True,add_book=False)
#wb = app.books.add()
#wb.sheets.add(name= "CPK result")
#wb.sheets[0].range('A1').value = sites_cpcpk


#wb=xw.Book(os.path.abspath('.')+'\Correlation Calc.xlsm')
wb=xw.Book(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Correlation Calc.xlsm')
sht=wb.sheets['CPK_Data']
sht.range('A1').value=sites_cpcpk

#wb.save(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Quad 30 Loops')