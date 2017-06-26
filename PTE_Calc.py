import xlwings as xw
import pandas as pd
import numpy as np
import os
from tkinter.filedialog import askdirectory
#select test time folder
#dir = askdirectory()


app=xw.App(visible=False)

#acquire time data
def time_data(path):
 wb = xw.books.open(path)

#step_name = np.delete(xw.Range('A2:C26').value,1,1) #delete second column
#array = np.array(step_name)
 sht = wb.sheets['Step']

 rowmax = sht.api.UsedRange.Rows.count

 step_name=xw.Range((2,1),(rowmax,1)).value

 step_time=xw.Range((2,3),(rowmax,3)).value
 d = {'step_name':pd.Series(step_name),'step_time':pd.Series(step_time),}

#sort data by ascending
 ds = pd.DataFrame(d)
 a = ds.sort_values(by='step_name', ascending=True)
 wb.close() #close book so no error  repeat call when visibel is false

 return a

def mean_time(folderpath, sitenumber):
 paths = os.listdir(folderpath)
 for index in range(len(paths)):
  path = folderpath + '\\' + paths[index]
  if index == 0:
   step_time = time_data(path)
  else:
   step_time = pd.merge(time_data(path),step_time, on="step_name")

#variable step_time include sorted original step time data
#calculate the mean and append to original dataframe

 mean = pd.Series()
 for x in range(len(step_time.index)):
     row = step_time.iloc[x:x+1, 1:]
     avg = row.sum(1) / float(len(row.columns))
     mean = mean.append(avg, ignore_index= True)
 step_time.insert( len(step_time.columns), sitenumber,mean)

 del step_time ['step_time_x']
 del step_time ['step_time_y']

 return step_time

#dir = os.path.abspath('.')+'\Test Time'
dir = r'C:\Users\luo\Desktop\HS6601 Correlation Report\Test Time'
site_1 = mean_time(dir+'\\1 site','1 Site')
site_2 = mean_time(dir+'\\2 site','2 Sites')
site_4 = mean_time(dir+'\\4 site','4 Sites')

raw_result = pd.merge(site_1,site_2, on = 'step_name')
raw_result = pd.merge(raw_result,site_4, on = 'step_name')
#raw_result = raw_result.drop('step_time', axis = 1) #may contain 'step_time/_x/_y'
print(raw_result)

#add columns of pte calculation result
pte_2 = pd.Series()
pte_4 = pd.Series()
for step in range(len(raw_result.index)):
 row = raw_result.iloc[step:step+1,1:]
 E2 = row['1 Site']
 F2 = row['2 Sites']
 H2 = row['4 Sites']
 pte2 = (1-(((F2-E2)/1)/E2)) * 100
 pte4 = ((1-(((H2-E2)/3)/E2))) * 100
 pte_2 = pte_2.append(pte2, ignore_index= True)
 pte_4 = pte_4.append(pte4, ignore_index= True)

raw_result.insert(3, '2 Sites PTE(%)', pte_2)
raw_result.insert(5, '4 Sites PTE(%)', pte_4)

#add book and sheet and print
#app=xw.App(visible=True,add_book=False)
#wb = app.books.add()
#wb.sheets.add(name= "Sorted Time Data")
#wb.sheets[0].range('A1').value = raw_result
#wb.save(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Test Time')

#wb=xw.Book(os.path.abspath('.')+'\Correlation Calc.xlsm')
wb=xw.Book(r'C:\Users\luo\Desktop\HS6601 Correlation Report\Correlation Calc.xlsm')
sht=wb.sheets['PTE_Data']
sht.range('A1').value=raw_result
