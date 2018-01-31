#-----------------------------Script for sales forcasting------------------------------------------------------------------
from pandas import Series, DataFrame
import pandas as pd
import numpy as np
import openpyxl as pyxl
import xlrd
import matplotlib.pyplot as plt

#---------------------To display Balance Sheet and Income statements in table format---------

wb = pyxl.load_workbook('D:/Python Work/python for finannce project/dataset.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
table = np.array([[cell.value for cell in col] for col in sheet['A1':'F9']])
print (table)                             

sales=table[1]
salesgr=sales[1:len(sales)+1]
salesgr=salesgr.astype(float)
print ("salesgr =", salesgr)

#-----------------------Method to calculate Growth rate of sales-------------------------

growthrate=[]
def growth():    
    for i in range(1,len(salesgr)):
        grate=(((salesgr[i]-salesgr[i-1])/salesgr[i-1])*100.00)
        growthrate.append(grate)
        
growth()
print ("growthrate =", growthrate)

averagegr = np.mean(growthrate)
print ("averagegr =", averagegr)

print (sales[len(sales)-1])

#-----------------------Method to calculate Forecasted sales-----------------------------

forecastsales=[]
def forecastingsales():
    for y in range (1,4):
        fsales=((1+(averagegr/100))**y)*sales[len(sales)-1]
        forecastsales.append(fsales)
forecastingsales()
print ("forecastsales =", forecastsales)


nrows=len(table)
nrows=nrows-1
ncols=len(table[0])


#-----------------------Method to calculate Percent to sales-----------------------------

percenttosales=[]
def percentsales():
    for s in range(0,nrows):
        persales=((table[s+1,5])/table[1,5])
        percenttosales.append(persales)

percentsales()
print ("percenttosales =", percenttosales)

#-----------------------Method to calculate Forecasted numbers-----------------------------

t2=[]
def futurenumbers():
    for x in range (0,3):
        for g in range(0,nrows):
            temp2=forecastsales[x]*percenttosales[g]
            t2.append(temp2)
            
futurenumbers()
print ("Forecasted numbers are ",t2)

#-----------------------External Fund Requirement-----------------------------


table1 = np.array([[cell.value for cell in col] for col in sheet['H10':'I12']])
as1=float(table1[0,1])
lb1=float(table1[1,1])
ir1=float(table1[2,1])
is0=float(sales[len(sales)-1])
is1=float(forecastsales[0]-sales[len(sales)-1])
efr = ((as1/is0)*(is1))-((lb1/is0)*(is1))-ir1
print ("External Fund requirement is",efr)


#-----------------------To plot Forecasted Sales--------------------------------------------
plt.plot(forecastsales)
plt.show()



