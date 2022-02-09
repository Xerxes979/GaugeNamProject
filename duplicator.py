import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl as xl
import os
from os.path import exists
import shutil
import numpy as np
from inspect import currentframe, getframeinfo




#function to check that the file exists at different places in the code
def file_exists(line_number): 
    existence = exists('Dummysheet2.xlsx')
    if existence:
        print('file exists on line ' + str(line_number))
    else:
        print('file does not exist on line ' + str(line_number))
    return existence

existenceBool = 0
frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)
if existenceBool:
    os.remove('DummySheet2.xlsx')
#this is to make sure the output is completely new each run

frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

original = r'DummySheet.xlsx'
target = r'DummySheet2.xlsx'
shutil.copyfile(original, target)
#making a fresh sheet each time it runs

inDF = pd.read_excel('DummySheet.xlsx', sheet_name='Sheet1')
outDF = pd.read_excel('DummySheet2.xlsx', sheet_name='Sheet1')

frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

print("Column headings:")
print(inDF.columns)
dftry = []

for i in inDF.index:
    if (len(str(inDF['GAUGE NUMBER'] [i])) > 5):
        dftry.append(True)
        print(inDF['GAUGE NUMBER'] [i])
        #at this point, I need to duplicate the row, make the entry in the first row
        #the first number, and the entry in the second row the second number
    else:
        dftry.append(False)

numrows = len(inDF)

x = 0
for i in range(0,numrows):
    #print(dftry[i])
    if dftry[i]:
        outDF.loc[i-1] = np.repeat(outDF.loc[i], 1)
        #why isn't this working ughhhhhhhh
        x = x + 1
print('x is ' + str(x))
