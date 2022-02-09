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
#this is to delete any old sheets

frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

original = r'DummySheet.xlsx'
target = r'DummySheet2.xlsx'
shutil.copyfile(original, target)
#making a fresh sheet

inDF = pd.read_excel('DummySheet.xlsx', sheet_name='Sheet1')
outDF = pd.read_excel('DummySheet2.xlsx', sheet_name='Sheet1')

frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

print("Column headings:")
print(inDF.columns)
dftry = []

#this is now just to print out the gauge numbers to be modified
#for i in inDF.index:
#    if (len(str(inDF['GAUGE NUMBER'] [i])) > 5):
#        print(inDF['GAUGE NUMBER'] [i])

#duplicating the rows, x is a var just to tell how many rows are duped
x = 0
for i in outDF.index:
    if (("/" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("(" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("#" in (str(outDF['GAUGE NUMBER'] [i])))): #this if actually works
        print ('found one')
        print(outDF['GAUGE NUMBER'] [i])
        outDF.loc[len(outDF.index)] = outDF.loc[i]
        
        #i = i + 1
        x = x + 1
print('x is ' + str(x))

#trying to sort
outDF['GAUGE NUMBER'] = outDF['GAUGE NUMBER'].astype(str)
outDF.sort_values(by=['GAUGE NUMBER'], ascending=True, inplace=True)

#trying to update and see if that helps
writer = pd.ExcelWriter('DummySheet2.xlsx')
outDF.to_excel(writer)
writer.save()

#now need to edit the numbers