import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl as xl
import os
from os.path import exists
import shutil

#check if file exists
file_exists = exists('DummySheet2.xlsx')
if file_exists:
    print('file exists on line 12\n')

if file_exists:
    os.remove('DummySheet2.xlsx')

#check if file exists
file_exists = exists('DummySheet2.xlsx')
if file_exists:
    print('file exists on line 20\n')
else:
    print('file does not exist on line 20\n')

original = r'DummySheet.xlsx'
target = r'DummySheet2.xlsx'
shutil.copyfile(original, target)
#making a fresh sheet each time it runs

"""
#just clean duplicating the excel files
path1 = "DummySheet.xlsx"
path2 = "DummySheet2.xlsx"
wb1 = xl.load_workbook(filename=path1)
ws1 = wb1.worksheets[0]
wb2 = xl.load_workbook(filename=path2)
ws2 = wb2.create_sheet(ws1.title)
for row in ws1:
    for cell in row:
        ws2[cell.coordinate].value = cell.value
wb2.save(path2)
#end duplicating code
"""

inputSheet = pd.read_excel('DummySheet.xlsx', sheet_name='Sheet1')
outputSheet = pd.read_excel('DummySheet2.xlsx', sheet_name='Sheet1')

#check if file exists
file_exists = exists('DummySheet2.xlsx')
if file_exists:
    print('file exists on line 48\n')

print("Column headings:")
print(inputSheet.columns)
for i in inputSheet.index:
    if (len(str(inputSheet['GAUGE NUMBER'] [i])) > 5):
        #print('need to do something here')
        print(inputSheet['GAUGE NUMBER'] [i])
        #at this point, I need to duplicate the row, make the entry in the first slot
        #the first #, and the entry in the second slot the second number

#check if file exists
file_exists = exists('DummySheet2.xlsx')
if file_exists:
    print('file exists on line 62\n')

