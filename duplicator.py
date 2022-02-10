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

#this is to delete any old sheets
existenceBool = 0
frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)
if existenceBool:
    os.remove('DummySheet2.xlsx')

#checking existence
frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

#making a fresh sheet
original = r'DummySheet.xlsx'
target = r'DummySheet2.xlsx'
shutil.copyfile(original, target)

#reading both into dataframes
inDF = pd.read_excel('DummySheet.xlsx', sheet_name='Sheet1')
outDF = pd.read_excel('DummySheet2.xlsx', sheet_name='Sheet1')

#checking existence
frameinfo = getframeinfo(currentframe())
existenceBool = file_exists(frameinfo.lineno)

#printing column headers
print("Column headings:")
print(inDF.columns)
dftry = []

#duplicating the rows, x is a var just to tell how many rows are duped
#first part for rows with dashed gauge and nam numbers
#second part for rows with nam XOR gauge
x = 0
for i in outDF.index:
    if  ((("/" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("(" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("#" in (str(outDF['GAUGE NUMBER'] [i]))))
        and 
        (("/" in (str(outDF['NAM NUMBER '] [i])))
        or ("(" in (str(outDF['NAM NUMBER '] [i])))
        or ("#" in (str(outDF['NAM NUMBER '] [i]))))
        ): #this if actually works
        #print(outDF['GAUGE NUMBER'] [i])
        outDF.loc[len(outDF.index)] = outDF.loc[i]
        x = x + 1
    elif((("/" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("(" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("#" in (str(outDF['GAUGE NUMBER'] [i]))))
        or 
        (("/" in (str(outDF['NAM NUMBER '] [i])))
        or ("(" in (str(outDF['NAM NUMBER '] [i])))
        or ("#" in (str(outDF['NAM NUMBER '] [i]))))
        ):
        outDF.loc[len(outDF.index)] = outDF.loc[i]
        x = x + 1
print('x is ' + str(x))

#sorting by gauge number
outDF['GAUGE NUMBER'] = outDF['GAUGE NUMBER'].astype(str)
outDF.sort_values(by=['GAUGE NUMBER'], ascending=True, inplace=True)


#now need to edit the numbers

#i tried to implement the functionality of this next line for 2 hours ...
#it literally just resets the indexes ... 
outDF.reset_index(drop=True, inplace=True)

#editing the values of gauge and nam numbers
truth = 0
for i in outDF.index:
    #think here: if 2 gauge numbers, if 2 nam numbers, if 2 of both, outliers ... 
    if truth == 0:
        if ("/" in (str(outDF['GAUGE NUMBER'] [i]))):
            temp = outDF['GAUGE NUMBER'] [i]
            temp1 = temp.split('/')[0]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            #print(outDF['GAUGE NUMBER'][i])
            #print(i)
            truth = 1
        if ('#' in (str(outDF['GAUGE NUMBER'][i]))):
            temp = outDF['GAUGE NUMBER'][i]
            temp1=temp.split('#')[0]
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            truth = 1
        if ('(' in (str(outDF['GAUGE NUMBER'][i]))):
            temp = outDF['GAUGE NUMBER'][i]
            temp1=temp.split('(')[0]
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            truth = 1
    else:
        if ("/" in (str(outDF['GAUGE NUMBER'] [i]))):
            #print('in elif')
            temp = outDF['GAUGE NUMBER'] [i]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'GAUGE NUMBER'] = temp2
            truth = 0
        if ("#" in (str(outDF['GAUGE NUMBER'] [i]))):
            #print('in elif')
            temp = outDF['GAUGE NUMBER'] [i]
            temp2 = temp.split('#')[0]
            outDF.loc[i,'GAUGE NUMBER'] = temp2
            truth = 0
        if ('(' in (str(outDF['GAUGE NUMBER'][i]))):
            temp = outDF['GAUGE NUMBER'][i]
            temp1=temp.split('(')[0]
            temp2=temp.split('(')[1]
            temp2=temp2[:-1]
            deletelength=len(temp2)
            temp1=temp1[-deletelength]
            temp1 = temp1 + temp2
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            truth = 0



#sorting by gauge number
# outDF['GAUGE NUMBER'] = outDF['GAUGE NUMBER'].astype(str)
# outDF.sort_values(by=['GAUGE NUMBER'], ascending=True, inplace=True)
#for some reason commenting the above 2 lines made things work right ... not sure...
outDF.reset_index(drop=True, inplace=True)

#have to push the dataframe to the excel sheet for results to show
writer = pd.ExcelWriter('DummySheet2.xlsx')
outDF.to_excel(writer)
writer.save()
