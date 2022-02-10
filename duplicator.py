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
#uncomment the lines inside if you are debugging and the file won't show up
def file_exists(line_number): 
    existence = exists('Dummysheet2.xlsx')
    #if existence:
    #    print('file exists on line ' + str(line_number))
    #else:
    #    print('file does not exist on line ' + str(line_number))
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
        or ("&" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("#" in (str(outDF['GAUGE NUMBER'] [i]))))
        and 
        (("/" in (str(outDF['NAM NUMBER'] [i])))
        or ("(" in (str(outDF['NAM NUMBER'] [i])))
        or ("&" in (str(outDF['NAM NUMBER'] [i])))
        or ("#" in (str(outDF['NAM NUMBER'] [i]))))
        ): #this if actually works
        #print(outDF['GAUGE NUMBER'] [i])
        outDF.loc[len(outDF.index)] = outDF.loc[i]
        x = x + 1
    elif((("/" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("(" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("&" in (str(outDF['GAUGE NUMBER'] [i])))
        or ("#" in (str(outDF['GAUGE NUMBER'] [i]))))
        or 
        (("/" in (str(outDF['NAM NUMBER'] [i])))
        or ("(" in (str(outDF['NAM NUMBER'] [i])))
        or ("&" in (str(outDF['NAM NUMBER'] [i])))
        or ("#" in (str(outDF['NAM NUMBER'] [i]))))
        ):
        outDF.loc[len(outDF.index)] = outDF.loc[i]
        x = x + 1
print('number of lines duplicated is ' + str(x))

######################################################################
#
# LOOK RIGHT HERE
#
# change 'GAUGE NUMBER' to 'NAM NUMBER' below if you want 
# the resulting sheet to sort by nam number rather than gauge number
######################################################################
#sorting by gauge number
outDF['GAUGE NUMBER'] = outDF['GAUGE NUMBER'].astype(str)
outDF.sort_values(by=['GAUGE NUMBER'], ascending=True, inplace=True)


#resetting the indexes 
outDF.reset_index(drop=True, inplace=True)

#editing the values of gauge numbers
truth = 0 #value to allow alternation between original and duplicate line code
for i in outDF.index: 
    if truth == 0:
        if ("/" in (str(outDF['GAUGE NUMBER'] [i]))):
            temp = outDF['GAUGE NUMBER'] [i]
            temp1 = temp.split('/')[0]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            truth = 1
        if ('#' in (str(outDF['GAUGE NUMBER'][i]))):
            temp = outDF['GAUGE NUMBER'][i]
            temp1=temp.split('#')[0]
            outDF.loc[i,'GAUGE NUMBER'] = temp1
            truth = 1
        if ('(' in (str(outDF['GAUGE NUMBER'][i]))):
            if ('LH' in (str(outDF['GAUGE NUMBER'][i]))):
                truth = 1
            elif ('RH' in (str(outDF['GAUGE NUMBER'][i]))):
                truth = 1
            else:
                temp = outDF['GAUGE NUMBER'][i]
                temp1=temp.split('(')[0]
                outDF.loc[i,'GAUGE NUMBER'] = temp1
                truth = 1
    else:
        if ("/" in (str(outDF['GAUGE NUMBER'] [i]))):
            temp = outDF['GAUGE NUMBER'] [i]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'GAUGE NUMBER'] = temp2
            truth = 0
        if ("#" in (str(outDF['GAUGE NUMBER'] [i]))):
            temp = outDF['GAUGE NUMBER'] [i]
            temp2 = temp.split('#')[0]
            outDF.loc[i,'GAUGE NUMBER'] = temp2
            truth = 0
        if ('(' in (str(outDF['GAUGE NUMBER'][i]))):
            if ('LH' in (str(outDF['GAUGE NUMBER'][i]))):
                truth = 0
            elif ('RH' in (str(outDF['GAUGE NUMBER'][i]))):
                truth = 0
            else:
                temp = outDF['GAUGE NUMBER'][i]
                temp1=temp.split('(')[0]
                temp2=temp.split('(')[1]
                temp2=temp2[:-1]
                deletelength=len(temp2)
                temp1=temp1[-deletelength]
                temp1 = temp1 + temp2
                outDF.loc[i,'GAUGE NUMBER'] = temp1
                truth = 0

truth = 0
for i in outDF.index:
    #doing the exact same as above for nam numbers
    if truth == 0:
        if ("/" in (str(outDF['NAM NUMBER'] [i]))):
            temp = outDF['NAM NUMBER'] [i]
            temp1 = temp.split('/')[0]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'NAM NUMBER'] = temp1
            truth = 1
        if ("&" in (str(outDF['NAM NUMBER'] [i]))):
            temp = outDF['NAM NUMBER'] [i]
            temp1 = temp.split('&')[0]
            temp2 = temp.split('&')[1]
            outDF.loc[i,'NAM NUMBER'] = temp1
            truth = 1
        if ('#' in (str(outDF['NAM NUMBER'][i]))):
            temp = outDF['NAM NUMBER'][i]
            temp1=temp.split('#')[0]
            outDF.loc[i,'NAM NUMBER'] = temp1
            truth = 1
        if ('(' in (str(outDF['NAM NUMBER'][i]))):
            if ('LH' in (str(outDF['NAM NUMBER'][i]))):
                truth = 1
            elif ('RH' in (str(outDF['NAM NUMBER'][i]))):
                truth = 1
            else:
                temp = outDF['NAM NUMBER'][i]
                temp1=temp.split('(')[0]
                outDF.loc[i,'NAM NUMBER'] = temp1
                truth = 1
    else:
        if ("/" in (str(outDF['NAM NUMBER'] [i]))):
            temp = outDF['NAM NUMBER'] [i]
            temp2 = temp.split('/')[1]
            outDF.loc[i,'NAM NUMBER'] = temp2
            truth = 0
        if ("&" in (str(outDF['NAM NUMBER'] [i]))):
            temp = outDF['NAM NUMBER'] [i]
            temp2 = temp.split('&')[1]
            outDF.loc[i,'NAM NUMBER'] = temp2
            truth = 0
        if ("#" in (str(outDF['NAM NUMBER'] [i]))):
            temp = outDF['NAM NUMBER'] [i]
            temp2 = temp.split('#')[0]
            outDF.loc[i,'NAM NUMBER'] = temp2
            truth = 0
        if ('(' in (str(outDF['NAM NUMBER'][i]))):
            if ('LH' in (str(outDF['NAM NUMBER'][i]))):
                truth = 0
            elif ('RH' in (str(outDF['NAM NUMBER'][i]))):
                truth = 0
            else:
                temp = outDF['NAM NUMBER'][i]
                temp1=temp.split('(')[0]
                temp2=temp.split('(')[1]
                temp2=temp2[:-1]
                deletelength=len(temp2)
                temp1=temp1[-deletelength]
                temp1 = temp1 + temp2
                outDF.loc[i,'NAM NUMBER'] = temp1
                truth = 0

outDF.reset_index(drop=True, inplace=True)

#pushing the dataframe to the excel sheet
writer = pd.ExcelWriter('DummySheet2.xlsx')
outDF.to_excel(writer)
writer.save()
