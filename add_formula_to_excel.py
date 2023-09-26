# this will output an excel file called output.xlsx with two random arrays and mathematical operations on those two arrays

import pandas as pd
import numpy as np

arraySize = np.random.randint(100)

# generate random integer
datasetA_list = np.random.randint(100, size=arraySize)

datasetB_list = np.random.randint(100, size=arraySize)

dataset_list = ('sum', 'average', 'median', 'std deviation', 'count', 'correlation')

datasetA = pd.DataFrame(datasetA_list, columns=['ValueA'])
datasetB = pd.DataFrame(datasetB_list, columns=['ValueB'])
dataset_list_calcs = pd.DataFrame(dataset_list, columns=['Calcs'])

path = "output.xlsx"

workbook = pd.ExcelWriter(path, engine='openpyxl')

datasetA.to_excel(workbook,sheet_name='Sheet1', startrow=1, index=False, header=True)
datasetB.to_excel(workbook,sheet_name='Sheet1', startrow=1, startcol=2, index=False, header=True)
dataset_list_calcs.to_excel(workbook,sheet_name='Sheet1', startrow=1, startcol=4, index=False, header=True)

# Creating Calculations for datasetA

sheet = workbook.sheets['Sheet1']
sheet['E2'] = 'CalcsA'
sheet['F3'] = '=SUM(A3:A' + str(datasetA_list.size + 2) + ')'
sheet['F4'] = '=AVERAGE(A3:A' + str(datasetA_list.size + 2) + ')'
sheet['F5'] = '=MEDIAN(A3:A' + str(datasetA_list.size + 2) + ')'
sheet['F6'] = '=STDEV(A3:A' + str(datasetA_list.size + 2) + ')'
sheet['F7'] = '=COUNT(A3:A' + str(datasetA_list.size + 2) + ')'
sheet['F8'] = '=CORREL(A3:A' + str(datasetA_list.size + 2) + ',C3:C' + str(datasetA_list.size + 2) + ')'

# Creating Calculations for datasetB

sheet = workbook.sheets['Sheet1']
sheet['H2'] = 'CalcsB'
sheet['H3'] = '=SUM(C3:C' + str(datasetA_list.size + 2) + ')'
sheet['H4'] = '=AVERAGE(C3:C' + str(datasetA_list.size + 2) + ')'
sheet['H5'] = '=MEDIAN(C3:C' + str(datasetA_list.size + 2) + ')'
sheet['H6'] = '=STDEV(C3:C' + str(datasetA_list.size + 2) + ')'
sheet['H7'] = '=COUNT(C3:C' + str(datasetA_list.size + 2) + ')'
sheet['H8'] = '=CORREL(A3:A' + str(datasetA_list.size + 2) + ',C3:C' + str(datasetA_list.size + 2) + ')'

# Use Numpy to calculate the values

a = np.sum(datasetA_list)
b = np.average(datasetA_list)
c = np.median(datasetA_list)
d = np.std(datasetA_list, ddof=1)       # setting ddof to 0 will give a different result
f = np.count_nonzero(datasetA_list)
g = np.corrcoef(datasetA_list, datasetB_list)

print(a,b,c,d,f,g)

sheet['E14'] = 'Numpy Calculations'
sheet['E15'] = 'Sum'
sheet['E16'] = 'Average'
sheet['E17'] = 'Median'
sheet['E18'] = 'Standard Deviation'
sheet['E19'] = 'Count'
sheet['E20'] = 'Correlation'


sheet['F15'] = a
sheet['F16'] = b
sheet['F17'] = c
sheet['F18'] = d
sheet['F19'] = f
sheet['F20'] = str(g)

workbook.close()
