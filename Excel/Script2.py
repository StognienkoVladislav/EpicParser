from itertools import islice
from openpyxl import load_workbook, Workbook

#Load in the workbook

wb = load_workbook('test.xlsx')

#Get sheet names
print(wb.get_sheet_names())



#Get a sheet by name
sheet = wb.get_sheet_by_name('Sheet3')

#Print the sheet title
print(sheet.title)

#Get currently active sheet
anotherSheet = wb.active

#Check `anotherSheet`
print(anotherSheet)


#########################################################


#Retrieve the value of a certain cell
print(sheet['A1'].value)

#Select element 'B2' of your sheet
c = sheet['B2']

#Retrieve the row number of your element
print('row : ' + str(c.row))

#Retrieve the column letter of your element
print('column : ' + str(c.column))

#Retrieve the coordinates of the cell
print("coord : " + str(c.coordinate))




#Retrieve cell value
print(sheet.cell(row = 1, column = 2).value)


#Print out values in column 2

for i in range(1, 4):
    print(i, sheet.cell(row = i, column = 2).value)


#################################################

from openpyxl.utils import get_column_letter, column_index_from_string

#Return 'A'
print(get_column_letter(1))

#Return '1'
print(column_index_from_string('A'))



#Print row per row
for cellObj in sheet['A1':'C3']:
    for cell in cellObj:
        print(cell.coordinate, cell.value)

    print('----END----')



#Retrieve the maximum amount of rows
print(sheet.max_row)

#Retrieve the maximum amount of column
print(sheet.max_column)



############################################

import pandas as pd

#Convert Sheet to DataFrame
df = pd.DataFrame(sheet.values)

#Put the sheet values in `data`
data = sheet.values

#Indicate the columns in the sheet values
cols = next(data)[1:]

#Convert your data to a list
data = list(data)

idx = [r[0] for r in data]

#Slice the data at index 1

data = (islice(r, 1, None)for r in data)

#Make your DataFrame
df = pd.DataFrame(data, index = idx, columns = cols)



#Import `dataframe_to_rows`
from openpyxl.utils.dataframe import  dataframe_to_rows

#Initialize a workbook
wb = Workbook()

#Get the workshet in the active workbook
ws = wb.active


#Append the rows of the DataFrame to your worksheet
for r in dataframe_to_rows(df, index = True, header = True):
    ws.append(r)

