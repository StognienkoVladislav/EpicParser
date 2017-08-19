

from openpyxl import load_workbook

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

