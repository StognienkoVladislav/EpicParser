

import os
import pandas as pd


cwd = os.getcwd()
print(cwd)

print(os.listdir('.'))

###########################33


file = 'cripto2.xlsx'

x1 = pd.ExcelFile(file)

#print the sheet names
print(x1.sheet_names)

#Load a sheet into a DataFrame by name: df1

df1 = x1.parse('cripto')


#Specify a writer
writer = pd.ExcelWriter('cripto2.xlsx', engine = 'xlsxwriter')


#Write your DataFrame to a file
df1.to_excel(writer, 'Sheet1')

writer.save()
