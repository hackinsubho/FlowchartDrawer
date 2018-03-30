import pandas as pd
import os
cwd = os.getcwd()
cwd
os.chdir("/home/cyrus/PycharmProjects/flowchart")
os.listdir('.')

file = 'example.xlsx'
xl = pd.ExcelFile(file)
print(xl.sheet_names)
df1 = xl.parse('Sheet1')
print (df1[0].count)