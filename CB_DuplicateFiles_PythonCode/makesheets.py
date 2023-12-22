import pandas as pd
import numpy as np
from pandas.api.types import is_number
from openpyxl import load_workbook

# Importing spreadsheet and making dataframe
df = pd.read_excel('CleanUpFileThing.xlsx', sheet_name='Duplicate Files')

df = pd.DataFrame(df)

# Gettign length of spreadsheet (number of rows)
rows = len(df.index)

# Getting al unique file owners
owners = df['Owner'].unique()
print (owners)

# and number of unique owners
numOwners = len(owners)

# Making a new empty df for each unique owner and exporting to excel
df2 = pd.DataFrame()
ind=0
for index in range(numOwners): 
    if index==0: 
        continue
    name = owners[ind]
    file = f"Sheets/{name}.xlsx"
    print (file)
    df2.to_excel(file)
    ind = index #update ind with current index

#df.to_excel('dftest.xlsx')
