import pandas as pd
import numpy as np
from pandas.api.types import is_number
from openpyxl import load_workbook

df = pd.read_excel('CleanUpFileThing.xlsx', sheet_name='Duplicate Files')

df = pd.DataFrame(df)

rows = len(df.index)

owners = df['Owner'].unique()
print (owners)

numOwners = len(owners)

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