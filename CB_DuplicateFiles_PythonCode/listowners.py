import pandas as pd
import numpy as np
from pandas.api.types import is_number
from openpyxl import load_workbook

# Importing spreadsheet and making dataframe
df = pd.read_excel('df26.xlsx')

df = pd.DataFrame(df)

# Finding the length of spreadsheet (number of rows)
rows = len(df.index)
print(rows)

# print (df.columns)

#owners = df['Owner'].value_counts()
#print (owners)

# Getting the index for all the files without multiple owners
idx = df.index.get_indexer_for(df[df.Owner != '[multiple]'].index)
print (idx)

location = df.loc[idx, 'NumFiles']
#print (location)

# listing the owners of duplications with multiple owners in the duplicate collection row
ind=0
j=0
for index in range(rows): 
    if index==0: #continue to start loop from second value
        continue
    owner = df.loc[ind]['Owner']
    num = df.loc[ind]['NumFiles']
    if num > 0:
        if owner == '[multiple]':
            j=ind
            for ppl in range(num.astype(np.int64)): 
                if ppl==index:
                    continue
                # Appending owners not already in the [multiple] cell
                if ppl<=num:
                    if df.loc[j+1]['Owner'] not in owner:
                        owner = f"{owner}, {df.iloc[j+1]['Owner']}"
                        print(f"{num}/{ppl}, {j}, {df.loc[j+1]['Owner']}")
                    j = j+1
            newOwner = f"ind:{ind}, NumFiles:{num}, Owner:{owner}"
            print (newOwner)
    df.at[ind, 'Owner'] = owner #updating cell
    ind = index #update ind with current index

# Exporting df as spreadsheet
df.to_excel('listedPeeps.xlsx')

#df2 = pd.read_excel('dftestfull.xlsx')
#df2 = pd.DataFrame(df2)
#print (len(df2.index))
