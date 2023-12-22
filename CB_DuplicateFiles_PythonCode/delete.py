import pandas as pd
import numpy as np
import xlsxwriter as writer
from pandas.api.types import is_number
from openpyxl import load_workbook

# Importing spreadsheets and making dataframe
#sheet = pd.read_excel('df26multiple.xlsx')
sheet = pd.read_excel('Sheets/[INSERTNAMEHERE].xlsx')

df = pd.DataFrame(sheet)

# Getting length of spreadsheet (number of rows)
rows = len(df.index)
print(rows)

# Making a temp df to put the selected rows
dftemp = pd.DataFrame(df.iloc[0])
df = pd.concat([df, dftemp], ignore_index=True)
# print (df.columns)

#owners = df['Owner'].value_counts()
#print (owners)

#idx = df.index.get_indexer_for(df[df.Owner == '[multiple]'].index)
#print (idx)

#location = df.loc[idx, 'NumFiles']
#print (location)

#owners = df['Owner'].unique()
#print (owners)
#owners = [ x for x in owners if "[multiple]" not in x ]
#print (owners)
#numOwners = len(owners)

# Deleting rows where 
ind=0
j=0
keep="Keep keep"
n=0
k=0
a=np.array([])
for index in range(rows): 
    if index==0: #continue to start loop from second value
        continue
    action = df.loc[ind]['Action']
    num = df.loc[ind]['NumFiles']+1
    if num > 1:
        if action == '-':
            j=ind
            for ppl in range(num.astype(np.int64)): 
                if ppl==index:
                    continue
                # Figuring out if Action column cell is empty
                if ppl<num:
                    if pd.isnull(df.loc[j+1]['Action']):
                        print ("null")
                    elif df.loc[j+1]['Action'] in keep: # if the file is to be kept
                        print (df.loc[j+1]['Action'])
                        n = n+1
                    j = j+1
            numKeep = f"total: {num}, keep: {num-n}"
            print (numKeep)
        # If the number of files to keep is the same as NumFiles
        if n==num-1:
            k=ind
            for ppl in range(num.astype(np.int64)): 
                if ppl==index:
                    continue
                if ppl<=num:
                    a = np.append(a,[k]) # append all the rows of that duplicated file to array a
                    print (f"del: {k}")
                    k = k+1
        n=0
    ind = index #update ind with current index
print (a)
df = df.drop(a) # drop array a from df

# Export df with deletions to spreadsheet
df.to_excel('Sheets/with deletions/[INSERTNAMEHERE].xlsx')

