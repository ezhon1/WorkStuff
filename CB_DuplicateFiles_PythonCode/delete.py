import pandas as pd
import numpy as np
import xlsxwriter as writer
from pandas.api.types import is_number
from openpyxl import load_workbook

#sheet = pd.read_excel('df26multiple.xlsx')
sheet = pd.read_excel('Sheets/[INSERTNAMEHERE].xlsx')

df = pd.DataFrame(sheet)

rows = len(df.index) #length of spreadsheet (number of rows)
print(rows)

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
                if ppl<num:
                    if pd.isnull(df.loc[j+1]['Action']):
                        print ("null")
                    elif df.loc[j+1]['Action'] in keep:
                        print (df.loc[j+1]['Action'])
                        n = n+1
                    j = j+1
            numKeep = f"total: {num}, keep: {num-n}"
            print (numKeep)
        if n==num-1:
            k=ind
            for ppl in range(num.astype(np.int64)): 
                if ppl==index:
                    continue
                if ppl<=num:
                    a = np.append(a,[k])
                    print (f"del: {k}")
                    k = k+1
        n=0
    ind = index #update ind with current index
print (a)
df = df.drop(a)
## https://stackoverflow.com/questions/57911508/pandas-iterate-and-return-index-next-index-and-row


df.to_excel('Sheets/with deletions/[INSERTNAMEHERE].xlsx')

