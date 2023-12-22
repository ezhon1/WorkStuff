import pandas as pd
import numpy as np
import xlsxwriter as writer
from pandas.api.types import is_number
from openpyxl import load_workbook

#sheet = pd.read_excel('df26multiple.xlsx')
sheet = pd.read_excel('listedPeeps.xlsx')

df = pd.DataFrame(sheet)

pf = pd.read_excel("Sheets/[INSERTNAMEHERE].xlsx")
pf = pd.DataFrame(pf)

rows = len(df.index) #length of spreadsheet (number of rows)
print(rows)

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
n=0
x=0
dftemp = pd.DataFrame()
for index in range(rows): 
    if index==0: #continue to start loop from second value
        continue
    owner = df.loc[ind]['Owner']
    num = df.loc[ind]['NumFiles']+1
    if num > 1:
        if "[INSERTNAMEHERE]" in owner:
            print (f"1owner: {owner}")
            print (num)
            j=ind
            for ppl in range(num.astype(np.int64)): 
                if ppl==index:
                    continue
                if ppl<num:
                    newRow=df.iloc[[j]]
                    #print(len(newRow))
                    #print (f"{newRow} ---{j}")
                    dftemp = pd.concat([dftemp, newRow], ignore_index=True)
                    j=j+1
                print (len(dftemp))
            file = "Sheets/[INSERTNAMEHERE].xlsx"
            print (f"3file: {file}")
            pf = pd.concat([pf, dftemp], ignore_index=True)
            pf.to_excel("Sheets/[INSERTNAMEHERE].xlsx")
        

        #if '[INSERTNAMEHERE]' in owner:
        #    for numPeeps in range(numOwners): 
        #        if numPeeps==0: 
        #            continue
        #        name = owners[x]
        #        print(f"2name: {x}{name}")
        #if "[INSERTNAMEHERE]" in owner:
        #    file = ".xlsx"
        #    print (f"3file: {file}")
        #    pf = pd.concat([pf, dftemp], ignore_index=True)
        #    pf.to_excel(".xlsx")

                    #pf = pd.ExcelFile('Sheets/[INSERTNAMEHERE].xlsx', engine='xlsxwriter')
                    #with pd.ExcelWriter("dftemp.xlsx", engine="openpyxl") as writer:
                     #   dftemp.to_excel(writer, startrow=writer.max_row, index = False,header= False)

                    #writer = pd.ExcelWriter('Sheets/[INSERTNAMEHERE].xlsx', engine='openpyxl')
                    #dftemp.to_excel(writer, startrow=writer.end, index = False,header= False)

                    #writer.save()  
            #    x = x+1
            #x=0    
            
        #if '[INSERTNAMEHERE]' not in owner:
        #    name = owner
        #    file = f"{name}.xlsx"
        #    print (f"3file: {file}")
        #    dftemp.to_excel(file)
        dftemp = pd.DataFrame() 
    ind = index #update ind with current index\

