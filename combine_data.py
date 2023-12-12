import pandas as pd

import warnings
warnings.filterwarnings("ignore")

table0 = pd.read_excel('shipadv to HQ 20231030.xlsx', header = 0, sheet_name='ShptComplete')
table1 = pd.read_excel('Shp adv-TWN 20231101.xlsx', header = 0, sheet_name='ShptComplete')
table2 = pd.read_excel('Price History-Cost Sample.xlsx', header = 2, skiprows=0)

table0['Vendor'] = table0['Vendor'].str.upper()
table1['Vendor'] = table1['Vendor'].str.upper()
table2['Vendor'] = table2['Vendor'].str.upper()


table0 = table0[["Vendor", "P/N", "Q\'ty ", "vendor confirm date", "ETD "]]
table0 = table0.fillna(0)
 
length = table0['vendor confirm date'].size
for i in range(length):
    if type(table0['vendor confirm date'][i]) == str:
        table0['vendor confirm date'][i] = 0
    if type(table0['ETD '][i]) == str:
        table0['ETD '][i] = 0

length = table0['vendor confirm date'].size
temp2 = table0
for i in range(length):
    temp2['vendor confirm date'][i] = str(temp2['vendor confirm date'][i])
    temp2['ETD '][i] = str(temp2['ETD '][i])
table0 = temp2
 
table0['vendor confirm date'] = table0['vendor confirm date'].replace('nan', 0)
table0['vendor confirm date'] = table0['vendor confirm date'].fillna(0)
table0['ETD '] = table0['ETD '].replace('nan', 0)
table0['ETD '] = table0['ETD '].fillna(0)
 
 
table0['vendor confirm date'] = table0['vendor confirm date'].str[:10]
table0['vendor confirm date'] = table0['vendor confirm date'].str.replace('-','')
table0['vendor confirm date'] = table0['vendor confirm date'].astype(int)
table0['ETD '] = table0['ETD '].str[:10]
table0['ETD '] = table0['ETD '].str.replace('-','')
table0['ETD '] = table0['ETD '].astype(int)



table1 = table1[["Vendor", "P/N", "Q\'ty ", "vendor confirm date", "ETD "]]
table1 = table1.fillna(0)
 
length = table1['vendor confirm date'].size
for i in range(length):
    if type(table1['vendor confirm date'][i]) == str:
        table1['vendor confirm date'][i] = 0
    if type(table1['ETD '][i]) == str:
        table1['ETD '][i] = 0

length = table1['vendor confirm date'].size
temp2 = table1
for i in range(length):
    temp2['vendor confirm date'][i] = str(temp2['vendor confirm date'][i])
    temp2['ETD '][i] = str(temp2['ETD '][i])
table1 = temp2
 
table1['vendor confirm date'] = table1['vendor confirm date'].replace('nan', 0)
table1['vendor confirm date'] = table1['vendor confirm date'].fillna(0)
table1['ETD '] = table1['ETD '].replace('nan', 0)
table1['ETD '] = table1['ETD '].fillna(0)
 
 
table1['vendor confirm date'] = table1['vendor confirm date'].str[:10]
table1['vendor confirm date'] = table1['vendor confirm date'].str.replace('-','')
table1['vendor confirm date'] = table1['vendor confirm date'].astype(int)
table1['ETD '] = table1['ETD '].str[:10]
table1['ETD '] = table1['ETD '].str.replace('-','')
table1['ETD '] = table1['ETD '].astype(int)

table = pd.concat([table0, table1], axis=0, join='inner', ignore_index=True)

#clean table2
vendor_cost = table2[["Item", "Vendor", "Cost to cost comp"]]
vendor_cost = vendor_cost.rename(columns={"Item" : "P/N"})
vendor_cost = vendor_cost.dropna(subset=["Vendor"])
vendor_cost['P/N'] = vendor_cost['P/N'].str.strip()

#combine table 0 and table 1
table_combined = vendor_cost[['P/N', 'Cost to cost comp']].merge(table,
                                                                   on = ['P/N'],
                                                                   how = "right")
print(table_combined.shape)
print(table_combined.head)

#calculate the value and put it in a new column
value = []
p = 0
length = table_combined['P/N'].size
for i in range(length):
    if table_combined['Cost to cost comp'].empty:
        value.append(None)
    else:
        p = table_combined['Cost to cost comp'][i] * table_combined['Q\'ty '][i]
        value.append(p)
            
table_combined['total_value'] = value

#get the date status for each
#0: invalid date
#1: on time
#-1: delay
status  = []
for i in range(length):
    #check if both date are valid
    if table['vendor confirm date'][i] != 0 and table['ETD '][i] != 0:
        if table['vendor confirm date'][i] <= table['ETD '][i]:
            status.append(1)
        else:
            status.append(-1)
    else:
        status.append(0)

table_combined['status'] = status

print(table_combined)

table3 = pd.read_excel("-----------NCR & QA Master List-----------.xlsx", header = 0, sheet_name='NCR List')

#make NCR table uniform
table_combined['Vendor'] = table_combined['Vendor'].str.upper()
table3['Vendor'] = table3['Vendor'].str.upper()
table3 = table3.replace('JY', 'JIAYE')

#get the date of the NCR and put them in a new column
table3['Ref Document'] = table3['Ref Document'].fillna(0)
ncr = table3[['Ref Document', 'Part Number']]
ncr['Ref Date'] = ncr['Ref Document'].str[:8]
ncr['Ref Date'] = ncr['Ref Date'].fillna(0)
ncr['Ref Date'].astype(int)
ncr = ncr.rename(columns={"Part Number" : "P/N"})

#merge two file together, all information needed are here
outputList = table_combined.merge(ncr,
                                    on = ['P/N'],
                                    how = "left")

#data type conversion. for calculations later
outputList['Ref Date'] = outputList['Ref Date'].fillna(0)

print(ncr)
print(outputList)
outputList.to_csv('combined_table.csv', mode='a', index=False, header=True)

