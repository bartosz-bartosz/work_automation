import pandas as pd

'''Put all customer numbers in customer_numbers list: '''
customer_numbers = []

''' Read file, drop columns '''
df = pd.read_excel(r'C:\Users\user\Desktop\east_tracker\export.XLSX')
drop_list = ['Profit Center', 'Company Code', 'Accounting Clerk', 'Payment Method']
df = df.drop(drop_list, axis=1)
df1 = df.iloc[:, 0:7]

''' Drop unnecessary rows '''
df1 = df1.dropna(subset=['Accounting Clerk Name'])
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'FOREIGN'])
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'FOREIGN LEGAL'])
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'Intercompany'])

''' Separate dataframes '''
customers_df = pd.DataFrame(columns=df1.columns)
for number in customer_numbers:
    customers_df = customers_df.append(df1.loc[df1['Customer'] == number])
    df1 = df1.drop(df1.index[df1['Customer'] == number])

df_media_old = df1.loc[df1['Accounting Clerk Name'] == 'Media/Saturn']
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'Media/Saturn'])

df_ar_legal_tps = df1.loc[df1['Accounting Clerk Name'] == 'AR LEGAL TPS']
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'AR LEGAL TPS'])

df_ar_cs_spareparts = df1.loc[df1['Accounting Clerk Name'] == 'CS-SPAREPARTS']
df1 = df1.drop(df1.index[df1['Accounting Clerk Name'] == 'CS-SPAREPARTS']) 

''' Save output file'''
with pd.ExcelWriter('east_tracker_output.xlsx') as writer:
    customers_df.to_excel(writer, sheet_name='Main customers')
    df_media_old.to_excel(writer, sheet_name='Media Old Entities')
    df_ar_legal_tps.to_excel(writer, sheet_name='AR LEGAL TPS')
    df_ar_cs_spareparts.to_excel(writer, sheet_name='CS-SPAREPARTS')
    df1.to_excel(writer, sheet_name='ALL THE REST')