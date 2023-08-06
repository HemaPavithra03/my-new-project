import glob
import pandas as pd
from functools import reduce

'''TASK 3a - dropping only the newly created columns'''

# To append 3 excel files into single DataFrame for analysing the data before writing it to a single excel file
df = pd.DataFrame()
for file in glob.glob(r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\AC_*.xlsx'):
    df = df.append(pd.read_excel(file, header=[0, 1]), ignore_index=True)
df.head()

# To merge the hierarchial/MultiIndex object to single Index object and updating the DataFrame
level_one = df.columns.get_level_values(0).astype(str)
level_two = df.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else i + ' ' + j for i, j in zip(level_one, level_two)]
df.columns = updated_columns

print(df.shape)  # for highlighting the change in the dimension of the dataframe

# To add new columns to the existing DataFrame
df.insert(3, 'Rate Status', 'ACTIVE')
df.insert(2, 'Serial Number', range(15042021, 15042021 + len(df)))

# To rename a few columns in the existing DataFrame
mappers = {"Serial Number": "Serial No.",
           "O.Via Code": "Origin Via Code",
           "D.Via  Code": "Destination Via Code"}
df1 = df.rename(columns=mappers, inplace=False)

print(df1.shape)  # for highlighting the change in the dimension of the dataframe

# To drop a few newly created columns (those which are created after combining the header levels)
df1.drop(columns=['D.Call Y/N',
                  'Rate(USD) Prefix',
                  'Rate(USD) CGO TYPE',
                  'Origin Via Code',
                  'Origin Transmode',
                  'Actual Customer Code',
                  'Rate Status',
                  'Serial No.'], axis=1, inplace=True)

print(df1.shape)  # for highlighting the change in the dimension of the dataframe

# First finding the position of from in Commodity Note
df1['from pos'] = df1['Commodity Note'].str.split('to').str.get(0)
df1['Valid From'] = df1['from pos'].str.split('from').str.get(1)
# Finding the position of to in Commodity Note
df1['to pos'] = df1['Commodity Note'].str.split('to').str.get(1)
df1['Valid To'] = df1['to pos'].str.split().str.get(0)
# Removing the columns that were used to get the position of the two date strings
df1.drop(columns=['from pos', 'to pos'], axis=1, inplace=True)
# Changing the newly added date columns into datetime format
df1['Valid From'] = pd.to_datetime(df1['Valid From'])
df1['Valid To'] = pd.to_datetime(df1['Valid To'])
print(df1.shape)

# not sure whether 'Commodity Note' column needs to be removed or not after extracting
# the 'valid from' and 'valid to' into two new columns
# if not needed we can drop 'Commodity Note' column while dropping the positioning
# columns in ln:46

# For saving the dataframe into new excel workbook under a sheet name "Task3a"
# df1.to_excel(r'C:\Users\Pavi\Desktop\Task3a.xlsx', sheet_name='Task3a', index=False, header=True)

'''TASK 3b - Transpose rows into column values for columns starting with "EQU" ; 
           - Rename column 'POD' to "Origin Locations" '''

# To read the excel file into a DataFrame
file = r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\HSD_Intern_Version.xlsx'
s = pd.read_excel(file, skiprows=14, header=0)
print(s.shape)  # for highlighting the change in the dimension of the dataframe

# Transposing 'EQU Size' and 'EQU Type' merged row values into column values based on price related columns
s1 = s
s1['EQUSize_EQUType'] = s['EQU Size'].map(str) + s['EQU Type'].map(str)
transposed_table = s1.pivot(index='Product ID', columns='EQUSize_EQUType',
                            values=['OFREIGHT', 'BAF', 'CUDE', 'DG ADD', 'ECA',
                                    'ISPS CAR', 'PAN CAN', 'CDC', 'GATE I O', 'PRECARR',
                                    'WHF EX', 'CAPAT IM', 'CLEANING',
                                    'CNT REL', 'CUST TEM', 'GENSET', 'ISPS IM',
                                    'ONCARR', 'THC IM']).fillna('')

# To merge the hierarchial/MultiIndex object to single Index object and updating the transposed_table
level_one = transposed_table.columns.get_level_values(0).astype(str)
level_two = transposed_table.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else j + ' ' + i for i, j in zip(level_one, level_two)]
transposed_table.columns = updated_columns
transposed_table.reset_index(inplace=True)

print("table before transposing the rows into columns is", s1.shape)
print("table after transposing the rows into columns is", transposed_table.shape)

# Considering a chain merge with 'reduce' on list of dataframes
dfs = [s1, transposed_table]
df_merged = reduce(lambda left, right: pd.merge(left, right, on=['Product ID'], how='outer'), dfs)
print("Overall shape of the newly updated table is", df_merged.shape)

# To rename a few columns in the existing DataFrame
mappers = {"POD": "Origin Locations",
           "40HC BAF": "40HC_BAF_Ocean_Charges",
           "20TK PRECARR": "20TK_PRECARR_Origin_Charges",
           "40RH PAN CAN": "40RH PAN CAN_Ocean_charges",
           "CDC": "CDC Price",
           "Pre-Carriage Tpt. Mode": "Pre-Carriage Transport Mode",
           "40RH THC IM": "40RH_THC_IM_destination_Charges"}
df_merged.rename(columns=mappers, inplace=True)

print(df_merged.columns)  # To highlight the transpose of 'EQU Size' and 'EQU Type' into multiple columns

# To drop a few columns in the existing DataFrame
df_merged.drop(columns=['Named Account Role',
                        'Named Account',
                        '40RH_THC_IM_destination_Charges',
                        'EQU Size',
                        'EQU Type',
                        '40RH PAN CAN_Ocean_charges',
                        '20TK_PRECARR_Origin_Charges',
                        'EQUSize_EQUType'], axis=1, inplace=True)
print("Overall table shape after renaming and dropping a few columns is", df_merged.shape)

df_output = df_merged
df_output.drop(df_output.loc[:, 'CDC Price':'CUR.18'], inplace=True, axis=1)
print("Table shape after retaining only the transposed price columns and dropping original price columns ", df_output.shape)

# For saving the dataframe into new excel workbook under a sheet name "Task3b"
df_output.to_excel(r'C:\Users\Pavi\Desktop\Task3b_price_col_transposed.xlsx', sheet_name='HSD_Intern_Version', index=False, header=True)
df_merged.to_excel(r'C:\Users\Pavi\Desktop\Task3b_price_col_double.xlsx', sheet_name='HSD_Intern_Version', index=False, header=True)