import pandas as pd
import numpy as np


# To read the excel file into a DataFrame
file = r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\HSDG -  AMGC0000164 -   AMD 1 -valid 30.09.2021.xlsx'
s = pd.read_excel(file, skiprows=14, header=0)
print(s.shape)  # for highlighting the change in the dimension of the dataframe

# Transposing 'EQU Size' and 'EQU Type' merged row values into column values based on price related columns
s1 = s
s1['EQUSize_EQUType'] = s1['EQU Size'].map(str) + s1['EQU Type'].map(str)
s1['EQUSize_EQUType_changed'] = np.where(s1['EQUSize_EQUType'] == '40HC', s1['EQUSize_EQUType'], '')
transposed_table = s1.pivot(index='Product ID', columns='EQUSize_EQUType_changed',
                            values=['OFREIGHT', 'BAF', 'CUDE', 'DG ADD', 'ECA',
                                    'ISPS CAR', 'PAN CAN', 'CDC', 'GATE I O', 'PRECARR',
                                    'WHF EX', 'CAPAT IM', 'CLEANING',
                                    'CNT REL', 'CUST TEM', 'GENSET', 'ISPS IM',
                                    'ONCARR', 'THC IM']).fillna('')

# To merge the hierarchial/MultiIndex object to single Index object and updating the transposed_table w.r.t "40HC"
level_one = transposed_table.columns.get_level_values(0).astype(str)
level_two = transposed_table.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else j + '' + i for i, j in zip(level_one, level_two)]
transposed_table.columns = updated_columns

s1['EQUSize_EQUType_changed'].replace({"40HC": "40STD"}, inplace=True)
transposed_table_std = s1.pivot(index='Product ID', columns='EQUSize_EQUType_changed',
                                values=['OFREIGHT', 'BAF', 'CUDE', 'DG ADD', 'ECA',
                                    'ISPS CAR', 'PAN CAN', 'CDC', 'GATE I O', 'PRECARR',
                                    'WHF EX', 'CAPAT IM', 'CLEANING',
                                    'CNT REL', 'CUST TEM', 'GENSET', 'ISPS IM',
                                    'ONCARR', 'THC IM']).fillna('')

# To merge the hierarchial/MultiIndex object to single Index object and updating the transposed_table w.r.t "40STD"
level_one_std = transposed_table_std.columns.get_level_values(0).astype(str)
level_two_std = transposed_table_std.columns.get_level_values(1).astype(str)
updated_columns_std = [j if i == j else j + '' + i for i, j in zip(level_one_std, level_two_std)]
transposed_table_std.columns = updated_columns_std

# To join the DataFrames after 40HC "EQUSize_EQUType" split and additional cells of 40STD
cols = transposed_table_std.columns.difference(transposed_table.columns)
transposed_table = transposed_table.join(transposed_table_std[cols])
cols_1 = transposed_table.columns.intersection(s1.columns)
s1 = s1.drop(columns=cols_1)
transposed_table.reset_index(inplace=True)
result = pd.merge(s1, transposed_table, on='Product ID')
result.drop(columns=['EQUSize_EQUType', 'EQUSize_EQUType_changed'], axis=1, inplace=True)
print(result.shape)

# For saving the dataframe into new excel workbook under a sheet name "HSDG"
result.to_excel(r'C:\Users\Pavi\Desktop\Task_HSDG.xlsx', sheet_name='HSDG', index=False, header=True)
