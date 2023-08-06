import glob
import pandas as pd
import numpy as np


# To append the 3 excel files into single DataFrame for analysing the data before writing it to a single excel file
df = pd.DataFrame()
for file in glob.glob(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC_*.xlsx'):
    df = df.append(pd.read_excel(file, header=[0, 1]), ignore_index=True)
print(df.shape)

# To merge the hierarchial/MultiIndex object to single Index object and updating the DataFrame
level_one = df.columns.get_level_values(0).astype(str)
level_two = df.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else i + ' ' + j for i, j in zip(level_one, level_two)]
df.columns = updated_columns
print(df.shape)

# First finding the position of 'from' in Commodity Note
df['from pos'] = df['Commodity Note'].str.split('to').str.get(0)
df['Valid From'] = df['from pos'].str.split('from').str.get(1)
# Finding the position of 'to' in Commodity Note
df['to pos'] = df['Commodity Note'].str.split('to').str.get(1)
df['Valid To'] = df['to pos'].str.split().str.get(0)
# Removing the columns that were used to get the position of the two date strings
df.drop(columns=['from pos', 'to pos'], axis=1, inplace=True)
# Changing the newly added date columns into datetime format
df['Valid From'] = pd.to_datetime(df['Valid From'])
df['Valid To'] = pd.to_datetime(df['Valid To'])
print(df.shape)

df['20TC'] = np.where(df['Rate(USD) Prefix'] == 'T', df['Rate(USD) 20'], '')
df['Rate(USD) 20'] = df['Rate(USD) 20'].astype(str)
df['Rate(USD) 20'] = df.apply(lambda x: x['Rate(USD) 20'].replace(x['20TC'], ''), axis=1)
df['Rate(USD) 20'] = pd.to_numeric(df['Rate(USD) 20'], errors='coerce')
print(df.shape)

# To add new columns to the existing DataFrame
# df.insert(3, 'Rate Status', 'ACTIVE')
# df.insert(2, 'Serial Number', range(15042021, 15042021 + len(df)))

# To rename a few columns in the existing DataFrame
# mappers = {"Destination Transmode": "D. Transmode",
#           "O.Via Code": "Origin Via Code",
#           "D.Via  Code": "Destination Via Code"}
# df1 = df.rename(columns=mappers, inplace=False)
# print(df1.columns)  # shows the change in the column names of the DataFrame used
# print(df1.shape)

# To drop a few columns in the existing DataFrame
# df1.drop(columns=['D.Call Y/N',
#                  'Rate(USD) Prefix',
#                  'Rate(USD) CGO TYPE',
#                  'Rate Status'
#                  'Origin Via Code',
#                  'Origin Transmode',
#                  'Serial Number',
#                  'Actual Customer Code'], axis=1, inplace=True)


# For saving the dataframe into new excel workbook under a sheet name "ONE"
df.to_excel(r'C:\Users\Pavi\Desktop\Task_ONE.xlsx', sheet_name='ONE', index=False, header=True)
