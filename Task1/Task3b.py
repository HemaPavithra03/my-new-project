import pandas as pd

# To read the excel files into a DataFrame
df = pd.DataFrame()
file = r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\HSD_Intern_Version.xlsx'
df= pd.read_excel(file, skiprows=13, header=[0,1])
df.head()

# To merge the hierarchial/MultiIndex object to single Index object and updating the DataFrame
df1 = df.rename(columns=lambda x: x if not 'Unnamed:' in str(x) else '')
level_one = df1.columns.get_level_values(0).astype(str)
level_two = df1.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else i + ' ' + j for i, j in zip(level_one, level_two)]
df1.columns = updated_columns

print(df.shape)
# To add new columns to the existing DataFrame
df1.insert(4, 'Rate Status', 'ACTIVE')
df1.insert(8, 'Serial Number', range(12021, 12021 + len(df1)))
# print(df.shape)

# To rename a few columns in the existing DataFrame
mappers = {"Origin Charges CUR": "Origin Charges Currency",
           "Ocean Charges CUR": "Ocean Charges Currency",
           "Destination Charges CUR": "Destination Charges Currency",
           "Pre-Carriage Tpt. Mode": "Pre-Carriage Transport Mode"}
df2 = df1.rename(columns=mappers, inplace=False)
print(df2.columns)  # shows the change in the column names of the DataFrame used
print(df2.shape)

# To drop a few columns in the existing DataFrame
df2.drop(columns=[' Named Account Role',
                  ' Named Account',
                  'Rate Status',
                  ' T/S3',
                  ' T/S4'], axis=1, inplace=True)

print(df2.shape)  # for highlighting the change in the dimension of the dataframe

# For saving the dataframe into new excel workbook under a sheet name "Task3b"
df2.to_excel(r'C:\Users\Pavi\Desktop\Task3b.xlsx', sheet_name='HSD_Intern_Version', index=False, header=True)