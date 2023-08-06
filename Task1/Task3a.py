import glob
import pandas as pd

# To append the 3 excel files into single DataFrame for analysing the data before writing it to a single excel file
df = pd.DataFrame()
for file in glob.glob(r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\AC_*.xlsx'):
    df = df.append(pd.read_excel(file, header=[0, 1]), ignore_index=True)
df.head()

# print(df.columns)  # it returns a MultiIndex object since the data is having multi level column (header)

# To merge the hierarchial/MultiIndex object to single Index object and updating the DataFrame
level_one = df.columns.get_level_values(0).astype(str)
level_two = df.columns.get_level_values(1).astype(str)
updated_columns = [j if i == j else i + ' ' + j for i, j in zip(level_one, level_two)]
df.columns = updated_columns

print(df.shape)
# To add new columns to the existing DataFrame
df.insert(3, 'Rate Status', 'ACTIVE')
df.insert(2, 'Serial Number', range(15042021, 15042021 + len(df)))

# To rename a few columns in the existing DataFrame
mappers = {"Destination Transmode": "D. Transmode",
           "O.Via Code": "Origin Via Code",
           "D.Via  Code": "Destination Via Code"}
df1 = df.rename(columns=mappers, inplace=False)
print(df1.columns)  # shows the change in the column names of the DataFrame used
print(df1.shape)

# To drop a few columns in the existing DataFrame
# df1.drop(columns=['Route Note',
#                   'Rate Status'], axis=1, inplace=True)
df1.drop(columns=['D.Call Y/N',
                  'Rate(USD) Prefix',
                  'Rate(USD) CGO TYPE',
                  'Origin Via Code',
                  'Origin Transmode',
                  'Actual Customer Code'], axis=1, inplace=True)

print(df1.shape)  # for highlighting the change in the dimension of the dataframe

# df2 = df1
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
df1.to_excel(r'C:\Users\Pavi\Desktop\Task3a.xlsx', sheet_name='Task3a', index=False, header=True)
