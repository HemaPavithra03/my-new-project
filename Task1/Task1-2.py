import pandas as pd
import numpy as np

# To parse multiple tables from an excel sheet into multiple data frame objects.
file = r'C:\Users\Pavi\PycharmProjects\TestProject\Excel_Test_Version.xlsx'
sheet_name = 0
threshold = 5  # to help in estimating the start and end of each table in the excel sheet
xl = pd.ExcelFile(file)
work_sheet = xl.parse(sheet_name=sheet_name)

# To count the number of non-Nan cells in each row and then the change in that number between adjacent rows
no_of_values = np.logical_not(work_sheet.isnull()).sum(axis=1)
no_of_values_deltas = no_of_values[1:] - no_of_values[:-1].values

# To define the beginning and end of each table in the sheet using delta in no_of_values

beginning_of_table = no_of_values_deltas > threshold
beginning_of_table = beginning_of_table[beginning_of_table].index
ending_of_table = no_of_values_deltas < -threshold
ending_of_table = ending_of_table[ending_of_table].index
# To make a list of data frames
dfs = []
for i in range(len(beginning_of_table)):
    start = beginning_of_table[i] + 1
    if i < len(ending_of_table):
        stop = ending_of_table[i]
    else:
        stop = work_sheet.shape[0]
    df = xl.parse(sheet_name=sheet_name, skiprows=start, nrows=stop - start)
    dfs.append(df)

# To make list of dataframes into a single dataframe
s = pd.concat(dfs)
s.columns = s.columns.map(str)
s.drop(s.columns[s.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)

# Task 2
# To add a new column to the existing table
a = s.shape
print(a)  # for highlighting the original dimension of the dataframe
s.insert(0, 'ID' , -1)
s.insert(1, 'Rate Status' , 'ACTIVE')
s.insert(2, 'Tariff Type' , 'buy')
s.insert(3, 'LegType' , 'FCL Mainleg')
s.insert(4, 'Transport Modes' , 'FCL')
s.insert(5, 'Supplier' , 'OOCL')
s.insert(6, 'Service Contract' , 'MT206737')
s.insert(7, 'Serial Number', range(15042021, 15042021 + len(s)))
print(s.shape)  # for highlighting the change in the dimension of the dataframe

# To delete a column from the existing table
s = s.drop(columns=['Notes',
                    'Remark',
                    'Serial Number',
                    'Svc Loop'], axis=1)
print(s.shape)  # for highlighting the change in the dimension of the dataframe

# To rename an existing column in the table
mapper = {"Origin": "Origin Locations", "Destination": "Destination Locations",
           "Via (Origin)": "Via Origin", "Via (Dest.)": "Via Destination",
           "Service Contract": "Service Mode",
           "Effective MM/DD/YY": "Validity From", "  Expiry   MM/DD/YY": "Validity To",
           "20": "20DC Price", "40": "40DC Price", "45": "45", "40H": "40HC Price"}
s1 = s.rename(columns=mapper, inplace=False)
# as per pandas documentation DataFrame.rename() function syntax "columns" is one of the passing parameter
# it works well
# but unfortunately here it is throwing an warning called "Unexpected argument"

# For saving the dataframe into new excel workbook under a sheet name "Task1"
s1.to_excel(r'C:\Users\Pavi\Desktop\Excel_Test_Version_Task1-2.xlsx', sheet_name='Rates', index=False, header=True)
