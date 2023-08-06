import pandas as pd
import numpy as np

# Task 1 and 2 - to read the 2 sheets separately in the excel file into 2 DataFrames
sheet_name_1 = 'EUROPE-USA (FAK)'
sheet_name_2 = 'Bullet Rates EUR to USA'
threshold_1 = 5
threshold_2 = 5


def read_file(sheet, threshold):
    xl = pd.ExcelFile(
        r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\MSC AC Container 18-544WW Appx (Rates) incl#32.xls')
    work_sheet = xl.parse(sheet_name=sheet)
    no_of_values = np.logical_not(work_sheet.isnull()).sum(axis=1)
    no_of_values_deltas = no_of_values[1:] - no_of_values[:-1].values
    beginning_of_table = no_of_values_deltas > threshold
    beginning_of_table = beginning_of_table[beginning_of_table].index
    ending_of_table = no_of_values_deltas < -threshold
    ending_of_table = ending_of_table[ending_of_table].index

    dfs = []
    for i in range(len(beginning_of_table)):
        start = beginning_of_table[i] + 1
        if i < len(ending_of_table):
            stop = ending_of_table[i]
        else:
            stop = work_sheet.shape[0]
        df = xl.parse(sheet_name=sheet, skiprows=start, nrows=stop - start)
        dfs.append(df)

    s = pd.concat(dfs)
    s.columns = s.columns.map(str)
    s.reset_index(drop=True)
    s = s.rename(columns=lambda x: x.replace("'", "").replace('"', '')).replace(" ", "")
    return s


s_sheet_name_1 = read_file(sheet_name_1, threshold_1)
s_sheet_name_2 = read_file(sheet_name_2, threshold_2)

s_sheet_name_1 = s_sheet_name_1.reset_index(drop=True)
s_sheet_name_2 = s_sheet_name_2.reset_index(drop=True)

pos_1 = s_sheet_name_1.columns.get_loc('40 DV HC')
s_sheet_name_1.insert(pos_1 + 1, '40STD', s_sheet_name_1['40 DV HC'])
s_sheet_name_1 = s_sheet_name_1.rename(columns={'40 DV HC': '40HC'})
# s_sheet_name_1['20 DV'] = s_sheet_name_1['20 DV'].astype(float).map('{}$'.format)
# s_sheet_name_1['40STD'] = s_sheet_name_1['40STD'].astype(float).map('{}$'.format)
# s_sheet_name_1['40HC'] = s_sheet_name_1['40HC'].astype(float).map('{}$'.format)
x = s_sheet_name_1.columns.get_loc('Type\nMove')
s_sheet_name_1.insert(x+1, 'Origin Shipment Mode', '')
s_sheet_name_1.insert(x+2, 'Destination Shipment Mode', '')
s_sheet_name_1[['Origin Shipment Mode', 'Destination Shipment Mode']] = s_sheet_name_1['Type\nMove']\
    .str.split('/', expand=True)
s_sheet_name_1.replace({"CY": "Port"}, inplace=True)
s_sheet_name_1.drop(columns=['Type\nMove'], inplace=True)

# To read the 2nd table in sheet 'EUROPE-USA (FAK)' into a single DataFrame
df_1 = pd.DataFrame()
df_1 = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\MSC AC Container 18-544WW Appx (Rates) '
                     r'incl#32.xls', sheet_name='EUROPE-USA (FAK)', skiprows=28, skipfooter=27)
df_1 = df_1.dropna(axis=1, how='all')
df_1 = df_1.dropna(how='all')
df_1.reset_index(drop=True, inplace=True)
A = df_1.iloc[np.r_[0:3]]
B = df_1.iloc[np.r_[3:7]]
C = df_1.iloc[np.r_[7:9]]

new_header = C.iloc[0] #grab the first row for the header
C = C[1:] #take the data less the header row
C.columns = new_header #set the header row as the df header
C = C.iloc[:, :-1]
col_names_c = ['40RF Additional', 'Reefer Remarks', 'Reefer Currency']
C.columns = col_names_c
C['40RF Additional'] = C['40RF Additional'].str.rsplit(',', 1).str.get(0)
C['40RF Additional'] = C['40RF Additional'].str.replace('$', '')
C['Reefer Currency'] = 'USD'

B.reset_index(drop=True, inplace=True)
B = B.append(B.iloc[0:2].agg(lambda x: ', '.join(x.dropna())).to_frame().T)
B = B[2:]
new_header_B = B.iloc[-1]
B.columns = new_header_B
B = B.iloc[:-1, :]
B = B.rename(columns=lambda x: x.replace("'", "").replace('"', '')).replace(" ", "")
B = B.rename(columns={'IMO Surcharge\nvatos': 'Remarks', '20': 'IMO Surcharges 20DC Price',
                      '40': 'IMO Surcharges 40DC Price', '': 'IMO Surcharge Additional Remarks'})
B['IMO Surcharges Currency'] = 'USD'
B.reset_index(drop=True, inplace=True)

A.reset_index(drop=True, inplace=True)
A = A.append(A.iloc[0:2].agg(lambda x: ', '.join(x.dropna())).to_frame().T)
A = A[2:]
new_header_A = A.iloc[-1]
A.columns = new_header_A
A = A.iloc[:-1, :]
A = A.rename(columns=lambda x: x.replace("'", "").replace('"', '')).replace(" ", "")
A = A.rename(columns={'20': 'Special Equipment Surcharge 20DC', '40': 'Special Equipment Surcharge 40DC',
                      '': 'Special Equipment Surcharge Currency'})
A['Special Equipment Surcharge Currency'] = 'USD'
A['Special Equipment Surcharge 20DC'] = A['Special Equipment Surcharge 20DC'].str.replace('$', '')
A['Special Equipment Surcharge 40DC'] = A['Special Equipment Surcharge 40DC'].str.replace('$', '')
A.reset_index(drop=True, inplace=True)

fr = [A, B, C]
re = pd.concat([A, B, C], axis=1)
result = pd.concat([s_sheet_name_1, re], axis=1)
print(result.shape)

pos_2 = s_sheet_name_2.columns.get_loc('40 DV HC')
s_sheet_name_2.insert(pos_2 + 1, '40STD', s_sheet_name_2['40 DV HC'])
s_sheet_name_2 = s_sheet_name_2.rename(columns={'40 DV HC': '40HC'})
# s_sheet_name_2['20 DV'] = s_sheet_name_2['20 DV'].astype(float).map('{}$'.format)
# s_sheet_name_2['40HC'] = s_sheet_name_2['40HC'].astype(float).map('{}$'.format)
# s_sheet_name_2['40STD'] = s_sheet_name_2['40STD'].astype(float).map('{}$'.format)
y = s_sheet_name_2.columns.get_loc('Type\nMove')
s_sheet_name_2.insert(x+1, 'Origin Shipment Mode', '')
s_sheet_name_2.insert(x+2, 'Destination Shipment Mode', '')
s_sheet_name_2[['Origin Shipment Mode', 'Destination Shipment Mode']] = s_sheet_name_2['Type\nMove']\
    .str.split('/', expand=True)
s_sheet_name_2.replace({"CY": "Port"}, inplace=True)
s_sheet_name_2.drop(columns=['Type\nMove'], inplace=True)

# Task 3 - to read the excel file into a single DataFrame
df = pd.DataFrame()
df = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\MSC AC Container 18-544WW Appx (Rates) incl#32.xls',
                   sheet_name=' ONC ', skiprows=12)
print(df.shape)

# To find the second table based on FAK Leschaco
pos = df.loc[df['FAK/NAC'] == 'FAK Leschaco'].index[0]
df1 = df.iloc[:pos-1, :]
df2 = df.iloc[pos+1:, :]
df1 = df1.dropna(how='all')
df2 = df2[np.where(df2['FAK/NAC'] == 'Origin Country')[0][0]:]
df2 = df2[1:]
df2.iloc[0, :] = df2.iloc[0, :].shift()
df1 = df1.rename(columns=lambda x: x.replace("'", "").replace('"', '')).replace(" ", "")
df2 = df2.rename(columns=lambda x: x.replace("'", "").replace('"', '')).replace(" ", "")
# df2['40 HR'] = pd.to_numeric(df2['40 HR'],errors='coerce')
df2['Valid as from'] = pd.to_datetime(df2['Valid as from'], format='%Y-%m-%d')


# Using the details before the main tables
# df1['Validity From'] = '2021-01-01 00:00:00'
# df1['Validity To'] = '2021-12-31 00:00:00'
# df2['Validity From'] = '2021-01-01 00:00:00'
# df2['Validity To'] = '2021-03-31 00:00:00'

# To merge the 2 tables into one single table and also updating the column names
l = [df1, df2]
data = pd.concat(l).reset_index(drop=True)
updated_columns = ['FAK/NAC', 'Origin Country', 'Destination Country', 'Destination City *', 'Destination State *',
                   'Destination Zip Code *', 'US POD', 'Transmode', 'Transittime (approx.)', '20 DV', '40 DV/HC',
                   '40 HR', 'Validity From', 'Validity To', 'Remarks']
data.columns = updated_columns

# To just pick the numbers from the columns
data['20 DV'] = data['20 DV'].str.replace('$', '')
data['40 DV/HC'] = data['40 DV/HC'].str.replace('$', '')
data['40 HR'] = data['40 HR'].str.replace('$', '')

# For saving the dataframe into new excel workbook under a sheet names "EUROPE-USA (FAK)", "Bullet Rates EUR to USA"
# and "MSC_ONC"
writer = pd.ExcelWriter(r'C:\Users\Pavi\Desktop\Task_MSC.xlsx', engine='xlsxwriter')

# To store the dataframes in a dict, where key is the sheet name
frames = {'EUROPE-USA (FAK)': result, 'Bullet Rates EUR to USA': s_sheet_name_2, 'MSC_ONE': data}
# To create a loop to insert each on a specific sheet
for sheet, frame in frames.items():
    frame.to_excel(writer, sheet_name=sheet, index=False, header=True)

writer.save()
