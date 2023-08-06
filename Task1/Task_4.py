import re
import pandas as pd
import numpy as np
import random

xl_file = pd.ExcelFile(r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\CMA 20-2615 _Task- valid 31.07.2021.xlsx')
worksheets = xl_file.sheet_names
df_table_1 = pd.DataFrame()
df_table_2 = pd.DataFrame()
for sheet in worksheets:
    if re.match('APPENDIX', sheet):  # when matching pattern add the dataframe to the list
        xl = pd.ExcelFile(xl_file)
        work_sheet = xl.parse(sheet_name=sheet, header=None, skiprows=10)
        work_sheet_1 = work_sheet[~work_sheet[0].str.split().str.len().gt(5)]
        work_sheet_1.reset_index(drop=True, inplace=True)
print(work_sheet_1.shape)

table_1 = ['RATES CONDITIONS']
table_2 = ['SPECIAL EQUIPMENT, HAZARDOUS, SOC, OOG...']


def read_file(table):
    xl_file = pd.ExcelFile(r'C:\Users\Pavi\PycharmProjects\TestProject\Task1\CMA 20-2615 _Task- valid 31.07.2021.xlsx')
    worksheets = xl_file.sheet_names
    df = pd.DataFrame()
    for sheet in worksheets:
        if re.match('APPENDIX', sheet):  # when matching pattern add the dataframe to the list
            xl = pd.ExcelFile(xl_file)
            work_sheet = xl.parse(sheet_name=sheet, header=None, skiprows=10)
            work_sheet_1 = work_sheet[~work_sheet[0].str.split().str.len().gt(5)]
            work_sheet_1.reset_index(drop=True, inplace=True)
            # to pick only the tables which we are interested in - to get the indexes for beginning of the tables
            idx = work_sheet_1.index[work_sheet_1.isin(table).any(1)]
            # to get the indexes for end of the tables
            val = [work_sheet_1.loc[i:].isnull().all(1).idxmax() for i in idx]
            inx = idx.to_list()
            # this helps to remove blank row right after the end of each table (which was used to differentiate 2 tables)
            inx = [id + 1 for id in inx]
            for i, j in zip(range(len(inx)), range(len(val))):
                df = df.append(work_sheet_1.iloc[inx[i]:val[j]])
            df = df.dropna(how='all', axis=1)
            df.reset_index(drop=True, inplace=True)
    return df


df_table_1 = read_file(table_1)
df_table_2 = read_file(table_2)
print(df_table_1.shape)
print(df_table_2.shape)

search = ['FAK/ BULLETS', 'IPI Construction']


def final_table(df1):
    # to obtain the index of the beginning of each table (i.e. column headers) and remove the redundant rows
    pos = df1.index[df1.isin(search).any(1)]
    rows = pos.values
    last_num = rows[0]
    new_list = []
    for x in rows[1:]:
        if x - last_num == 1:
            new_list.append(last_num)
        last_num = x
    v = df1.index[new_list]
    df1.drop(v, inplace=True)
    ix = df1.iloc[:, :1].dropna(how='all').index.tolist()
    df_final = df1.loc[ix]

    # to make sure headers which are added everytime when each table is appended is addressed in separate DataFrame
    mask = df1.index[df1.isin(search).any(1)]
    s = mask.values
    l = mask.to_list()
    for x in s[0:]:
        a = x + 1
        l.append(a)
    l1 = sorted(l)
    df_final = df_final.loc[df_final.index.difference(l1)]  # this DataFrame holds only data fields of all the tables
    df_headers = df1.reindex([x for x in df1.index if x in l1])  # this holds only the column fields of all the tables
    o = df_headers.count(axis=1).idxmax()  # to identify the column fields for resultant table -
    # based on the row which has max non-nan values

    # to make sure the multi-header format to single header format pattern is met (to some extent, it follows the pattern)
    df_headers.loc[o + 1] = df_headers.loc[o + 1].replace(np.nan, ' ')
    df_headers = df_headers.astype(str)
    col = df_headers.loc[o] + ' ' + df_headers.loc[o + 1]
    df_final.columns = col
    df_final.reset_index(drop=True, inplace=True)
    return df_final, df_headers


df_final_table_1, df_headers_table_1 = final_table(df_table_1)  # comprises of all the table1 values across diff sheets
df_final_table_2, df_headers_table_2 = final_table(df_table_2)  # comprises of all the table2 values across diff sheets

# to help merge the similar column fields from table1 and table2 into one single DataFrame
df_final_table_1 = df_final_table_1.rename(columns={'D20  ': '20  ', 'D40  ': '40  '})
df_final = pd.concat([df_final_table_1, df_final_table_2], axis=0, ignore_index=True)
print(df_final.shape)

# to split the column "MODE" into 2 columns and it is appended at the end of the resultant table
df_final = pd.concat([df_final.loc[:, ~df_final.columns.isin(['MODE  ', 'Origin Mode', 'Destination Mode'])],
                      df_final['MODE  '].str.split('/', expand=True).rename(columns={0: 'Origin Mode', 1:
                                                                                     'Destination Mode'})], axis=1)

# to reverse the abbreviations used under the column "MODE"
df_final[['Origin Mode', 'Destination Mode']] = df_final[['Origin Mode', 'Destination Mode']].replace(['CY', 'R', 'M',
                                                                                                       'RM', 'B', 'RB',
                                                                                                       'BM'],
                                                                                                      ['PORT', 'RAIL',
                                                                                                       'MOTOR',
                                                                                                       'RAIL/MOTOR',
                                                                                                       'BARGE',
                                                                                                       'RAIL/BARGE',
                                                                                                       'MOTOR/BARGE'])

# To add new columns to the existing table - 3 ways are used here
print(df_final.shape)  # for highlighting the original dimension of the dataframe
df_final['Check Flag'] = pd.Series(random.choices(['yes', 'no'], weights=[5, 2], k=len(df_final)))
df_final.insert(0, 'Serial Number', range(28052021, 28052021 + len(df_final)))
sWeight = len(df_final['20  '])
df_final = df_final.assign(Weight=pd.Series(np.random.randn(sWeight)).values)
print(df_final.shape)

# To rename a few columns in the existing DataFrame
mappers = {"20  ": "D20 ",
           "40  ": "D40 ",
           "Effective date  ": "Effective Date ",
           "P/C EIFS Appl/ Not Appl": "P/C (EIFS Appl/ Not Appl)",
           "FR Appl/ Not Appl": "FR (Appl/ Not Appl)",
           "Hazardous Yes / No": "Hazardous (Yes / No)",
           "OT Appl/ Not Appl": "OT (Appl/ Not Appl)",
           "EIS Org Appl/ Not Appl": "EIS Org (Appl/ Not Appl)",
           "Seal Fee Appl/ Not Appl": "Seal Fee (Appl/ Not Appl)"}
df_final_renamed = df_final.rename(columns=mappers, inplace=False)

# To drop a few columns in the existing DataFrame
df_final_renamed.drop(columns=['Place of Delivery  ', 'nan  ', 'Note  ', 'Weight', 'Serial Number', 'FAK/ BULLETS  ',
                               'Check Flag', 'IPI Construction  ', 'Shipper own SOC\nCOC',
                               'OOG IG (In gauge)\nOOG (Out Of Gauge)', 'SDD (Origin; Dest)  '], axis=1, inplace=True)
print("Overall table shape after adding, renaming and dropping a few columns is", df_final_renamed.shape)
print(df_final_renamed.columns)

# For saving the dataframe into new excel workbook
df_final_renamed.to_excel(r'C:\Users\Pavi\Desktop\CMA 20-2615_Task_4.xlsx', sheet_name='APPENDIX', index=False,
                          header=True)
