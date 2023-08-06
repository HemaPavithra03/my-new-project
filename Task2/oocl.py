import pandas as pd
import numpy as np
import itertools

# To read the 2 sheets separately in the excel file into 2 DataFrames
sheet_name_1 = 'Inland'
sheet_name_2 = 'Rates'
threshold_1 = 6
threshold_2 = 7

def read_file(sheet, threshold):
    xl = pd.ExcelFile(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC Container MT216737-070121-B + TPTWB GRI September 1 2021.xlsx')
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
    s.drop(s.columns[s.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)
    s.reset_index(drop=True)
    return s


s1_sheet_name_1 = read_file(sheet_name_1, threshold_1)
s1_sheet_name_2 = read_file(sheet_name_2, threshold_2)


# To read the 1st table in sheet 'Inland' into a single DataFrame
df_1 = pd.DataFrame()
# df_1 = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC Container MT216737-070121-B + TPTWB GRI September 1 2021.xlsx',
#                    sheet_name='Inland', skiprows=5, skipfooter=280)
df_1 = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC Container MT216737-070121-B + TPTWB GRI September 1 2021.xlsx',
                     sheet_name='Inland', skiprows=5, skipfooter=267)
commodity_1 = df_1['Commodity']

s1_sheet_name_1 = s1_sheet_name_1.reset_index()
N1 = s1_sheet_name_1[s1_sheet_name_1['index'] == 0].index
N1 = N1.to_list()
N1.append(len(s1_sheet_name_1))
n1 = [t - u for u, t in zip(N1, N1[1:])]
value1 = list(itertools.chain(*(itertools.repeat(elem, n1) for elem, n1 in zip(commodity_1, n1))))
s1_sheet_name_1['Commodity'] = value1
s1_sheet_name_1.drop(columns=['index'], inplace=True)

# To read the 1st table in sheet 'Rates' into a single DataFrame
df_2 = pd.DataFrame()
# df_2 = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC Container MT216737-070121-B + TPTWB GRI September 1 2021.xlsx',
#                    sheet_name='Rates', skiprows=5, skipfooter=2405)
df_2 = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC Container MT216737-070121-B + TPTWB GRI September 1 2021.xlsx',
                   sheet_name='Rates', skiprows=5, skipfooter=992)

df_2 = df_2.dropna(axis=1, how='all')
df_2 = df_2.iloc[:, 1:]
df_2['C'] = df_2.apply(lambda x: ', '.join(x.dropna()), axis=1)
commodity_2 = df_2['C']

s1_sheet_name_2 = s1_sheet_name_2.reset_index()
N2 = s1_sheet_name_2[s1_sheet_name_2['index'] == 0].index
N2 = N2.to_list()
N2.append(len(s1_sheet_name_2))
n2 = [t - u for u, t in zip(N2, N2[1:])]
# x1 = sum(n2[3:11])
# x2 = sum(n2[11:13])
# x3 = sum(n2[14:16])
# x4 = sum(n2[18:20])
x1 = sum(n2[3:5])
x2 = sum(n2[9:11])
# n = [n2[0:3], x1, x2, n2[13], x3, n2[16:18], x4, n2[20]]
n = [n2[0:3], x1, n2[5:9], x2, n2[11]]


def flatten(xs):
    result = []
    if isinstance(xs, (list, tuple)):
        for x in xs:
            result.extend(flatten(x))
    else:
        result.append(xs)
    return result


n = flatten(n)
value2 = list(itertools.chain(*(itertools.repeat(elem, n) for elem, n in zip(commodity_2, n))))
s1_sheet_name_2['Commodity'] = value2
s1_sheet_name_2.drop(columns=['index'], inplace=True)

# Task a - Sheet 'Inland' change the datetime format and delimiter format in 3 columns
s1_sheet_name_1['Effective MM/DD/YY'] = pd.to_datetime(s1_sheet_name_1['Effective MM/DD/YY'],
                                                       errors='coerce').dt.strftime('%Y-%m-%d')
# s1_sheet_name_1['Expiry MM/DD/YY'] = pd.to_datetime(s1_sheet_name_1['Expiry MM/DD/YY'],
#                                                     errors='coerce').dt.strftime('%Y-%m-%d')
s1_sheet_name_1['Cargo Nature'] = s1_sheet_name_1['Cargo Nature'].str.replace(';', ',')

# mapper = {"Effective MM/DD/YY": "Validity From", "Expiry MM/DD/YY": "Validity To",
#           "  Expiry   MM/DD/YY": "Validity To"}
mapper = {"Effective MM/DD/YY": "Validity From"}
s1_inland = s1_sheet_name_1.rename(columns=mapper)
print(s1_inland.shape)

# Task b - Sheet 'Rates' change the datetime format and delimiter format in 3 columns
s1_sheet_name_2['Effective MM/DD/YY'] = pd.to_datetime(s1_sheet_name_2['Effective MM/DD/YY'],
                                                       errors='coerce').dt.strftime('%Y-%m-%d')
s1_sheet_name_2['  Expiry   MM/DD/YY'] = pd.to_datetime(s1_sheet_name_2['  Expiry   MM/DD/YY'],
                                                    errors='coerce').dt.strftime('%Y-%m-%d')
s1_sheet_name_2['Cargo Nature'] = s1_sheet_name_2['Cargo Nature'].str.replace(';', ',')
# s1_sheet_name_2.drop(columns=['20RF', '40RF', '40RQ', '45RF'], axis=1, inplace=True)
mapper_1 = {"Effective MM/DD/YY": "Validity From", "Expiry MM/DD/YY": "Validity To",
           "  Expiry   MM/DD/YY": "Validity To"}
s1_rates = s1_sheet_name_2.rename(columns=mapper_1)
print(s1_rates.shape)

# For saving the dataframe into new excel workbook under a sheet names "OOCL_Inland", "OOCL_Rates"
writer = pd.ExcelWriter(r'C:\Users\Pavi\Desktop\Task_OOCL_new.xlsx', engine='xlsxwriter')

# To store the dataframes in a dict, where key is the sheet name
frames = {'Inland': s1_inland, 'Rates': s1_rates}
# To create a loop to insert each on a specific sheet
for sheet, frame in frames.items():
    frame.to_excel(writer, sheet_name=sheet, index=False, header=True)

writer.save()
