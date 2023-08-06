import pandas as pd


# To read the excel file into a single DataFrame
df = pd.DataFrame()
df = pd.read_excel(r'C:\Users\Pavi\PycharmProjects\TestProject\Task2\AC HON21060 COSCO Customer rate 032421.xlsx')
print(df.shape)

# To split column Mode into 2 columns and replacing/reversing the abbreviation Y to Port
x = df.columns.get_loc('Mode')
df.insert(x+1, 'Origin Shipment Mode', df['Mode'].str[0])
df.insert(x+2, 'Destination Shipment Mode', df['Mode'].str[1:])
df.replace({"Y": "Port"}, inplace=True)
print(df.shape)

# Creating extra columns with default values and skipping index at the end
df['ID'] = '-1'
df['Rate Status'] = 'INACTIVE'
df['Tariff Type'] = 'buy'
df['Legtype'] = 'FCL Mainleg'
df['Transport Modes '] = 'FCL'
df['Supplier'] = 'COSCO'
df['Service Contract'] = 'HON21060'
df['Validity From'] = '2021-01-13 00:00:00'

# Renaming the column names
df = df.rename(columns={'Effective End Date': 'Validity To',
                          'Commodity Group': 'Commodities',
                          'Destination Via': 'Via Destination',
                          'Included Surcharges': 'Routing Info Origin',
                          'Subject To Surcharges': 'Remarks',
                          'Rate 20': '20DC Price ',
                          'Per Unit': '20DC Pricing Type',
                          'Rate 40': '40DC Price ',
                          'Rate 40H': '40HC Price'})

df['Origin Locations'] = df['Origin']
df['Destination Locations'] = df['Destination']
df['Remarks'] = "Subject to Tariff Charges - " + df['Remarks']
df['Routing Info Origin'] = "Included Surcharges - " + df['Routing Info Origin']

# To drop the column Mode
df.drop(columns=['Mode'], axis=1, inplace=True)

# For saving the dataframe into new excel workbook under a sheet name "COSCO"
df.to_excel(r'C:\Users\Pavi\Desktop\Task_COSCO.xlsx', sheet_name='COSCO', index=False, header=True)
