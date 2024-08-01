#%% import pandas and files and read files (you need to change the directory of the files)
import pandas as pd

Pais2020 = pd.read_excel(r'C:\Users\LEGION\Lenovo\Documents\VistoConsulting\expo_pais_sa_peso_0112_2020_v2.xlsx')
Pais2021 = pd.read_excel(r'C:\Users\LEGION\Lenovo\Documents\VistoConsulting\expo_pais_sa_peso_0112_2021_v2.xlsx')
Pais2022 = pd.read_excel(r'C:\Users\LEGION\Lenovo\Documents\VistoConsulting\expo_pais_sa_peso_0112_2022.xlsx')
Pais2023 = pd.read_excel(r'C:\Users\LEGION\Lenovo\Documents\VistoConsulting\expo_pais_sa_peso_0112_2023_v2.xlsx')
Pais2024 = pd.read_excel(r'C:\Users\LEGION\Lenovo\Documents\VistoConsulting\expo_pais_sa_peso_0106_2024.xlsx')
#%% Concat all the Excel files
Pais = pd.concat([Pais2020, Pais2021, Pais2022, Pais2023, Pais2024])
#%% Drop the first two rows
Pais2 = Pais._drop_axis(Pais.index[0:2], axis=0)
#%% drop the first column Unnamed: 0
Pais3 = Pais2.drop('Unnamed: 0', axis=1)
#%% filter the code 1001 and 1005 and conserve the countries
Pais4 = Pais3[Pais3['Unnamed: 1'].str.startswith(('1001', '1005')) | Pais3['Unnamed: 1'].str.match(r'^[A-Z]')]
#%%
Pais5 = Pais4.copy()
Pais5['Country'] = ''
Pais5 = Pais5.reset_index()
for i in range(len(Pais5)):
    current_value = Pais5.at[i, 'Unnamed: 1']

    if isinstance(current_value, str) and current_value[0].isalpha():
        Pais5.at[i, 'Country'] = current_value
    elif isinstance(current_value, str) and current_value[0].isdigit():
        if i > 0:
            Pais5.at[i, 'Country'] = Pais5.at[i - 1, 'Country']
#%% drop the column index
Pais6 = Pais5.drop(columns='index')
#%% Reorder columns in Pais6
column_order = ['Unnamed: 1', 'Country', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6',
                'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9', 'Unnamed: 10', 'Unnamed: 11', 'Unnamed: 12',
                'Unnamed: 13', 'Unnamed: 14']

Pais7 = Pais6[column_order]
#%% drop the rows
Pais8 = Pais7[~((Pais7['Country'] == Pais7['Unnamed: 1']) & (Pais7['Country'] != 'País / Código Arancelario'))]
Pais8.reset_index(drop=True, inplace=True)
#%% rename columns
Pais8 = Pais8.rename(
    columns={'Unnamed: 1': 'Code', 'Unnamed: 2': 'January', 'Unnamed: 3': 'February', 'Unnamed: 4': 'March',
             'Unnamed: 5': 'April', 'Unnamed: 6': 'May'
        , 'Unnamed: 7': 'June', 'Unnamed: 8': 'July', 'Unnamed: 9': 'August', 'Unnamed: 10': 'September',
             'Unnamed: 11': 'October', 'Unnamed: 12': 'November', 'Unnamed: 13': 'December'})
Pais8 = Pais8.drop(columns='Unnamed: 14')
#%% add a column date and fill it with the year
Pais8['Date']=''
current_date = None
for i in range(len(Pais8)):
    if "Enero - 2020" in str(Pais8.at[i, 'January']):
        current_date = "2020"
    elif "Enero - 2021" in str(Pais8.at[i, 'January']):
        current_date = "2021"
    elif "Enero - 2022" in str(Pais8.at[i, 'January']):
        current_date = "2022"
    elif "Enero - 2023" in str(Pais8.at[i, 'January']):
        current_date = "2023"
    elif "Enero - 2024" in str(Pais8.at[i, 'January']):
        current_date = "2024"

    Pais8.at[i, 'Date'] = current_date
Pais8 = Pais8.iloc[:, [0,1,14,2,3,4,5,6,7,8,9,10,11,12,13]]
#%% drop the row País / Código Arancelario
Pais8 = Pais8[Pais8['Country'] != 'País / Código Arancelario'].reset_index(drop=True)
#%% Melt the table
Pais9 = pd.melt(Pais8, id_vars=['Code', 'Country', 'Date'],
                  var_name='Month', value_name='Value')

# Convert 'Month' and 'Date' into a single datetime column
Pais9['Date'] = pd.to_datetime(Pais9['Month'] + ' ' + Pais9['Date'].astype(str), format='%B %Y')

# Drop the 'Month' column
Pais9 = Pais9.drop(columns=['Month'])

Pais9 = Pais9.sort_values(by='Date').reset_index(drop=True)
#%% verifying my work
filtered = Pais9[(Pais9['Date'] == '2024-03-01') &
                 (Pais9['Country'] == 'Italia') &
                 (Pais9['Code'] == '10051090')]

if not filtered.empty:
    x = filtered['Value'].iloc[0]
    print(x)
else:
    print('NO')
#%% export to excel file
Pais9.to_excel('Pais9.xlsx', index=False)