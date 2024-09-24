import pandas as pd
import xlwings as xw

excel_app = xw.App(visible=False)
path = r"C:\Users\tsla\GitHub\smallthings\BOHR\UPDATE_Deep (Cluster 1) - Copy.xlsx"
data_sheet = "Surface areas overview"
calc_sheet = 'GACP Calculation'

print('reading excel...\n')
df = pd.read_excel(path, data_sheet, index_col=0)
print('creating sum column...\n')
df['sum'] = df['J-tubes (immersed, coated)'] + df['Boat landing (immersed, coated)'] + df['Jacket (immersed, coated)']

df_values = df[['Piles (embedded, bare steel)', 'Piles (immersed, bare steel)', 'sum']]

df_results = pd.DataFrame()
print('iterating through rows...\n')
for index, row in df_values.iterrows():
    print(f'calculating {index}...\n')
    doc_calc = xw.Book(path)
    sheet_calc = doc_calc.sheets[calc_sheet]
    sheet_calc['H21'].value = row.iloc[1]
    sheet_calc['H22'].value = row.iloc[2]
    sheet_calc['H25'].value = row.iloc[0]

    doc_calc.save(path)
    doc_calc.close()
    # just_open(path)

    df_result = pd.read_excel(path, calc_sheet, index_col=0, usecols = "B, H", nrows = 6, skiprows=43, header=None)
    df_result.rename(columns={df_result.columns[0]:index}, inplace=True)

    df_results = pd.concat([df_results, df_result], axis=1)

print(df_results)

df_results.to_excel('resultater.xlsx')

excel_app.quit()