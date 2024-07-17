import pandas as pd

excel_file = 'input.xlsx'
df_sheet_multi = pd.read_excel(excel_file, sheet_name=['入力項目定義','出力項目定義'], header=6)


print(df_sheet_multi)