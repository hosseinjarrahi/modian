import os
import pandas as pd

dir_path = './files'

dfs = []

for file_name in os.listdir(dir_path):
    if file_name.endswith('.xls'):
        df = pd.read_excel(os.path.join(dir_path, file_name), header=7)
        df = df[df.apply(lambda row: len(row.dropna()) >= 7, axis=1)]
        df = df.dropna(axis=1, how='all')
        dfs.append(df)
    if  file_name.endswith('.xlsx'):
        df = pd.read_excel(os.path.join(dir_path, file_name))
        
        dfs.append(df)

writer = pd.ExcelWriter('./output/output.xlsx', engine='xlsxwriter')

for i, df in enumerate(dfs):
    sheet_name = 'Sheet{}'.format(i+1)
    df.to_excel(writer, sheet_name=sheet_name, index=False)

writer._save()