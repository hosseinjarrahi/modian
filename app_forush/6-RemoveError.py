import pandas as pd
import pyodbc
import os

path = os.getcwd() + './../TTMS.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

print("Connected To Database")
conn = pyodbc.connect(db)

sql_query = 'SELECT Radif, KharidarNationalCode FROM Foroush_Detail'

excel_file = './error.xlsx'
df = pd.read_excel(excel_file)
headers = list(df.columns)
headers = headers[3:]

map = {
    'شخصیت': 'HCKharidarTypeCode',
    'شناسه خریدار': 'KharidarNationalCode',
    'نام خریدار': 'KharidarName',
    'نام شرکت/نام خانوادگی خریدار': 'KharidarLastNameSherkatName',
}

for i, row in df.iterrows():
    row_number = row['ردیف']
    sql_update = f"UPDATE Foroush_Detail SET "

    for header in headers:
        value = row[header]
        
        if value == '*':
            continue

        if value == None:
            value = ''

        if value == 'nan':
            value = ''

        if value == '.':
            value = ''

        if value == 'حقوقی':
            value = 2

        if value == 'حقیقی':
            value = 1

        engHeader = map[header]
        if engHeader in ['KharidarLastNameSherkatName','KharidarName','KharidarNationalCode']:
            sql_update += f"{map[header]} = '{value}',"
        else:
            sql_update += f"{map[header]} = {value},"

    sql_update = sql_update[:-1]
    sql_update += f" WHERE Radif = {row_number}"
    print(sql_update)
    conn.execute(sql_update)
    conn.commit()

conn.close()
