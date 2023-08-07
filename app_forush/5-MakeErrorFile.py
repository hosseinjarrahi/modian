import pandas as pd
import pyodbc
import os

def isNan(value):
    try:
        if value == None or math.isnan(value):
            return True
    except:
        pass
    return False


path = os.getcwd() + r'./../TTMS.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

print("Connected To Database")
conn = pyodbc.connect(db)
cursor = conn.cursor()

excel_file = './error.xlsx'
df = pd.read_excel(excel_file)
headers = list(df.columns)

map = {
    'شخصیت': 'HCKharidarTypeCode',
    'شناسه خریدار': 'KharidarNationalCode',
    'نام خریدار': 'KharidarName',
    'نام شرکت/نام خانوادگی خریدار': 'KharidarLastNameSherkatName',
}

for i, row in df.iterrows():
    row_number = row['ردیف']
    sql = f"SELECT HCKharidarTypeCode,KharidarNationalCode,KharidarName,KharidarLastNameSherkatName FROM Foroush_Detail WHERE Radif = {row_number}"

    cursor.execute(sql)
    row = cursor.fetchone()
    data = {
        'HCKharidarTypeCode': row[0],
        'KharidarNationalCode': row[1],
        'KharidarName': row[2],
        'KharidarLastNameSherkatName': row[3],
    }

    for key, value in map.items():
        if isNan(data[value]):
            df.at[i,key] = '.'
        else:
            df.at[i,key] = data[value]

conn.close()
writer = pd.ExcelWriter('./error_template.xlsx')
df.to_excel(writer, index=False)
writer._save()
