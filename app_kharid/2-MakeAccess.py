import pandas as pd
import pyodbc
import os
import pandas as pd
import time
import math
import numpy as np
import array

start_time = time.time()

df_list = []
i = 0
pathExcelFile = './output/output.xlsx'

excel_file = pd.ExcelFile(pathExcelFile)

sheet_names = excel_file.sheet_names

for sheet_name in sheet_names:
    df = pd.read_excel(pathExcelFile, sheet_name=sheet_name)
    df['SheetName'] = sheet_name
    i = i + 1
    df_list.append(df)
    print(i)

# Concatenate the selected DataFrames into a new DataFrame
new_df = pd.concat(df_list)
new_df['Access'] = range(1, len(new_df) + 1)
new_df['ردیف'] = range(1, len(new_df) + 1)

path = os.getcwd() + r'/../TTMS.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

print(path)
print("Connected To Database")

conn = pyodbc.connect(db)
cursor = conn.cursor()
i = 0


def log(*args):
    print('**********************')
    for arg in args:
        print(arg)
    print('**********************')


def stateConvert(item):
    cursor.execute('SELECT OstanCode FROM Zone WHERE Ostan = ?', item)
    row = cursor.fetchone()
    return row[0]


def cityConvert(item):
    cursor.execute('SELECT ShahrCode FROM Zone WHERE Shahr = ?', item)
    row = cursor.fetchone()
    return row[0]


def personType(item):
    if item == 'حقیقی':
        return 1
    return 2


def typeNational(item):
    if item == '':
        return ''
    
    item = str(int(item))
    if len(item) == 9:
        item = '0'+item
    if len(item) == 8:
        item = '00'+item
        
    return item


mapKeys = {
    'KalaType': {'row': -1, 'convert': lambda item: 12},
    'KalaKhadamatName': {'row': 1},
    'Price': {'row': 15, "type": lambda item: float(item)},
    'MaliatArzeshAfzoodeh': {'row': 18, "type": lambda item: round(item)},
    'AvarezArzeshAfzoodeh': {'row': -1, 'convert': lambda item: ''},
    'SayerAvarez': {'row': 19},
    'TakhfifPrice': {'row': 16},
    'MaliatMaksoore': {'row': 17},
    'HCForoushandeTypeCode': {'row': 3, 'convert': personType},
    'ForoushandeAddress': {'row': 13},
    'ForoushandeName': {'row': 9},
    'ForoushandeLastNameSherkatName': {'row': [8, 5]},
    'ForoushandeEconomicNO': {'row': 6, "type": lambda item: str(item)},
    'ForoushandeNationalCode': {'row': [7, 10], "type": typeNational},
    'HCForoushandeType1Code': {'row': -1, 'convert': lambda item: 5},
    'StateCode': {'row': 11, 'convert': stateConvert, "type": lambda item: int(item)},
    'CityCode': {'row': 12, 'convert': cityConvert, "type": lambda item: int(item)},
}


def is_array(arr):
    return type(arr) == list


try:
    for index, row in new_df.iterrows():
        i += 1
        Radif = i
        sql = "INSERT INTO Kharid_Detail (Radif,Sarjam,IsHagholAmalKari,BargashtType,"
        sql_conc = ") VALUES (?,?,?,?,"
        values = []

        values.append(Radif)
        values.append(0)
        values.append(0)
        values.append(0)

        for key in mapKeys:
            value = ''
            sql += key + ","
            sql_conc += "?,"
            handler = mapKeys[key]
            if 'convert' in handler:
                value = handler['convert'](row[handler['row']])
            elif is_array(handler['row']):
                for item in handler['row']:
                    if row[item]:
                        try:
                            if math.isnan(row[item]):
                                continue
                        except:
                            pass
                        value = row[item]
            else:
                value = row[handler['row']]

            isNan = False
            try:
                if math.isnan(value):
                    value = ''
            except:
                pass

            if "type" in handler:
                value = handler['type'](value)

            values.append(value)

        sql = sql[:-1]
        sql_conc = sql_conc[:-1]
        sql += sql_conc + ")"

        cursor.execute(sql, values)
        conn.commit()

        print(f"!...{i} Record inserted successfully...!")
    conn.close()
except pyodbc.Error as e:
    print("Error in Connection: ", e)

print(i)
print("--- %s seconds ---" % (time.time() - start_time))

new_df.to_excel('./output/output_TTMS.xlsx', index=False)

new_df['Access'] = range(1, len(new_df) + 1)
