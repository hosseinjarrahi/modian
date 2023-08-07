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


def stateConvert(item):
    cursor.execute('SELECT OstanCode FROM Zone WHERE Ostan = ?', item)
    row = cursor.fetchone()
    return row[0]


def cityConvert(item):
    cursor.execute('SELECT ShahrCode FROM Zone WHERE Shahr = ?', item)
    row = cursor.fetchone()
    return row[0]


def nationalCodeConvert(item):
    if len(item) == 9:
        item = '0'+item
    if len(item) == 8:
        item = '00'+item
    return item


def personType(item):
    if item == 'حقیقی':
        return 1
    return 2


mapKeys = {
    'KalaType': {'row': -1, 'convert': lambda item: 12},
    'KalaKhadamatName': {'row': -1, "convert": lambda item: 'حق ورود به پارکینگ'},
    'Price': {'row': 3},
    'MaliatArzeshAfzoodeh': {'row': 4},
    'SayerAvarez': {'row': -1, "convert": lambda item: 0},
    'TakhfifPrice': {'row': -1, "convert": lambda item: 0},
    'MaliatMaksoore': {'row': -1, "convert": lambda item: 0},
    'HCKharidarTypeCode': {'row': -1, "convert": personType},
    'KharidarAddress': {'row': -1, "convert": lambda item: 'اسکله شرقی شهید رجایی'},
    'KharidarName': {'row': 6},
    'KharidarLastNameSherkatName': {'row': 6},
    'KharidarNationalCode': {'row': 5, "convert": nationalCodeConvert},
    'HCKharidarType1Code': {'row': -1, 'convert': lambda item: 5},
    'StateCode': {'row': 11, 'convert': stateConvert},
    'CityCode': {'row': 12, 'convert': cityConvert},
}


def is_array(arr):
    return type(arr) == list


def log(*args):
    print('**********************')
    for arg in args:
        print(arg)
    print('**********************')

try:
    Radif = 0
    for index, row in new_df.iterrows():
        Radif += 1
        sql = "INSERT INTO Foroush_Detail (Radif,"
        sql_conc = ") VALUES (?,"
        values = []

        values.append(Radif)

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

            values.append(value)

        # Convert all int items to floats
        values = [float(x) if type(x) is int else x for x in values]

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

new_df.to_excel('./../output/kharid/output_TTMS.xlsx', index=False)

new_df['Access'] = range(1, len(new_df) + 1)
