import pandas as pd
import pyodbc
import os
import pandas as pd
import time
import math
import numpy as np
start_time = time.time()
print(pyodbc.version)

sheetHeades = ['ردیف', 'شماره صورتحساب', 'پرداخت', 'مبلغ', 'ماليات و عوارض ماليات', 'شناسه خريدار', 'نام خريدار']

df_list = []
i = 0

excel_file = pd.ExcelFile('./output/output.xlsx')

sheet_names = excel_file.sheet_names

for sheet_name in sheet_names:
    df = pd.read_excel('./output/output.xlsx', sheet_name=sheet_name)
    df = df[sheetHeades]
    df['SheetName'] = sheet_name
    i = i + 1
    df_list.append(df)
    print(i)

new_df = pd.concat(df_list)
new_df['Access'] = range(1, len(new_df) + 1)
new_df['ردیف'] = range(1, len(new_df) + 1)


path = os.getcwd() + r'/../TTMS.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

print(db)
print("Connected To Database")

conn = pyodbc.connect(db)
cursor = conn.cursor()
i = 0
try:
    for index, row in new_df.iterrows():

        i += 1
        r = i
        n = str(0)
        k = 'حق ورود به پارکینگ'
        p = str(row[3])
        m = str(row[4])
        s = 0
        t = 0
        kh = row[6]
        cm = str(row[5]) 
        if len(cm) == 9:
            cm = '0'+cm
        if len(cm) == 8:
            cm = '00'+cm
        address = 'اسکله شرقی شهید رجایی'
        noe_kharid = 5
        SCode = 23
        Ccode = 2301000
        KT = 12
        maliat_m = str(0)
        if len(cm) > 10:
            noe_moshtari = 2
            sql = "INSERT INTO Foroush_Detail (Radif,KalaKhadamatName,Price,MaliatArzeshAfzoodeh,SayerAvarez,TakhfifPrice,KharidarLastNameSherkatName\
        ,KharidarNationalCode,KharidarAddress,HCKharidarTypeCode,HCKharidarType1Code,StateCode,CityCode,KalaType,MaliatMaksoore) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            record = (r, k, p, m, s, t, kh, cm, address, noe_moshtari,
                      noe_kharid, SCode, Ccode, KT, maliat_m)
        if len(cm) == 10:
            noe_moshtari = 1
            sql = "INSERT INTO Foroush_Detail (Radif,KalaKhadamatName,Price,MaliatArzeshAfzoodeh,SayerAvarez,TakhfifPrice,KharidarName,KharidarLastNameSherkatName\
        ,KharidarNationalCode,KharidarAddress,HCKharidarTypeCode,HCKharidarType1Code,StateCode,CityCode,KalaType,MaliatMaksoore) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            record = (r, k, p, m, s, t, kh, kh, cm, address,
                      noe_moshtari, noe_kharid, SCode, Ccode, KT, maliat_m)

        cursor.execute(sql, record)
        conn.commit()

        print(f"!...{r} Record inserted successfully...!")
    conn.close()
except pyodbc.Error as e:
    print("Error in Connection: ", e)
print(i)
print("--- %s seconds ---" % (time.time() - start_time))


new_df.to_excel('./output/output_TTMS.xlsx', index=False)
new_df['Access'] = range(1, len(new_df) + 1)
