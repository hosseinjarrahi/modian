import sys
import os
sys.path.append(os.getcwd() + r'/../databaseMaker')
import ORM
import pyodbc



path = os.getcwd() + r'./../TTMS.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

print("Connected To Database")
conn = pyodbc.connect(db)


companies = ORM.fetchAll()

i = 0
for company in companies:
    HCKharidarTypeCode = '' if company.HCKharidarTypeCode == None else company.HCKharidarTypeCode
    KharidarNationalCode = '' if company.KharidarNationalCode == None else company.KharidarNationalCode
    KharidarName = '' if company.KharidarName == None else company.KharidarName
    KharidarLastNameSherkatName = '' if company.KharidarLastNameSherkatName == None else company.KharidarLastNameSherkatName

    if KharidarName != '':
        sql_update = f"UPDATE Foroush_Detail SET HCKharidarTypeCode = '{HCKharidarTypeCode}' ,KharidarNationalCode = '{KharidarNationalCode}',KharidarLastNameSherkatName = '{KharidarLastNameSherkatName}'  WHERE KharidarName = '{KharidarName}'"
        conn.execute(sql_update)

    if KharidarLastNameSherkatName != '':
        sql_update = f"UPDATE Foroush_Detail SET HCKharidarTypeCode = '{HCKharidarTypeCode}' ,KharidarNationalCode = '{KharidarNationalCode}',KharidarName = '{KharidarName}'  WHERE KharidarLastNameSherkatName = '{KharidarLastNameSherkatName}'"
        conn.execute(sql_update)

    if KharidarNationalCode != '':
        sql_update = f"UPDATE Foroush_Detail SET HCKharidarTypeCode = '{HCKharidarTypeCode}' ,KharidarName = '{KharidarName}',KharidarLastNameSherkatName = '{KharidarLastNameSherkatName}'  WHERE KharidarNationalCode = '{KharidarNationalCode}'"
        conn.execute(sql_update)

    conn.commit()

    print(i)
    i = i + 1
