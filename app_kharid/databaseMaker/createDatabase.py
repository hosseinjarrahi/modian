import pyodbc
import sqlite3
import ORM
from time import sleep
# Connect to the Access database
path = r'C:\\Users\\a\Desktop\\safar\\database.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

conn = pyodbc.connect(db)
access_cursor = conn.cursor()

access_cursor.execute('SELECT * FROM Foroush_Detail')
rows = access_cursor.fetchall()

all = rows[::-1]

for row in all:
    print(row[0])
    try:
        ORM.create(
            KharidarName=row[18],
            KharidarLastNameSherkatName=row[19],
            KharidarNationalCode=row[21],
            HCKharidarTypeCode=row[13],
        )
    except:
        pass
    
access_cursor.close()
conn.close()
