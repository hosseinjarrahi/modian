import pyodbc
import ORM
import threading
# function that inserts a batch of records into ORM

def insert_batch(records):
    for row in records:
        try:
            ORM.create(
                KharidarName=row[18],
                KharidarLastNameSherkatName=row[19],
                KharidarNationalCode=row[21],
                HCKharidarTypeCode=row[13],
            )
        except:
            pass


# Connect to the Access database
path = r'C:\\Users\\a\Desktop\\safar\\database.mdb;'
db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + path

conn = pyodbc.connect(db)
access_cursor = conn.cursor()

access_cursor.execute('SELECT * FROM Foroush_Detail')
rows = access_cursor.fetchall()

# divide rows into 5 batches
num_threads = 5
batch_size = len(rows) // num_threads
batches = [rows[i:i+batch_size] for i in range(0, len(rows), batch_size)]

# create threads to insert batches of records into ORM
threads = []
for i in range(num_threads):
    t = threading.Thread(target=insert_batch, args=(batches[i],))
    threads.append(t)
    t.start()

# wait for all threads to finish
for t in threads:
    t.join()

access_cursor.close()
conn.close()
