
import pyodbc
import pandas as pd

df = pd.read_excel('C:/Users/mrtar/Desktop/Database link/Excel/O-Risk stock.xlsm', engine='openpyxl')
df = df.fillna(value=0)

server = 'localhost'
database = 'postgres'
username = 'postgres'
password = '123456789'

cnxn = pyodbc.connect('DRIVER={Devart ODBC Driver for PostgreSQL};SEVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

for index, row in df.iterrows(): 

    try:
        cursor.execute('INSERT INTO demoone ("KEYID", "ITNBR", "ITDSC", "HOUSE", "BEGIN", "MOHTQ", "PROCS", "NOSUFFIX", "WH") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
            row['KEYID'], row['ITNBR'], row['ITDSC'], row['HOUSE'], row['BEGIN'], row['MOHTQ'], row['PROCS'], row['NOSUFFIX'], row['WH'])
        print("Insert Done")
    except:
        print("Error Insert Data")
        try:
            cursor.execute('UPDATE demoone SET "ITNBR" = ?, "ITDSC" = ?, "HOUSE" = ?, "BEGIN" = ?, "MOHTQ" = ?, "PROCS" = ?, "NOSUFFIX" = ?, "WH" = ? WHERE "KEYID" = ?',
                row['ITNBR'], row['ITDSC'], row['HOUSE'], row['BEGIN'], row['MOHTQ'], row['PROCS'], row['NOSUFFIX'], row['WH'], row['KEYID'])
            print("Update Done")
        except:
            print("Error Update Data")

cnxn.commit()
cursor.close()
