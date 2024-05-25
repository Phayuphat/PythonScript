import pyodbc
import pandas as pd
import sys

server = 'localhost'
database = 'postgres'
username = 'postgres'
password = '123456789'

cn_DB = pyodbc.connect('DRIVER={Devart ODBC Driver for PostgreSQL};SEVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cn_DB.cursor()

sql = "SELECT * FROM demoone"
df_DB = pd.read_sql_query(sql, cn_DB)

df_excel = pd.read_excel('C:/Users/mrtar/Desktop/Database link/Excel/O-Risk stock.xlsm', engine='openpyxl')
df_excel = df_excel.fillna(value=0)  #Replace NaN with zero

df_diff = pd.concat([df_excel, df_DB], keys=['x', 'y']).drop_duplicates(keep=False)

#check df_diff not empty
if df_diff.empty:
    sys.exit()
else:
    df_change = df_diff.loc['x']

for index, row in df_change.iterrows():
    try:
        if (df_change['KEYID'].isin(df_DB['KEYID'])).any(): 
            cursor.execute('UPDATE demoone SET "ITNBR" = ?, "ITDSC" = ?, "HOUSE" = ?, "BEGIN" = ?, "MOHTQ" = ?, "PROCS" = ?, "NOSUFFIX" = ?, "WH" = ? WHERE "KEYID" = ?',
                row['ITNBR'], row['ITDSC'], row['HOUSE'], row['BEGIN'], row['MOHTQ'], row['PROCS'], row['NOSUFFIX'], row['WH'], row['KEYID'])
            print("Update Done")
        else:
            cursor.execute('INSERT INTO demoone ("KEYID", "ITNBR", "ITDSC", "HOUSE", "BEGIN", "MOHTQ", "PROCS", "NOSUFFIX", "WH") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
                row['KEYID'], row['ITNBR'], row['ITDSC'], row['HOUSE'], row['BEGIN'], row['MOHTQ'], row['PROCS'], row['NOSUFFIX'], row['WH'])
            print("Insert Done")
    except Exception as e:
        print(f"Error : {e}")

cn_DB.commit()
cursor.close()
cn_DB.close()