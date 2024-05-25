import pyodbc
import pandas as pd
# import psycopg2

server = 'localhost'
database = 'postgres'
username = 'postgres'
password = '123456789'

# Read Excel to DataFrame
df = pd.read_excel('C:/Users/mrtar/Desktop/Database link/Excel/O-Risk stock.xlsm', engine='openpyxl')
df = df.fillna(value=0)


# Connect DB
cnxn = pyodbc.connect('DRIVER={Devart ODBC Driver for PostgreSQL};SEVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password) 
cursor = cnxn.cursor()

sql = "SELECT * FROM demoone"
df_postgresql = pd.read_sql_query(sql, cnxn)

for index, row in df_postgresql.iterrows():
    Key_postgres = {
    row['KEYID']: list([row['ITNBR'], row['ITDSC'], row['HOUSE'], row['BEGIN'], row['MOHTQ'], row['PROCS'], row['NOSUFFIX'], row['WH']])
}

print(Key_postgres)