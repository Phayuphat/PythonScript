import pyodbc
import pandas as pd

df = pd.read_excel(f'//10.122.77.1/M_PEM-IoT/1007017_PANKANOK/Inventory/O-Risk stock.xlsm', engine = 'openpyxl')
df = df.fillna(value=0)
print(df)

server = '10.122.77.1'
database = 'inventory'
username = 'densoinfo'
password = 'densoinfo'

cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

# .iterrows >> go to data row by row
# row save data in row
for index, row in df.iterrows(): 
     try:
          
          cursor.execute("INSERT INTO risk_stock (KEYID, ITNBR, ITDSC, HOUSE, BEGINQTY, MOHTQ, PROCS, NOSUFFIX, WH) values(?,?,?,?,?,?,?,?,?)", row.KEYID, row.ITNBR, row.ITDSC, row.HOUSE, row.BEGINQTY, row.MOHTQ, row.PROCS, row.NOSUFFIX, row.WH)
          print("Insert Risk_stock") 
     except:
          try:
               cursor.execute("UPDATE risk_stock SET ITNBR = ?, ITDSC = ?, HOUSE = ?, BEGINQTY = ?, MOHTQ = ?, PROCS = ?, NOSUFFIX = ?, WH = ? WHERE KEYID = ?;", row.ITNBR, row.ITDSC, row.HOUSE, row.BEGINQTY, row.MOHTQ, row.PROCS, row.NOSUFFIX, row.WH, row.KEYID)
               print("Update Risk_stock")
          except:
               pass
print("Done")
cnxn.commit()
cursor.close()