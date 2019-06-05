import csv, pyodbc

# set up some constants
MDB = 'd:/Libraries/Documents/Bascula/DBTRUCK.mdb' #TODO: Change it to the other computer path
DRV = '{Microsoft Access Driver (*.mdb)}'
PWD = 'pw'

# connect to db
con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
cur = con.cursor()

# run a query and get the results 
SQL = 'SELECT observacion FROM HISTORIA WHERE CLIENTE=1 AND PLACAS=\'088-SU-9\';' # your query goes here [Selected Interenlace as CLIENTE]
rows = cur.execute(SQL).fetchall()
cur.close()
con.close()

print(rows)