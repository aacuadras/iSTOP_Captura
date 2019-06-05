from tkinter import *
from tkinter import ttk
import csv, pyodbc

MDB = 'c:/path/to/file.mdb' #TODO: Change it to the other computer path
DRV = '{Microsoft Access Driver (*.mdb)}'
PWD = 'pw'

con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
cur = con.cursor()

#Client query selector
SQL = 'SELECT NOMBRE FROM cliente;'
clientes = cur.execute(SQL).fetchall()


#Tkinter kit creator
top = Tk()
top.title("Seleccionar Cliente")
frame = Frame(top)
frame.grid(columnspan=2)
label2 = Label(frame, text="Cliente")
label2.grid()

##Start query list
list1 = Listbox(frame, selectmode=SINGLE, height=15)
counter = 0
for rows in clientes:
    counter =+ 1
    for el in rows:
        list1.insert(counter, el)
list1.grid(column=1, row=0)
##End query list

#Display element from the query where the client is the one selected
#and has the values for 'Fecha' and 'Placas'
def displayQuery():
    return 1

#TODO: Save "Folio" and "Cliente/Destino" to mdb file  
def save_client():
    return 1

'''Creates new windows with information for the specific client'''
def tablaCliente(cl, value):
    tableWindow = Toplevel(top)
    tableWindow.title("Cliente: " + value)
    tableWindow.geometry("800x390")
    
    label2 = Label(tableWindow, text="iSTOP")
    label2.place(x=5, y=5)
    chat = ttk.Treeview(tableWindow, height="15", columns=("Nick","Mensaje","Hora"), selectmode="extended")
    chat.place(x=5, y=60)
    chat.heading('#1', text='Fecha', anchor=W)
    chat.heading('#2', text='Placas', anchor=W)
    chat.heading('#3', text='Alpha', anchor=W)
    chat.column('#1', stretch=NO, minwidth=0, width=130)
    chat.column('#2', stretch=NO, minwidth=0, width=65)
    chat.column('#3', stretch=NO, minwidth=0, width=65)
    chat.column('#0', stretch=NO, minwidth=0, width=0)
    chat.bind("<Double-1>", displayQuery)

    SQL = "SELECT * FROM HISTORIA WHERE CLIENTE=" + str(cl) + ";"
    query = cur.execute(SQL).fetchall()
    for row in query:
        chat.insert('',1, text = row[1], values =(row[1],row[3], row[32]))
    
    label3 = Label(tableWindow, text='Fecha')
    label3.place(x=350, y=80)
    info1 = Entry(tableWindow, state=DISABLED)
    info1.place(x=390, y=80)
    label4 = Label(tableWindow, text="Placas")
    label4.place(x=550, y=80)
    info2 = Entry(tableWindow, state=DISABLED)
    info2.place(x=590, y=80)
    label5 = Label(tableWindow, text="Economico")
    label5.place(x=350, y=120)
    info3 = Entry(tableWindow, state=DISABLED)
    info3.place(x=420, y=120)
    label6 = Label(tableWindow, text="Alpha")
    label6.place(x=550, y=120)
    info4 = Entry(tableWindow, state=DISABLED)
    info4.place(x=590, y=120)
    #TODO: Lower this two labels/entries
    label7 = Label(tableWindow, text="Folio")
    label7.place(x=350, y=160)
    info5 = Entry(tableWindow)
    info5.place(x=390, y=160)
    label8 = Label(tableWindow, text="Cliente/Destino")
    label8.place(x=550, y=160)
    info6 = Entry(tableWindow)
    info6.place(x=650, y=160)

    buttonSave = Button(tableWindow, command=save_client, text="Guardar")
    buttonSave.place(x=450, y=300)

    buttonClear =Button(tableWindow, text="Borrar Campos")
    buttonClear.place(x=550, y=300)
    
    

def selecc():
    value = str(list1.get(ACTIVE))
    SQL = "SELECT numero FROM cliente WHERE nombre='" + value + "';"
    s = cur.execute(SQL).fetchall()
    print("Seleccionando cliente: " + value)
    tablaCliente(s[0][0], value)

#TODO: Create form to add a cliente *****Probably won't be used since the database already stores the clients****** 
def create_window():
    ag_cliente = Toplevel(top)
button1 = Button(frame, text="Seleccionar", command=selecc)
button1.grid(row=1)
button2 = Button(frame, text="Crear Cliente", command=create_window)
button2.grid(row=1, column=1)

top.mainloop()
cur.close()
con.close()
