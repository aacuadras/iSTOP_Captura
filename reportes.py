from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import csv, pyodbc

MDB = 'c:/SetupPV/Data/DBTRUCK.mdb' 
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



'''Creates new windows with information for the specific client'''
def tablaCliente(cl, value):
    
    def displayQuery(a):
        curItem = chat.focus()
        curTruck = chat.item(curItem)

        info1.delete(0,END)
        info1.insert(0, str(curTruck['values'][0]))
        info2.delete(0,END)
        info2.insert(0, str(curTruck['values'][1]))
        SQL = "SELECT embarque FROM HISTORIA WHERE PLACAS='" + str(curTruck['values'][1]) + "' AND observacion='" + str(curTruck['values'][2]) +"' AND FECHA1='" + str(curTruck['values'][0]) + "';"
        info3.delete(0, END)
        econ = cur.execute(SQL).fetchall()
        info3.insert(0, str(econ[0][0]))
        info4.delete(0,END)
        info4.insert(0, str(curTruck['values'][2]))
        SQL = "SELECT FOLIOGEN FROM HISTORIA WHERE PLACAS='" + str(curTruck['values'][1]) + "' AND observacion='" + str(curTruck['values'][2]) +"' AND FECHA1='" + str(curTruck['values'][0]) + "';"
        econ = cur.execute(SQL).fetchall()
        info5.delete(0,END)
        if(econ[0][0] != None):
            info5.insert(0, str(econ[0][0]))
        SQL = "SELECT PROVEEDOR FROM HISTORIA WHERE PLACAS='" + str(curTruck['values'][1]) + "' AND observacion='" + str(curTruck['values'][2]) +"' AND FECHA1='" + str(curTruck['values'][0]) + "';"
        econ = cur.execute(SQL).fetchall()
        info6.delete(0,END)
        if(econ[0][0] != None):
            info6.insert(0, str(econ[0][0]))
        SQL = "SELECT guia FROM HISTORIA WHERE PLACAS='" + str(curTruck['values'][1]) + "' AND observacion='" + str(curTruck['values'][2]) +"' AND FECHA1='" + str(curTruck['values'][0]) + "';"
        econ = cur.execute(SQL).fetchall()
        info7.delete(0,END)
        if(econ[0][0] != None):
            info7.insert(0, str(econ[0][0]))
        SQL= "SELECT * FROM PERMISOS WHERE FOLIOBAS='" + str(info7.get()) + "';"
        per = cur.execute(SQL).fetchall()
        if(per == []):
            permit1.deselect()
        else:
            permit1.select()

    
    def save_client():
        #If all entries have data, save the cliente. Otherwise display an error message
        if(info1.get() != "" and info2.get() != "" and info4.get() != "" and info5.get() != "" and info6.get() != "" and info7.get() != ""):
            SQL = "UPDATE HISTORIA SET guia='" + str(info7.get()) + "', PROVEEDOR='" + str(info6.get()) + "', FOLIOGEN='" + str(info5.get()) + "' WHERE FECHA1='" + str(info1.get()) + "' AND PLACAS='" + str(info2.get()) + "' AND observacion='" + str(info4.get()) + "';"
            cur.execute(SQL)
            cur.commit()
            messagebox.showinfo("Creado", "Folio y cliente creado")
        else:
            messagebox.showerror("Error", "Llenar todos los campos")

    def del_entries():
        info1.delete(0, END)
        info2.delete(0, END)
        info3.delete(0, END)
        info4.delete(0, END)
        info5.delete(0, END)
        info6.delete(0, END)
        info7.delete(0, END)
        

    def filter_date():
        if(entry1.get() != '' and entry2.get() != ''):
            SQL = "SELECT * FROM HISTORIA WHERE FECHA1 BETWEEN '" + str(entry1.get()) + "' AND '" + str(entry2.get()) + "' AND CLIENTE=" + str(cl) + ";"
            query = cur.execute(SQL).fetchall()
            for i in chat.get_children():
                chat.delete(i)

            for row in query:
                chat.insert('',1, text = row[1], values =(row[1],row[3], row[32]))
        else:
            for i in chat.get_children():
                chat.delete(i)
            SQL = "SELECT * FROM HISTORIA WHERE CLIENTE=" + str(cl) + ";"
            query = cur.execute(SQL).fetchall()
            for row in query:
                chat.insert('',1, text = row[1], values =(row[1],row[3], row[32]))
        
    
    def generate_report():
        def create_csv():
            SQL = "SELECT FECHA1, PROVEEDOR, FOLIO, embarque, observacion, PLACAS, guia, PRODUCTO FROM HISTORIA WHERE FECHA1='" + str(date.get()) + "' AND CLIENTE=" + str(cl) + ";"
            rows = cur.execute(SQL).fetchall()
            if(cl != 4):
                totalItems = 0
                for i in rows:
                    SQL = "SELECT costo from cliente WHERE nombre='" + value + "';"
                    curCost = cur.execute(SQL).fetchall()
                    i[7] = curCost[0][0]
                    totalItems += 1
                    '''
                    i[7] = "$280.00"
                    totalItems += 1
                    '''
            else:
                for i in rows:
                    curCost = 0
                    i[7] = "CANCELADO"
                
            with open('REPORTE_DIARIO.csv', 'w', newline='') as fou:
                csv_writer = csv.writer(fou) # default field-delimiter is ","
                csv_writer.writerow(['FECHA', 'CLIENTE/DESTINO', 'FACTURA/FOLIO', 'ECONOMICO', 'ALPHA', 'PLACAS', 'FOLIO BASCULA', '$'])
                csv_writer.writerows(rows)
                if(cl != 4):
                    csv_writer.writerow(['','','','','','TOTAL',str(totalItems), "$" + str(curCost[0][0] * totalItems) + ".00"])
            messagebox.showinfo("Creado", "Reporte Creado")
            r.destroy()


        r = Toplevel(tableWindow)
        r.title('Crear Reporte')
        label11 = Label(r, text="Ingresar Fecha [DD/MM/AAAA]:")
        label11.pack()
        date = Entry(r)
        date.pack()
        submit = Button(r, text="Seleccionar", command=create_csv)
        submit.pack()


    def show_permit():
        def save_permit():
            SQL= "SELECT * FROM PERMISOS WHERE FOLIOBAS='" + str(info7.get()) + "';"
            test = cur.execute(SQL).fetchall()
            #If the id is not found in the table
            if(test == []):
                SQL = "INSERT INTO PERMISOS (FOLIOBAS, FOLIO, NOMBRE, TIPO) VALUES ('" + str(info7.get()) + "', '" + str(permitIDInfo.get()) \
                    + "', '" + str(permitNameInfo.get()) + "', '" + str(permitTypeInfo.get()) + "');" 
                cur.execute(SQL)  
                cur.commit()
            else:
                SQL = "UPDATE PERMISOS SET FOLIO='" + str(permitIDInfo.get()) + "', NOMBRE='" + str(permitNameInfo.get()) + "', TIPO='" + str(permitTypeInfo.get()) + "' WHERE FOLIOBAS='" + str(info7.get()) + "';"
                cur.execute(SQL)
                cur.commit()

        if(permitCheck.get() == 1 and info7.get() != ''):
            P = Toplevel(tableWindow)
            P.title('Informacion de Permiso')
            P.geometry("300x100")
            ti = Label(P)
            ti.grid(columnspan=4)
            permitID = Label(P, text='Número: ')
            permitID.grid(column=0, row=0)
            permitIDInfo = Entry(P)
            permitIDInfo.grid(column=1, row=0)
            permitName = Label(P, text='Nombre: ')
            permitName.grid(column=0, row=1)
            permitNameInfo = Entry(P, width=35)
            permitNameInfo.grid(column=1, row=1)
            permitType = Label(P, text="Tipo: ")
            permitType.grid(column=0, row=2)
            permitTypeInfo = Entry(P, width=25)
            permitTypeInfo.grid(column=1, row=2)
            saveBtn = Button(P, text="Guardar", command=save_permit)
            saveBtn.grid(column=1)

            SQL = "SELECT * FROM PERMISOS WHERE FOLIOBAS='" + str(info7.get()) + "';"
            permit = cur.execute(SQL).fetchall()
            permitIDInfo.delete(0, END)
            permitNameInfo.delete(0, END)
            permitTypeInfo.delete(0, END)
            if(permit != []):
                permitIDInfo.insert(0, permit[0][1])
                permitNameInfo.insert(0, permit[0][2])
                permitTypeInfo.insert(0, permit[0][3])
        else:
            messagebox.showerror('Error', 'Permiso no habilitado')

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

    label9 = Label(tableWindow, text="Desde [DD/MM/AAAA]:")
    label9.place(x=5, y=30)
    entry1 = Entry(tableWindow)
    entry1.place(x=140, y=30)
    label10 = Label(tableWindow, text="Hasta: [DD/MM/AAAA]:")
    label10.place(x=240, y=30)
    entry2 = Entry(tableWindow)
    entry2.place(x=375, y=30)
    buttonDate = Button(tableWindow, command=filter_date, text="Filtrar por Fecha")
    buttonDate.place(x=500, y=30)
    
    label3 = Label(tableWindow, text='Fecha')
    label3.place(x=350, y=80)
    info1 = Entry(tableWindow)
    info1.place(x=390, y=80)
    label4 = Label(tableWindow, text="Placas")
    label4.place(x=550, y=80)
    info2 = Entry(tableWindow)
    info2.place(x=590, y=80)
    label5 = Label(tableWindow, text="Economico")
    label5.place(x=350, y=120)
    info3 = Entry(tableWindow)
    info3.place(x=420, y=120)
    label6 = Label(tableWindow, text="Alpha")
    label6.place(x=550, y=120)
    info4 = Entry(tableWindow)
    info4.place(x=590, y=120)
    label7 = Label(tableWindow, text="Folio")
    label7.place(x=350, y=160)
    info5 = Entry(tableWindow)
    info5.place(x=390, y=160)
    label8 = Label(tableWindow, text="Cliente/Destino")
    label8.place(x=550, y=160)
    info6 = Entry(tableWindow)
    info6.place(x=650, y=160)
    label13 = Label(tableWindow, text="Folio Bascula")
    label13.place(x=400, y=200)
    info7 = Entry(tableWindow)
    info7.place(x=480, y=200)
    permitCheck = IntVar()
    permit1 = Checkbutton(tableWindow, text="Permiso", variable=permitCheck)
    permit1.place(x=480, y=230)

    buttonSave = Button(tableWindow, command=save_client, text="Guardar")
    buttonSave.place(x=350, y=300)

    buttonClear =Button(tableWindow, text="Borrar Campos", command=del_entries)
    buttonClear.place(x=430, y=300)

    buttonReport = Button(tableWindow, text="Generar Reporte", command=generate_report)
    buttonReport.place(x=550, y=300)

    buttonPermit = Button(tableWindow, text="Info. Permiso", command=show_permit)
    buttonPermit.place(x=670, y=300)

    
    
#Display element from the query where the client is the one selected
#and has the values for 'Fecha' and 'Placas'

def selecc():
    value = str(list1.get(ACTIVE))
    SQL = "SELECT numero FROM cliente WHERE nombre='" + value + "';"
    s = cur.execute(SQL).fetchall()
    print("Seleccionando cliente: " + value)
    tablaCliente(s[0][0], value)


def create_cancel():
    def cancel_folio():
        SQL = "UPDATE HISTORIA SET PROVEEDOR='CANCELADO', FOLIO=00, embarque='CANCELADO', observacion='CANCELADO', PLACAS='CANCELADO', CLIENTE=4 WHERE guia='" + str(entry4.get()) + "';"
        cur.execute(SQL)
        cur.commit()
        messagebox.showinfo("Cancelado", "Folio Cancelado")
        cancelar.destroy()

    cancelar = Toplevel(top)
    cancelar.title('Cancelar Folio')
    label12 = Label(cancelar, text="Ingresar Folio Bascula a Cancelar")
    label12.pack()
    entry4 = Entry(cancelar)
    entry4.pack()
    submit_cancel = Button(cancelar, text="Cancelar", command=cancel_folio)
    submit_cancel.pack()

def change_charge():
    def save_cost():
        SQL = "UPDATE cliente SET costo=" + entryCost.get() + " WHERE nombre='" + activeClient + "';"
        cur.execute(SQL)
        cur.commit()

        messagebox.showinfo("Guardado", "Nuevo costo para " + activeClient + ": $" + entryCost.get())
        c.destroy()

    c = Toplevel(top)
    c.title("Editar Costo")

    activeClient = str(list1.get(ACTIVE))
    SQL = "SELECT costo FROM cliente WHERE nombre='" + activeClient + "';"
    cost = cur.execute(SQL).fetchall()

    label14 = Label(c, text='Cambiar costo para cliente ' + activeClient)
    label14.pack()
    #Fetching cost from database
    entryCost = Entry(c)
    entryCost.pack()
    entryCost.delete(0,END)
    entryCost.insert(0,cost[0][0])
    button4 = Button(c, text='Guardar', command=save_cost)
    button4.pack()



button1 = Button(frame, text="Seleccionar", command=selecc)
button1.grid(row=1)
button2 = Button(frame, text="Cancelar Folio", command=create_cancel)
button2.grid(row=1, column=1)
button3 = Button(frame, text="Cambiar Precio", command=change_charge)
button3.grid(row=2,column=1)

top.mainloop()
cur.close()
con.close()
