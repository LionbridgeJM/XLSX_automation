

from concurrent.futures import process
from fileinput import filename
from posixpath import basename
import threading
from tkinter import *
from tkinter.tix import NoteBook
from tkinter import filedialog 
from tkinter import ttk 
import tkinter.messagebox
from tkinter import messagebox
#from excelClass import *
from excelClass import *
import time


import os

from threading import Thread


                                    
def browseFiles():

    filename = filedialog.askopenfilename(initialdir = "/",
                                        
                                          title = "Select a File",
                                          filetypes = (("Text files",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*"))) 
      #Esto es para que sea accesible fuera de la funcion                                              
    browseFiles.filename=filename                                                       
    # Change label contents
    basename=os.path.basename(filename)
    
   # print(basename)
    label_file_explorer.configure(text="File Opened: "+basename )
    
def opcionesMenu(eventObject):
    #obtenemos el valor de la primera lista
    opt = cbbPrincipal.get()


    #obtenemos el indice del valor escogido de la lista
    opcionesMenu.index = opcionesPrincipales.index(opt)
    
    #llenamos el segundo combobox
   
    cbbSegundario.config(values=opcionesSegundarias[ opcionesMenu.index])
   


def codigo(): 
        
        threads = [None]
        results = [None] 

     
        if label_file_explorer['text'] == "File Explorer:"  or label_file_explorer['text']=="File Opened: ":
            tkinter.messagebox.showinfo("info", "Select a File:")
        #print(tabControl.index(tab2))
        else:
            ###################
            #seleccion de idex del tab
            tabName = tabControl.index(tab_id='current')
            tab = tabName

            path=browseFiles.filename   

            if  tab == 0:
                if cbbTMSSegundario.get() == "Choose an Option":
                           tkinter.messagebox.showinfo("info", "Select an Option:")
                else:
                    optTMSproceso = cbbTMSSegundario.get()       
                    indexTMSsegundo = opcionesTMSSegundarias[opcionesTMSMenu.index].index(optTMSproceso)
            
                    if  opcionesTMSMenu.index==0:
                          
                            proceso=ExcelProcesses(path, indexTMSsegundo)
                            t1= threading.Thread(target= proceso.TMSftbt)
                            t1.start()
                            t1.join()
                            #retorno=proceso.mensajeretornado
                            #print(proceso.mensajeretornado)
                            tkinter.messagebox.showinfo("info",proceso.mensajeretornado)
                            #proceso.TMSftbt()
                         
                            cbbTMSSegundario.set("Choose an Option")
                           
                    else:
                            proceso=ExcelProcesses(path, indexTMSsegundo)
                            t2= threading.Thread(target= proceso.TMSmigration)
                            t2.start()
                            t2.join()
                            tkinter.messagebox.showinfo("info",proceso.mensajeretornado)
                            cbbTMSSegundario.set("Choose an Option")
            else:
                if cbbSegundario.get() == "Choose an Option":
                           tkinter.messagebox.showinfo("info", "Select an Option:")
                else:
                
                    optProceso = cbbSegundario.get()  
                    indexSegundo = opcionesSegundarias[opcionesMenu.index].index(optProceso)
                    if  opcionesMenu.index==0:
                        
                        
                            proceso=ExcelProcesses(path, indexSegundo)
                            t3= threading.Thread(target= proceso.ftbt)
                            t3.start()
                            t3.join()
                            tkinter.messagebox.showinfo("info",proceso.mensajeretornado)
                                                   
                            
                    else:

                            
                            proceso=ExcelProcesses(path, indexSegundo)
                            t4=threading.Thread(target=proceso.migration)
                            t4.start()
                            t4.join()
                            tkinter.messagebox.showinfo("info", proceso.mensajeretornado)
                            cbbSegundario.set("Choose an Option")
        
        
# Create the root window
window = Tk()      

# Set window title
window.title('PEUL')  

# Set window size
window.geometry("700x300")
window.resizable(False, False)

#Set window background color
window.config(background = "white")

# Frame 1
frame1 = Frame(window,bg="black",width=500,height=300)
frame1.pack(pady=0,padx=0)

# Create a File Explorer label
label_file_explorer = Label(frame1,
                            text = "File Explorer:",
                            width = 100, height = 4,
                            fg = "blue")  
      
button_explore = Button(frame1,
                        text = "Browse Files",
                        command = browseFiles)  

# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
label_file_explorer.grid(column = 1, row = 1)  
button_explore.grid(column = 1, row = 2)

# Frame 2
frame2 = Frame(window,bg="grey",width=500,height=300)
frame2.pack(pady=0,padx=0)
tabControl = ttk.Notebook(frame2)

style = ttk.Style(frame2)
style.configure('TNotebook.Tab',width=frame2.winfo_screenwidth(),height=600)

tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl) 
tabControl.add(tab1, text ='TMS')
tabControl.add(tab2, text ='OO TMS')
tabControl.pack(expand = 1, fill ="both")


#opciones 1
opcionesTMSPrincipales = ["FTBT", "Migration"]
#opciones 2
opcionesTMSSegundarias = [
        ["0 )Default TMS FTBT prep"],
        ["1 )Default TMS Migration prep",]
    ]

#dibujamos el primer combobox, llenamos sus valores
cbbTMSPrincipal = ttk.Combobox(tab1,state = "readonly", width=37, values=(opcionesTMSPrincipales))
cbbTMSPrincipal.current(0)

cbbTMSPrincipal.pack(pady=0,padx=0)

#creamos la funcion para el llenado del combobox 2

def opcionesTMSMenu(eventObject):

    #obtenemos el valor de la primera lista
    TMSopt = cbbTMSPrincipal.get()   
    #obtenemos el indice del valor escogido de la lista
    opcionesTMSMenu.index = opcionesTMSPrincipales.index(TMSopt)
    #llenamos el segundo combobox
    cbbTMSSegundario.config(values=opcionesTMSSegundarias[opcionesTMSMenu.index])


    

cbbTMSSegundario = ttk.Combobox(tab1,state = "readonly", width=37,)

cbbTMSSegundario.pack(pady=1,padx=1)

#llenamos el segundo combobox, aquí se hace el llamado a la funcion
cbbTMSSegundario.bind('<Button>', opcionesTMSMenu)
cbbTMSSegundario.set("Choose an Option")


#opciones 1
opcionesPrincipales = ["FTBT", "Migration"]
#opciones 2
opcionesSegundarias = [
        ["0 )Default FTBT prep","1)Default FTBT prep/Split Tabs","2 )Only FTBT colored lines","3)Only FTBT colored lines/Split Tabs"],
        ["0 )Default Migration prep","1 )Default Migration prep/Split Tabs","2 )Only Migration colored lines","3 )Only Migration colored lines/Split Tabs"]
    ]    

#dibujamos el primer combobox, llenamos sus valores
cbbPrincipal = ttk.Combobox(tab2, width=37,state = "readonly", values=(opcionesPrincipales))
cbbPrincipal.current(0)


#posicion del combobox
cbbPrincipal.pack(pady=0,padx=0)
#creamos la funcion para el llenado del combobox 2


cbbSegundario = ttk.Combobox(tab2,state = "readonly", width=37)
cbbSegundario.pack(pady=1,padx=1)

#llenamos el segundo combobox, aquí se hace el llamado a la funcion
cbbSegundario.bind('<Button>', opcionesMenu)
cbbSegundario.set("Choose an Option")

frame3 = Frame(window,bg="blue",width=500,height=300)
frame3.pack(pady=0,padx=0)   

button_run = Button(frame3,
                     text = "Prepare!",
                     command=codigo) 
button_run.grid(column=0,row=0 )

def boton():
   
    path=filename
etiqueta = Label(tab2)
etiqueta.pack()                        
window.mainloop()
