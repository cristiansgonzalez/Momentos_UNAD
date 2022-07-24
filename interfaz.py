# -*- coding: utf-8 -*-
"""
Created on Thu Jul 21 17:12:56 2022

@author: Cristian Gonzalez
"""

from tkinter import filedialog
from tkinter import *
import os

global aux

aux=0


def abrirArchivo():    
    archivo_abierto = filedialog.askopenfilename(initialdir = os.getcwd(),
                                                 title = "Seleccinar archivo",
                                                 filetypes = (("xlsx files","*.xlsx"),
                                                 ("all files","*.*")))
    print(type(archivo_abierto))
    if archivo_abierto != "":
        aux=1
        lbinformacion.config(text="Documento cargado exitosamente\n")

def bh(e):
    btOpen["bg"]="white"
def bhl(e):
    btOpen["bg"]="SystemButtonFace"
def bh1(e):
    btOk["bg"]="white"
def bhl1(e):
    btOk["bg"]="SystemButtonFace"
    
def sel():
    lbinformacion.config(text=f"Se analizara la actividad {var.get()}\n")
def actividades():
    
    if entrada.get() =="4":       
        
        R1 = Radiobutton(ventana, text = "1", state = NORMAL ,variable = var, value = 1, command = sel)
        R2 = Radiobutton(ventana, text = "2", state = NORMAL ,variable = var, value = 2, command = sel)
        
        print(var)
        R1.place(x=300 ,y=230)
        R2.place(x=350 ,y=230)
        
    elif entrada.get() =="5":
        R2 = Radiobutton(ventana, text = "1")#, command = sel)
        R2.place(x=300 ,y=220)
ventana = Tk()
ventana.title('Analisis')   
ventana.geometry("500x500")   
ventana.config(background = "red")
ventana.iconbitmap('unad.ico')

lbintrucciones=Label(ventana,bg="white",text="\n-Descargar el reporte de notas desde el \n"+
                     "centralizador.\n"+
                     "-BUSCAR: Seleccionar el documento \n"+
                     "descargado\n"+
                     "-Despues de que aparezca el mensaje \n"+
                     "escriba el numero de la actividad que desea \n"+
                     "analizar\n\n\n\n\n\n\n\n\n\n\n\n",
                     width=35)
lbintrucciones.place(x=15,y=15)

btOpen = Button(ventana, text ="Buscar",bd = 3,command = abrirArchivo,width=24)
btOpen.place(x=302 ,y=20)
btOpen.bind("<Enter>",bh)
btOpen.bind("<Leave>",bhl)



lbNdeActivades=Label(ventana,text="# total de actividades",bd = 3,width=25)
lbNdeActivades.place(x=300,y=110)

entrada = Entry(ventana, bd = 3,width=20)
entrada.place(x=300,y=145)

var = IntVar()
btOk = Button(ventana, text ="Ok", bd = 3,command = actividades, width=5)
btOk.place(x=437 ,y=143)
btOk.bind("<Enter>",bh1)
btOk.bind("<Leave>",bhl1)

lbActividad=Label(ventana,text="Seleccione la actividad que\ndesea analizar",bd = 3,width=25)
lbActividad.place(x=300,y=180)

lbinformacion=Label(ventana,bg="white",text="\n",bd = 3,width=66)
lbinformacion.place(x=15,y=350)

'''
salir = Button(ventana, text = "Exit",bd = 3, command = ventana.destroy, width=10) 
salir.grid(column = 1,row = 3)

entrada = Entry(ventana, bd = 5)
entrada.grid(column = 1,row = 4)
ok = Button(ventana, text = "Ok",bd = 3, command = lambda: print(entrada.get()), width=10) 
print(ok)
ok.grid(column = 2,row = 4)

print(entrada.get())

'''



ventana.mainloop() 


