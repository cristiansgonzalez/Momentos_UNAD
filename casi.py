# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 17:21:35 2022

@author: Cristian Gonzalez
"""

from tkinter import *
from tkinter.messagebox import showinfo
from tkinter import filedialog
import os

class UNAD(Tk):
    
    def __init__(self):
        
        super().__init__()
        
        global entrada
        self.aux=0
        self.var = IntVar()
        
        # configure the root window
        self.title('Analisis')   
        self.geometry("500x500")   
        self.config(background = "red")
        self.iconbitmap("Imagenes/unad.ico")
        
        #Explica que hace el software
        self.Tutorial()
      
        # Boton para abrir el documento
        self.btOpen = Button(self, text ="Abrir Documento",bd = 3,command = self.abrirArchivo,width=24)
        self.btOpen.place(x=302 ,y=20)
        self.btOpen.bind("<Enter>",self.bh)
        self.btOpen.bind("<Leave>",self.bhl)
      
        # Letrero que dice numero total de actividades
        lbNdeActivades=Label(self,text="# total de actividades",bd = 3,width=25)
        lbNdeActivades.place(x=300,y=110)
        
        #Espacio donde se ingresa la cantidad de actividades
        entrada = Entry(self, bd = 3,width=20)
        entrada.place(x=300,y=145)
      
        # boton Ok  para que aparezcan las opciones
        self.btOk = Button(self, text ="Ok", bd = 3,command = self.actividades, width=5)
        self.btOk.place(x=437 ,y=143)
        self.btOk.bind("<Enter>",self.bh1)
        self.btOk.bind("<Leave>",self.bhl1)
        
        # Letrero actividades a escoger
        lbActividad=Label(self,text="Seleccione la actividad que\ndesea analizar",bd = 3,width=25)
        lbActividad.place(x=300,y=180)
        
        # letrero ingrese el puntaje de la actividad
        lbpuntaje=Label(self,text="Puntaje de la actividad",bd = 3,width=25)
        lbpuntaje.place(x=300,y=340)
        
        #self.actividades()
        
        
        # Informa todo lo que sucede en el programa
        self.lbinformacion=Label(self,bg="white",text="\n",bd = 3,width=66)
        self.lbinformacion.place(x=15,y=450)
        # button
        
    def abrirArchivo(self):    
        archivo_abierto = filedialog.askopenfilename(initialdir = os.getcwd(),
                                                     title = "Seleccinar archivo",
                                                     filetypes = (("xlsx files","*.xlsx"),
                                                     ("all files","*.*")))
        a=archivo_abierto.split("/")
        if archivo_abierto != "":
            self.aux=1
            print(type(a))
            self.lbinformacion.config(text=f"El documento {a[-1]} a sido cargado exitosamente\n")

        
    def bh(self,e):
        self.btOpen["bg"]="white"
    def bhl(self,e):
        self.btOpen["bg"]="SystemButtonFace"
    def bh1(self,e):
        self.btOk["bg"]="white"
    def bhl1(self,e):
        self.btOk["bg"]="SystemButtonFace"
        
    
        
        
    def actividades(self):
        
        lbNdeActivades=Label(self,bg="red",text="",bd = 3,height=6,width=26)
        lbNdeActivades.place(x=300,y=230)
        
        print(entrada.get())
        print(self.aux)
        if entrada.get() =="6" and self.aux==1:       
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R5 = Radiobutton(self, text = "5", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R6 = Radiobutton(self, text = "6", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 8, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 9, command = self.sel)
            
            self.R1.place(x=300 ,y=230)
            self.R2.place(x=350 ,y=230)
            self.R3.place(x=400 ,y=230)
            self.R4.place(x=450 ,y=230)
            self.R5.place(x=300 ,y=266)
            self.R6.place(x=350 ,y=266)
            self.R75.place(x=304 ,y=303)
            self.R25.place(x=369 ,y=303)
            self.RDef.place(x=434 ,y=303)
            
        elif entrada.get() =="5" and self.aux==1:
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R5 = Radiobutton(self, text = "5", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 8, command = self.sel)
            
            self.R1.place(x=300 ,y=230)
            self.R2.place(x=350 ,y=230)
            self.R3.place(x=400 ,y=230)
            self.R4.place(x=450 ,y=230)
            self.R5.place(x=300 ,y=266)
            self.R75.place(x=304 ,y=303)
            self.R25.place(x=369 ,y=303)
            self.RDef.place(x=434 ,y=303)
            
        elif entrada.get() =="4" and self.aux==1:
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            
            self.R1.place(x=300 ,y=230)
            self.R2.place(x=350 ,y=230)
            self.R3.place(x=400 ,y=230)
            self.R4.place(x=450 ,y=230)
            self.R75.place(x=304 ,y=266)
            self.R25.place(x=369 ,y=266)
            self.RDef.place(x=434 ,y=266)
        
        else:
            self.lbinformacion.config(text=f"=====================\nNo a seleccionado el documento")
     
            
    def sel(self):
        
        self.lbinformacion.config(text=f"Se analizara la actividad {self.var.get()}\n")
    
    def Tutorial(self):        
        # Tutoria de uso
        lbintrucciones=Label(self,bg="white",text="\n-Descargar el reporte de notas desde el \n"+
                 "centralizador.\n"+
                 "-BUSCAR: Seleccionar el documento \n"+
                 "descargado\n"+
                 "-Despues de que aparezca el mensaje \n"+
                 "escriba el numero de la actividad que desea \n"+
                 "analizar\n\n\n\n\n\n\n\n\n\n\n\n",
                 width=35)
        lbintrucciones.place(x=15,y=15)
        
        
app = UNAD()
app.mainloop()
