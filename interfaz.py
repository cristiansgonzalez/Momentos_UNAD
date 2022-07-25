# -*- coding: utf-8 -*-
"""
Created on Fri Jul 22 17:21:35 2022

@author: Cristian Gonzalez
"""

from tkinter import *
from tkinter.messagebox import showinfo
from tkinter import filedialog
import os
import test
import Analisis2

class UNAD(Tk):
    
    def __init__(self):
        
        super().__init__()
        
        global NdeActivades, puntajeActivades
        self.aux = 0
        self.var = IntVar(self, 0)
        self.TDAVar = StringVar(self, "0")
        self.archivo_abierto = ""
        self.bajar = 27
        
        # configure the root window
        self.title('Analisis')   
        self.geometry("500x527")   
        self.config(background = "red")
        self.iconbitmap("Imagenes/unad.ico")
        self.resizable(0,0)
        
        #Explica que hace el software
        self.Tutorial()
      
        # Boton para abrir el documento
        self.btOpen = Button(self, text = "Abrir Documento", bd = 3, command = self.abrirArchivo, width = 24)
        self.btOpen.place(x = 302 , y = 20)
        self.btOpen.bind("<Enter>", self.bh)
        self.btOpen.bind("<Leave>", self.bhl)

        # Inicializa los radiobuton estapa paso...
        self.tipoDeActividad()        

        # Letrero que dice numero total de actividades
        lbNdeActivades=Label(self, text = "# total de actividades", bd = 3, width = 25)
        lbNdeActivades.place(x = 300, y = 110 + self.bajar)
        
        #Espacio donde se ingresa la cantidad de actividades
        NdeActivades = Entry(self, bd = 3, width = 20)
        NdeActivades.place(x = 300, y = 145 + self.bajar)
      
        # boton Ok  para que aparezcan las opciones
        self.btOk = Button(self, text = "Ok", bd = 3, command = self.actividades, width = 5)
        self.btOk.place(x = 437 ,y = 143 + self.bajar)
        self.btOk.bind("<Enter>", self.bh1)
        self.btOk.bind("<Leave>", self.bhl1)
        
        # Letrero actividades a escoger
        lbActividad = Label(self, text = "Seleccione la actividad que\ndesea analizar", bd = 3, width = 25)
        lbActividad.place(x = 300, y = 180 + self.bajar)
        
        # letrero ingrese el puntaje de la actividad
        lbpuntaje = Label(self, text = "Puntaje de la actividad", bd = 3, width = 25)
        lbpuntaje.place(x = 300, y = 340 + self.bajar)
        
        #Espacio donde se ingresa el puntaje de la actividad a analizar
        puntajeActivades = Entry(self, bd = 3, width = 29)
        puntajeActivades.place(x = 300, y = 375 + self.bajar)

        # boton Analizar Ejecuta todo el codigo 
        self.btAnalizar = Button(self, text = "Analizar", bd = 3, command = self.Analizar, width = 24)
        self.btAnalizar.place(x = 302, y = 410 + self.bajar)
        self.btAnalizar.bind("<Enter>", self.bh2)
        self.btAnalizar.bind("<Leave>", self.bhl2)
        
        # Informa todo lo que sucede en el programa
        self.lbinformacion = Label(self, bg = "white", text = "\n", bd = 3, width = 65)
        self.lbinformacion.place(x = 16, y = 450 + self.bajar)
        # button
        
    def abrirArchivo(self):    
        self.archivo_abierto = filedialog.askopenfilename(initialdir = os.getcwd(),
                                                     title = "Seleccinar archivo",
                                                     filetypes = (("xlsx files", "*.xlsx"),
                                                     ("all files", "*.*")))
        a = self.archivo_abierto.split("/")
        if self.archivo_abierto != "":
            self.aux = 1
            print(type(a))
            self.lbinformacion.config(text = f"El documento {a[-1]} a sido cargado exitosamente\n")

    def tipoDeActividad(self):
        self.TDA1 = Radiobutton(self, text = "Paso", state = NORMAL, variable = self.TDAVar, value = "Paso", command = self.blanco)
        self.TDA2 = Radiobutton(self, text = "Etapa", state = NORMAL, variable = self.TDAVar, value = "Etapa", command = self.blanco)
        self.TDA3 = Radiobutton(self, text = "Tarea", state = NORMAL, variable = self.TDAVar, value = "Tarea", command = self.blanco)
        self.TDA4 = Radiobutton(self, text = "Fase", state = NORMAL, variable = self.TDAVar, value = "Fase", command = self.blanco)
        
        self.TDA1.place(x=300 ,y=60)
        self.TDA2.place(x=365 ,y=60)
        self.TDA3.place(x=300 ,y=98)
        self.TDA4.place(x=367 ,y=98)

    def blanco(self):
        self.lbinformacion.config(text="Selecciono " + self.TDAVar.get() + "\n")

    def bh(self,e):
        self.btOpen["bg"] = "white"
    def bhl(self,e):
        self.btOpen["bg"] = "SystemButtonFace"
    def bh1(self,e):
        self.btOk["bg"] = "white"
    def bhl1(self,e):
        self.btOk["bg"] = "SystemButtonFace"
    def bh2(self,e):
        self.btAnalizar["bg"] = "white"
    def bhl2(self,e):
        self.btAnalizar["bg"] = "SystemButtonFace"
        
    def Analizar(self):
        test.visualizar(self.archivo_abierto, int(puntajeActivades.get()), self.var.get(), int(NdeActivades.get()), "si", self.TDAVar.get())
        informacion = Analisis2.Analisis_Curso(self.archivo_abierto, int(puntajeActivades.get()), self.var.get(), int(NdeActivades.get()), "si", self.TDAVar.get())
        self.lbinformacion.config(text = informacion)
        
    def actividades(self):
        
        lbNdeActivades=Label(self,bg="red",text="",bd = 3,height=6,width=26)
        lbNdeActivades.place(x=300,y=230)
        
        print("Numero total de actividades")
        print(type(NdeActivades.get()))
        print(self.aux)
        if NdeActivades.get() =="6" and self.aux==1:       
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R5 = Radiobutton(self, text = "5", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R6 = Radiobutton(self, text = "6", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 8, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 9, command = self.sel)
            
            self.R1.place(x=300 ,y=230 + self.bajar)
            self.R2.place(x=350 ,y=230 + self.bajar)
            self.R3.place(x=400 ,y=230 + self.bajar)
            self.R4.place(x=450 ,y=230 + self.bajar)
            self.R5.place(x=300 ,y=266 + self.bajar)
            self.R6.place(x=350 ,y=266 + self.bajar)
            self.R75.place(x=304 ,y=303 + self.bajar)
            self.R25.place(x=369 ,y=303 + self.bajar)
            self.RDef.place(x=434 ,y=303 + self.bajar)
            
        elif NdeActivades.get() =="5" and self.aux==1:
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R5 = Radiobutton(self, text = "5", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 8, command = self.sel)
            
            self.R1.place(x=300 ,y=230 + self.bajar)
            self.R2.place(x=350 ,y=230 + self.bajar)
            self.R3.place(x=400 ,y=230 + self.bajar)
            self.R4.place(x=450 ,y=230 + self.bajar)
            self.R5.place(x=300 ,y=266 + self.bajar)
            self.R75.place(x=304 ,y=303 + self.bajar)
            self.R25.place(x=369 ,y=303 + self.bajar)
            self.RDef.place(x=434 ,y=303 + self.bajar)
            
        elif NdeActivades.get() =="4" and self.aux==1:
            
            self.R1 = Radiobutton(self, text = "1", state = NORMAL ,variable = self.var, value = 1, command = self.sel)
            self.R2 = Radiobutton(self, text = "2", state = NORMAL ,variable = self.var, value = 2, command = self.sel)
            self.R3 = Radiobutton(self, text = "3", state = NORMAL ,variable = self.var, value = 3, command = self.sel)
            self.R4 = Radiobutton(self, text = "4", state = NORMAL ,variable = self.var, value = 4, command = self.sel)
            self.R75 = Radiobutton(self, text = "75%", state = NORMAL ,variable = self.var, value = 5, command = self.sel)
            self.R25 = Radiobutton(self, text = "25%", state = NORMAL ,variable = self.var, value = 6, command = self.sel)
            self.RDef = Radiobutton(self, text = "Def", state = NORMAL ,variable = self.var, value = 7, command = self.sel)
            
            self.R1.place(x=300 ,y=230 + self.bajar)
            self.R2.place(x=350 ,y=230 + self.bajar)
            self.R3.place(x=400 ,y=230 + self.bajar)
            self.R4.place(x=450 ,y=230 + self.bajar)
            self.R75.place(x=304 ,y=266 + self.bajar)
            self.R25.place(x=369 ,y=266 + self.bajar)
            self.RDef.place(x=434 ,y=266 + self.bajar)
        
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
                 "analizar",
                 height=27,width=35)
        lbintrucciones.place(x=15,y=15)
        
        
app = UNAD()
app.mainloop()
