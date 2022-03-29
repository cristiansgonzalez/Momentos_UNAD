# -*- coding: utf-8 -*-
"""
Created on Mon Feb 28 16:33:48 2022

@author: Cristian Gonzalez
"""

import pandas as pd
import numpy as np

archivo = 'Sabana_de_notas_1141_243005.xlsx'
df = pd.read_excel(archivo, sheet_name='Sabana_de_notas')
df=df.fillna('N')

Actividad_1=25
N_intentos=10

fil ,col=df.shape
estudiante=[]
documento=[]
correo=[]
Aprobado=0
Ceros=0
Center=[]
Zone=[]
Program=[]
inten=[]
generacion=[]
Reprobaron=0

intentos=range(1,1+N_intentos)
intento2=(np.zeros((len(intentos))))
generacion1=0
generacion2=0

Centros=['Acacias','Aguachica','Barrancabermeja','Boavita',
         'Bucaramanga','Cali','Cartagena','Chiquinquirá',
         'Corozal','Cubará','Cúcuta','Curumaní','Dos quebradas',
         'Duitama','Facatativa','Florencia','Fusagasugá',
         'Garagoa','Girardot','Ibagué','José Acevedo y Gomez',
         'La Dorada','La Guajira','La Plata','Leticia','libano',
         'Mariquita','Medellin','Neiva','Ocaña','Palmira',
         'Pamplona','Pasto','PATIA (EL BORDO)','Pitalito',
         'Popayán','Puerto Asis','Puerto carreño',
         'Puerto Colombia','Sahagún ','Santa Marta',
         'Santander de quilichao','Soacha','Sogamoso',
         'Tumaco','Tunja','Turbo','Valle del Guamuez',
         'Valledupar','Velez','Yopal','Zipaquirá']

Programa=['INGENIERIA DE TELECOMUNICACIONES (Resolución 14518)',
         'INGENIERÍA ELECTRÓNICA (Resolución 13155)',
         'TECNOLOGIA EN AUTOMATIZACION ELECTRONICA']

Zonas=['Zona Amazonia Orinoquia','Zona Caribe',
       'Zona Centro Bogotá Cundinamarca','Zona Centro Boyacá',
       'Zona Centro Sur','Zona CentroOriente',
       'Zona Occidente','Zona Sur']

Programa1=(np.zeros((len(Programa))))
Programa2=(np.zeros((len(Programa))))
Cent1=(np.zeros((len(Centros))))
Cent2=(np.zeros((len(Centros))))
zona1=(np.zeros((len(Zonas))))
zona2=(np.zeros((len(Zonas))))

for X in range(15, fil):
    if df.iloc[X, 8]=='NO PRESENTADA':
        #print('RESPUESTA ', str(Ceros))
        estudiante.append(df.iloc[X, 3])
        documento.append(df.iloc[X, 2])
        correo.append(df.iloc[X, 4])
        Ceros=Ceros+1
        
        if (df.iloc[X, 27])[0]=='G':
            generacion1=generacion1+1
            
        for M in range(0, len(intentos)):
            if df.iloc[X, 7]==intentos[M]:
                intento2[M]=intento2[M]+1
        
        for Y in range(0,len(Zonas)):    
            if df.iloc[X, 20]==Zonas[Y]:
                zona1[Y]=zona1[Y]+1
        
        for Y in range(0,len(Centros)):    
            if df.iloc[X, 21]==Centros[Y]:
                Cent1[Y]=Cent1[Y]+1
                
        for Y in range(0,len(Programa)):    
            if df.iloc[X, 22]==Programa[Y]:
                Programa1[Y]=Programa1[Y]+1
    
    else:
        if float(df.iloc[X, 8])>=Actividad_1*0.6:
            Aprobado=Aprobado+1
            #print('Aprobado: ', str(Aprobado))
        
        if float(df.iloc[X, 8])<Actividad_1*0.6:
            Reprobaron=Reprobaron+1
            #print('Reprobaron: ', str(Reprobaron))
            for Y in range(0,len(Zonas)):    
                if df.iloc[X, 20]==Zonas[Y]:
                    zona2[Y]=zona2[Y]+1
            
            for Y in range(0,len(Centros)):    
                if df.iloc[X, 21]==Centros[Y]:
                    Cent2[Y]=Cent2[Y]+1
                    
            for Y in range(0,len(Programa)):    
                if df.iloc[X, 22]==Programa[Y]:
                    Programa2[Y]=Programa2[Y]+1 
                    
            if (df.iloc[X, 27])[0]=='G':
                generacion2=generacion2+1
                #print(generacion2)
        
productos = np.transpose([documento,estudiante,correo ])

grafica=[]
grafica.append(['Total Estudiantes',fil-15])
grafica.append(['Estudiantes Participaron Etapa 1',fil-15-Ceros])
grafica.append(['Estudiantes No Participaron Etapa 1',Ceros])

grafica.append(['',''])
grafica.append(['Total Estudiantes',fil-15])
grafica.append(['Estudiantes Participaron Etapa 1',fil-15-Ceros])
grafica.append(['Estudiantes Participaron y Aprobaron Etapa 1',Aprobado])
grafica.append(['Estudiantes Participaron y No Aprobaron Etapa 1',Reprobaron])

grafica.append(['',''])
grafica.append(['Aprobaron %',(Aprobado/(fil-15))*100])
grafica.append(['No participo %',(Ceros/(fil-15))*100])
grafica.append(['Reprobaron %',(Reprobaron/(fil-15))*100])

grafica.append(['',''])
grafica.append(['Generacion E',''])
grafica.append(['Estudiantes No Participaron Etapa 1',generacion1])
grafica.append(['Estudiantes Reprobaron ',generacion2])


for Z in range(0,len(Centros)):
    Center.append([Centros[Z],Cent1[Z],Cent2[Z]])
for Z in range(0,len(Programa)):
    Program.append([Programa[Z],Programa1[Z],Programa2[Z]])
for Z in range(0,len(Zonas)):
    Zone.append([Zonas[Z],zona1[Z],zona2[Z]])
a=np.arange(0,N_intentos)+1
for Z in range(0,len(intento2)):
    inten.append([a[Z],intento2[Z]])

est = pd.DataFrame (productos)
po = pd.DataFrame (Program)
ce = pd.DataFrame (Center)
gra = pd.DataFrame (grafica)
zo = pd.DataFrame (Zone)
te =pd.DataFrame (inten)

with pd.ExcelWriter('Momento 1.xlsx') as writer:  
    est.to_excel(writer, sheet_name='Estudiantes',header=['Cedula','Estudiante','Correo'],index=False)
    gra.to_excel(writer, sheet_name='Grafica',header=False,index=False)
    ce.to_excel(writer, sheet_name='Centros',header=['Centros','Ceros','Reprobaron'],index=False)
    po.to_excel(writer, sheet_name='Programa',header=['Programa','Ceros','Reprobaron'],index=False)
    te.to_excel(writer, sheet_name='Intentos',header=['Intentos','Ceros'],index=False)
    zo.to_excel(writer, sheet_name='Zonas',header=['Zonas','Ceros','Reprobaron'],index=False)
    
    
    
