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

Actividad_1=25 #puntaje de la actividad
N_intentos=10

fil ,col=df.shape
estudiante=[]
documento=[]
correo=[]
Aprobado=0
Ceros=0
Reprobaron=0



for X in range(15, fil):
    if df.iloc[X, 8]=='NO PRESENTADA':
        #print('RESPUESTA ', str(Ceros))
        estudiante.append(df.iloc[X, 3])
        documento.append(df.iloc[X, 2])
        correo.append(df.iloc[X, 4])
        Ceros=Ceros+1
    
    else:
        if float(df.iloc[X, 8])>=Actividad_1*0.6:
            Aprobado=Aprobado+1
            #print('Aprobado: ', str(Aprobado))
        
        if float(df.iloc[X, 8])<Actividad_1*0.6:
            Reprobaron=Reprobaron+1
            #print('Reprobaron: ', str(Reprobaron))
        
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



est = pd.DataFrame (productos)
gra = pd.DataFrame (grafica)

with pd.ExcelWriter('Momento 1.xlsx') as writer:  
    est.to_excel(writer, sheet_name='Estudiantes',header=['Cedula','Estudiante','Correo'],index=False)
    gra.to_excel(writer, sheet_name='Grafica',header=False,index=False)

    
    