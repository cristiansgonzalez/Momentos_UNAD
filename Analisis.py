# -*- coding: utf-8 -*-
"""
Created on Fri Apr  1 22:13:40 2022

@author: Cristian Gonzalez
"""

def Analisis_Curso(*arg):
    
    import pandas as pd
    import numpy as np

    df = pd.read_excel(arg[0]+'.xlsx', sheet_name='Sabana_de_notas')    
    df=df.fillna('**********')
    
  
    inicial=0
    #Etapa=1
    fil ,col=df.shape
    for x in range(0,fil):
        if df.iloc[x,0]=='Grupo':
            inicial=x+1
            
            break
        
    estudiante=[]
    documento=[]
    correo=[]
    Aprobado=0
    Ceros=0
    Reprobaron=0
    
    for X in range(inicial, fil):
        
        if df.iloc[X, 5]=='Matriculado':
            #print(f'{df.iloc[X, 5]} {X}')
                    
            if df.iloc[X, 7+arg[2]]=='NO PRESENTADA' or df.iloc[X, 7+arg[2]]==0:
                
                #print('RESPUESTA ', str(Ceros))
                estudiante.append(df.iloc[X, 3])
                documento.append(df.iloc[X, 2])
                correo.append(df.iloc[X, 4])
                Ceros=Ceros+1
            
            else:
                if float(df.iloc[X, 7+arg[2]])>=arg[1]*0.6:
                    #print('Aprobado  '+df.iloc[X, 7+arg[2]])
                    
                    Aprobado=Aprobado+1
                    #print('Aprobado: ', str(Aprobado))
                
                if float(df.iloc[X, 7+arg[2]])<arg[1]*0.6:
                    Reprobaron=Reprobaron+1
                    estudiante.append(df.iloc[X, 3])
                    documento.append(df.iloc[X, 2])
                    correo.append(df.iloc[X, 4])
                    #print('Reprobaron: ', str(Reprobaron)
            
    productos = np.transpose([documento,estudiante,correo ])
    
    grafica=[]
    grafica.append(['Total Estudiantes',fil-inicial])
    grafica.append(['Estudiantes Participaron Etapa '+str(arg[2]),fil-inicial-Ceros])
    grafica.append(['Estudiantes No Participaron Etapa '+str(arg[2]),Ceros])
    
    grafica.append(['',''])
    grafica.append(['Total Estudiantes',fil-inicial])
    grafica.append(['Estudiantes Participaron Etapa '+str(arg[2]),fil-inicial-Ceros])
    grafica.append(['Estudiantes Participaron y Aprobaron Etapa '+str(arg[2]),Aprobado])
    grafica.append(['Estudiantes Participaron y No Aprobaron Etapa '+str(arg[2]),Reprobaron])
    
    grafica.append(['',''])
    grafica.append(['Aprobaron %',(Aprobado/(fil-inicial))*100])
    grafica.append(['No participo %',(Ceros/(fil-inicial))*100])
    grafica.append(['Reprobaron %',(Reprobaron/(fil-inicial))*100])
    
    est = pd.DataFrame (productos)
    gra = pd.DataFrame (grafica)
 
    
    try:
        with pd.ExcelWriter('Reporte '+str(arg[2])+'.xlsx') as writer:  
            est.to_excel(writer, sheet_name='Estudiantes',header=['Cedula','Estudiante','Correo'],index=False)
            gra.to_excel(writer, sheet_name='Grafica',header=False,index=False)
    except:        
        print('\n\nNo se puedo crear el '+'Reporte '+str(arg[2])+'.xlsx\n')
    finally:        
        print('El Reporte '+str(arg[2])+' Se creo Exitosamente')
        
        