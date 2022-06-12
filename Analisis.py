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
    
    #puntaje_actividad=25 #puntaje de la actividad
    
    #Etapa=1
    fil ,col=df.shape
    for X in range(0,fil):
        if df.iloc[X,0]=='Grupo':
            inicial=X+1            
            break
    
    titulos=[]
    for X in  df.iloc[inicial-1,:]:
        titulos.append(X)
        
    N_intentos=max(df.iloc[inicial::, titulos.index('Num Intento')])
    #np.argwhere(df.iloc[:,:]=='Grupo')
    
    estudiante=[]
    documento=[]
    correo=[]
    
    Aprobado=0
    Ceros=0
    Reprobaron=0
    aplazado=0
    
    Center=[]
    Zone=[]
    Program=[]
    inten=[]
    
    
    condi_estudiante=[]
    condi_documento=[]
    condi_correo=[]
    condicion=[]
    
    #Etapas_totales=5
    
    intentos=range(1,1+N_intentos)
    intento2=(np.zeros((len(intentos))))
    generacion1=0
    generacion2=0
    
    # se buscan las columnas con los parametros solicitados
    col_centro=titulos.index('Centro')
    col_programa=titulos.index('Programa')
    col_zona=titulos.index('Zona')
    col_matricula=titulos.index('Estado Matricula')
    col_generacion_e=titulos.index('Convenios')+1
    col_condi_especiales=titulos.index('Necesidades Especiales')
    
    
    #Se guardan todos los centros, programas y zonas
    Centros=list(set((df.iloc[inicial::,col_centro])))
    Programa=list(set((df.iloc[inicial::,col_programa])))
    Zonas=list(set((df.iloc[inicial::,col_zona])))
        
    Programa1=(np.zeros((len(Programa))))
    Programa2=(np.zeros((len(Programa))))
    Cent1=(np.zeros((len(Centros))))
    Cent2=(np.zeros((len(Centros))))
    zona1=(np.zeros((len(Zonas))))
    zona2=(np.zeros((len(Zonas))))
    aprobado_programa=(np.zeros((len(Programa))))
    aprobado_zona=(np.zeros((len(Zonas))))
    aprobado_Cent=(np.zeros((len(Centros))))
    
    for X in range(inicial, fil):
        
        if df.iloc[X, col_matricula]=='Reportado' or df.iloc[X, col_matricula]=='Matriculado':
            
            
            if df.iloc[X, col_condi_especiales]!='**********':
                print(df.iloc[X, col_condi_especiales])
                condi_estudiante.append(df.iloc[X, 3])
                condi_documento.append(df.iloc[X, 2])
                condi_correo.append(df.iloc[X, 4])
                condicion.append(df.iloc[X, col_condi_especiales])
                
            print(f'{df.iloc[X, 5]} {X}')
                    
            if df.iloc[X, 7+arg[2]]=='NO PRESENTADA' or df.iloc[X, 7+arg[2]]==0:
                
                #print('RESPUESTA ', str(Ceros))
                estudiante.append(df.iloc[X, 3])
                documento.append(df.iloc[X, 2])
                correo.append(df.iloc[X, 4])
                Ceros=Ceros+1
                
                if arg[4]=='si':
                    if (df.iloc[X, col_generacion_e])[0]=='G':
                        generacion1=generacion1+1
                    
                for M in range(0, len(intentos)):
                    if df.iloc[X, 7]==intentos[M]:
                        intento2[M]=intento2[M]+1
                
                for Y in range(0,len(Zonas)):    
                    if df.iloc[X, col_zona]==Zonas[Y]:
                        zona1[Y]=zona1[Y]+1
                
                for Y in range(0,len(Centros)):    
                    if df.iloc[X, col_centro]==Centros[Y]:
                        Cent1[Y]=Cent1[Y]+1
                        
                for Y in range(0,len(Programa)):    
                    if df.iloc[X, col_programa]==Programa[Y]:
                        Programa1[Y]=Programa1[Y]+1 
                        
                    
                print(f'{df.iloc[X, 7+arg[2]]} {X}')
            
            else:
                if float(df.iloc[X, 7+arg[2]])>=arg[1]*0.6:
                    #print('Aprobado  '+df.iloc[X, 7+arg[2]])
                    
                    Aprobado=Aprobado+1
                    for Y in range(0,len(Programa)):    
                        if df.iloc[X, col_programa]==Programa[Y]:
                            aprobado_programa[Y]=aprobado_programa[Y]+1
                            
                    for Y in range(0,len(Centros)):    
                        if df.iloc[X, col_centro]==Centros[Y]:
                            aprobado_Cent[Y]=aprobado_Cent[Y]+1
                            
                    for Y in range(0,len(Zonas)):    
                        if df.iloc[X, col_zona]==Zonas[Y]:
                            aprobado_zona[Y]=aprobado_zona[Y]+1
                    
                    #print('Aprobado: ', str(Aprobado))
                
                if float(df.iloc[X, 7+arg[2]])<arg[1]*0.6:
                    Reprobaron=Reprobaron+1
                    estudiante.append(df.iloc[X, 3])
                    documento.append(df.iloc[X, 2])
                    correo.append(df.iloc[X, 4])
                    #print('Reprobaron: ', str(Reprobaron))
                    for Y in range(0,len(Zonas)):    
                        if df.iloc[X, col_zona]==Zonas[Y]:
                            zona2[Y]=zona2[Y]+1
                    
                    for Y in range(0,len(Centros)):    
                        if df.iloc[X, col_centro]==Centros[Y]:
                            Cent2[Y]=Cent2[Y]+1
                            
                    for Y in range(0,len(Programa)):    
                        if df.iloc[X, col_programa]==Programa[Y]:
                            Programa2[Y]=Programa2[Y]+1
                            
                    
                    if arg[4]=='si':
                        if (df.iloc[X, col_generacion_e])[0]=='G':
                            generacion2=generacion2+1
                        #print(generacion2)
        elif df.iloc[X, col_matricula]=='Aplazado':
            aplazado+=1
            
    productos = np.transpose([documento,estudiante,correo ])
    
    for z in range(0,len(Centros)):
        Center.append([Centros[z],Cent1[z],Cent2[z],aprobado_Cent[z]])   
    for z in range(0,len(Programa)):
        Program.append([Programa[z],Programa1[z],Programa2[z],aprobado_programa[z]])
    for z in range(0,len(Zonas)):
        Zone.append([Zonas[z],zona1[z],zona2[z],aprobado_zona[z]])
    
    inten=np.transpose([np.arange(0,N_intentos)+1,intento2])
    condiciones=np.transpose([condi_documento,condi_estudiante,condi_correo,condicion])
    
    grafica=[]
    grafica.append(['Total Estudiantes',fil-inicial-aplazado])
    grafica.append(['Estudiantes Participaron Etapa '+str(arg[2]),fil-inicial-Ceros-aplazado])
    grafica.append(['Estudiantes No Participaron Etapa '+str(arg[2]),Ceros])
    
    grafica.append(['',''])
    grafica.append(['Total Estudiantes',fil-inicial-aplazado])
    grafica.append(['Estudiantes Participaron Etapa '+str(arg[2]),fil-inicial-Ceros-aplazado])
    grafica.append(['Estudiantes Participaron y Aprobaron Etapa '+str(arg[2]),Aprobado])
    grafica.append(['Estudiantes Participaron y No Aprobaron Etapa '+str(arg[2]),Reprobaron])
    
    grafica.append(['',''])
    grafica.append(['Aprobaron %',(Aprobado/(fil-inicial-aplazado))*100])
    grafica.append(['No participo %',(Ceros/(fil-inicial-aplazado))*100])
    grafica.append(['Reprobaron %',(Reprobaron/(fil-inicial-aplazado))*100])
    
    grafica.append(['',''])
    grafica.append(['Generacion E',''])
    grafica.append(['Estudiantes No Participaron Etapa '+str(arg[2]),generacion1])
    grafica.append(['Estudiantes Reprobaron ',generacion2])
    
    est = pd.DataFrame (productos)
    po = pd.DataFrame (Program)
    ce = pd.DataFrame (Center)
    gra = pd.DataFrame (grafica)
    zo = pd.DataFrame (Zone)
    te =pd.DataFrame (inten)
    condi =pd.DataFrame (condiciones)
    
    try:
        with pd.ExcelWriter(arg[0]+' Reporte '+str(arg[2])+'.xlsx') as writer:  
            est.to_excel(writer, sheet_name='Estudiantes',header=['Cedula','Estudiante','Correo'],index=False)
            gra.to_excel(writer, sheet_name='Grafica',header=False,index=False)
            ce.to_excel(writer, sheet_name='Centros',header=['Centros','Ceros','Reprobaron','Aprobaron'],index=False)
            po.to_excel(writer, sheet_name='Programa',header=['Programa','Ceros','Reprobaron','Aprobaron'],index=False)
            te.to_excel(writer, sheet_name='Intentos',header=['Intentos','Ceros'],index=False)
            zo.to_excel(writer, sheet_name='Zonas',header=['Zonas','Ceros','Reprobaron','Aprobaron'],index=False)
            condi.to_excel(writer, sheet_name='Condiciones',header=['Cedula','Estudiante','Correo','Condicion especial'],index=False)
    except:        
        print('\n\nNo se puedo crear el '+'Reporte '+str(arg[2])+'.xlsx\n')
    finally:        
        print('El Reporte '+str(arg[2])+' Se creo Exitosamente')
        
        