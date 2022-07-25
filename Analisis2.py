
def Analisis_Curso(*arg):
    
    import pandas as pd
    import numpy as np
    import statistics
    nombre = str(arg[0])
    df = pd.read_excel(arg[0], sheet_name='Sabana_de_notas')    
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
    total_gen=[]
    total_matri0=[]
    prom_bene=[]
    prom_total=[]
    
    Aprobado=0
    Ceros=0
    Reprobaron=0
    aplazado=0
    
    Center=[]
    Zone=[]
    Program=[]
    inten=[]
    
    # Variables numero de intentos
    Ceros2=0
    generacion12=0
    matricula12=0
    Aprobado2=0
    apro_Gen2=0
    apro_Matri02=0
    Reprobaron2=0
    generacion22=0
    matricula22=0
    total_estudiantes2=0
    
    total_matri02=[]
    total_gen2=[]
    prom_bene2=[]
    prom_total2=[]
    
    
    condi_estudiante=[]
    condi_documento=[]
    condi_correo=[]
    condicion=[]
    
    #Etapas_totales=5
    
    intentos=range(1,1+N_intentos)
    intento2=(np.zeros((len(intentos))))
    apro_intento=(np.zeros((len(intentos))))
    repro_intento=(np.zeros((len(intentos))))
    generacion1=0
    generacion2=0
    matricula1=0
    matricula2=0
    apro_Gen=0
    apro_Matri0=0
    
    # se buscan las columnas con los parametros solicitados
    col_centro=titulos.index('Centro')
    col_programa=titulos.index('Programa')
    col_zona=titulos.index('Zona')
    col_matricula=titulos.index('Estado Matricula')
    col_generacion_e=titulos.index('Convenios')
    col_condi_especiales=titulos.index('Necesidades Especiales')
    #col_definitiva=titulos.index('Definitiva')
    col_75=titulos.index('75 %')
    
    
    #Se guardan todos los centros, programas y zonas
    Centros=list(set((df.iloc[inicial::,col_centro])))
    Centros.sort()
    Programa=list(set((df.iloc[inicial::,col_programa])))
    Programa.sort()
    Zonas=list(set((df.iloc[inicial::,col_zona])))
    Zonas.sort()
    
    Programa1=(np.zeros((len(Programa))))
    Programa2=(np.zeros((len(Programa))))
    Cent1=(np.zeros((len(Centros))))
    Cent2=(np.zeros((len(Centros))))
    zona1=(np.zeros((len(Zonas))))
    zona2=(np.zeros((len(Zonas))))
    aprobado_programa=(np.zeros((len(Programa))))
    aprobado_zona=(np.zeros((len(Zonas))))
    aprobado_Cent=(np.zeros((len(Centros))))
    
    
    if arg[2]-arg[3]==1:
        col_definitiva=titulos.index('75 %')
        arg[3]
    elif arg[2]-arg[3]==2:
        col_definitiva=titulos.index('25 %')
    else:
        col_definitiva=titulos.index('Definitiva')
        
 
    
    if df.iloc[inicial,col_75]=='**********':        
        for m in range(inicial,fil):
            for mm in range (0,3):
                df.iloc[m,col_75+mm]=0
    
    for X in range(inicial, fil):
        
        if df.iloc[X, col_matricula]=='Reportado' or df.iloc[X, col_matricula]=='Matriculado' or df.iloc[X, col_matricula]=='Reportado 75%':
            
            
            if df.iloc[X, col_condi_especiales]!='**********':
                condi_estudiante.append(df.iloc[X, 3])
                condi_documento.append(df.iloc[X, 2])
                condi_correo.append(df.iloc[X, 4])
                condicion.append(df.iloc[X, col_condi_especiales])
                    
            if df.iloc[X, 7+arg[2]]=='NO PRESENTADA' or df.iloc[X, 7+arg[2]]==0:
                
                #print('RESPUESTA ', str(Ceros))
                estudiante.append(df.iloc[X, 3])
                documento.append(df.iloc[X, 2])
                correo.append(df.iloc[X, 4])
                Ceros=Ceros+1
                #print('Ceros: ',Ceros)
                
                if arg[4]=='si':
                    if (df.iloc[X, col_generacion_e])[0]=='G':
                        generacion1=generacion1+1
                        total_gen.append(df.iloc[X, col_definitiva])                          
                
                    if (df.iloc[X, col_generacion_e])[0]=='M':
                        matricula1=matricula1+1
                        total_matri0.append(df.iloc[X, col_definitiva])
                        
                if (df.iloc[X, col_generacion_e])[0]!='G' and (df.iloc[X, col_generacion_e])[0]!='M':
                    prom_bene.append(df.iloc[X, col_definitiva])
                    
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
                        
                    
                #print(f'{df.iloc[X, 7+arg[2]]} {X}')
            
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
                            
                    for M in range(0, len(intentos)):
                        if df.iloc[X, 7]==intentos[M]:
                            apro_intento[M]=apro_intento[M]+1
                    
                    if arg[4]=='si':
                        if (df.iloc[X, col_generacion_e])[0]=='G':
                            apro_Gen=apro_Gen+1
                            total_gen.append(df.iloc[X, col_definitiva])
                    
                        if (df.iloc[X, col_generacion_e])[0]=='M':
                            apro_Matri0=apro_Matri0+1
                            total_matri0.append(df.iloc[X, col_definitiva])
                    
                    #print('Aprobado: ', str(Aprobado))
                
                if float(df.iloc[X, 7+arg[2]])<arg[1]*0.6:
                    Reprobaron=Reprobaron+1
                    estudiante.append(df.iloc[X, 3])
                    documento.append(df.iloc[X, 2])
                    correo.append(df.iloc[X, 4])
                    #print('Reprobaron: ', str(Reprobaron))
                    for M in range(0, len(intentos)):
                        if df.iloc[X, 7]==intentos[M]:
                            repro_intento[M]=repro_intento[M]+1
                    
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
                            total_gen.append(df.iloc[X, col_definitiva])
                            
                        if (df.iloc[X, col_generacion_e])[0]=='M':
                            matricula2=matricula2+1
                            total_matri0.append(df.iloc[X, col_definitiva])
                        #print(generacion2)
                
                if (df.iloc[X, col_generacion_e])[0]!='G' and (df.iloc[X, col_generacion_e])[0]!='M':
                    prom_bene.append(df.iloc[X, col_definitiva])
                
                
                #########################################
            if df.iloc[X, 7]>=2:
                
                total_estudiantes2=total_estudiantes2+1
                #print(df.iloc[X, 7+arg[2]])
                if df.iloc[X, 7+arg[2]]=='NO PRESENTADA' or df.iloc[X, 7+arg[2]]==0:

                    Ceros2=Ceros2+1
                    
                    if arg[4]=='si':
                        if (df.iloc[X, col_generacion_e])[0]=='G':
                            generacion12=generacion12+1
                            total_gen2.append(df.iloc[X, col_definitiva])
                            print(df.iloc[X, col_definitiva])
                    
                        if (df.iloc[X, col_generacion_e])[0]=='M':
                            matricula12=matricula12+1
                            total_matri02.append(df.iloc[X, col_definitiva])
                            
                    if (df.iloc[X, col_generacion_e])[0]!='G' and (df.iloc[X, col_generacion_e])[0]!='M':
                        prom_bene2.append(df.iloc[X, col_definitiva])

                else:
                    if float(df.iloc[X, 7+arg[2]])>=arg[1]*0.6:
                        
                        Aprobado2=Aprobado2+1
                        
                        if arg[4]=='si':
                            if (df.iloc[X, col_generacion_e])[0]=='G':
                                apro_Gen2=apro_Gen2+1
                                total_gen2.append(df.iloc[X, col_definitiva])
                        
                            if (df.iloc[X, col_generacion_e])[0]=='M':
                                apro_Matri02=apro_Matri02+1
                                total_matri02.append(df.iloc[X, col_definitiva])
                    
                    if float(df.iloc[X, 7+arg[2]])<arg[1]*0.6:
                        
                        Reprobaron2=Reprobaron2+1                   
                        
                        if arg[4]=='si':
                            if (df.iloc[X, col_generacion_e])[0]=='G':
                                generacion22=generacion22+1
                                total_gen2.append(df.iloc[X, col_definitiva])
                                
                            if (df.iloc[X, col_generacion_e])[0]=='M':
                                matricula22=matricula22+1
                                total_matri02.append(df.iloc[X, col_definitiva])
                            #print(generacion2)
                    
                if (df.iloc[X, col_generacion_e])[0]!='G' and (df.iloc[X, col_generacion_e])[0]!='M':
                        prom_bene2.append(df.iloc[X, col_definitiva])
                
         
                
        elif df.iloc[X, col_matricula]=='Aplazado':
            aplazado+=1
      
    
    
    productos = np.transpose([documento,estudiante,correo ])

    for z in range(0,len(Centros)):
        Center.append([Centros[z],Cent1[z],Cent2[z],aprobado_Cent[z]])   
    for z in range(0,len(Programa)):
        Program.append([Programa[z],Programa1[z],Programa2[z],aprobado_programa[z]])
    for z in range(0,len(Zonas)):        
        Zone.append([Zonas[z],zona1[z],zona2[z],aprobado_zona[z],zona1[z]+zona2[z]+aprobado_zona[z]])
    
    inten=np.transpose([np.arange(0,N_intentos)+1,intento2,repro_intento,apro_intento,repro_intento+apro_intento+intento2])
    condiciones=np.transpose([condi_documento,condi_estudiante,condi_correo,condicion])
    
    grafica=[]
    
    total_estudiantes=fil-inicial-aplazado
    grafica.append(['Total Estudiantes',total_estudiantes])
    grafica.append(['Estudiantes Participaron '+arg[5]+' '+str(arg[2]),total_estudiantes-Ceros])
    grafica.append(['Estudiantes No Participaron '+arg[5]+' '+str(arg[2]),Ceros])
    print("hola")
    grafica.append(['',''])
    grafica.append(['Total Estudiantes',total_estudiantes])
    grafica.append(['Estudiantes Participaron '+arg[5]+' '+str(arg[2]),total_estudiantes-Ceros])
    grafica.append(['Estudiantes No Participaron '+arg[5]+' '+str(arg[2]),Ceros])
    grafica.append(['Estudiantes Participaron y Aprobaron '+arg[5]+' '+str(arg[2]),Aprobado])
    grafica.append(['Estudiantes Participaron y No Aprobaron '+arg[5]+' '+str(arg[2]),Reprobaron])
    
    grafica.append(['',''])
    grafica.append(['Aprobaron %',(Aprobado/(total_estudiantes))*100])
    grafica.append(['No participo %',(Ceros/(total_estudiantes))*100])
    grafica.append(['Reprobaron %',(Reprobaron/(total_estudiantes))*100])
    
    grafica.append(['',''])
    grafica.append(['Total Generacion E',len(total_gen)])
    grafica.append(['Estudiantes No Participaron '+arg[5]+' '+str(arg[2]),generacion1])
    grafica.append(['Estudiantes Reprobaron ',generacion2])
    
    grafica.append(['',''])
    grafica.append(['Total Matricula 0',len(total_matri0)])
    grafica.append(['Estudiantes No Participaron '+arg[5]+' '+str(arg[2]),matricula1])
    grafica.append(['Estudiantes Reprobaron ',matricula2])
    
    prom_total.append(total_gen+total_matri0+prom_bene)    
    total_sin_bene=total_estudiantes-len(total_matri0)-len(total_gen)
    apro_sin=Aprobado-apro_Gen-apro_Matri0
    repo_sin=Reprobaron-generacion2-matricula2
    ceros_sin=Ceros-generacion1-matricula1
    print(f'\nPromedio sin bene {total_sin_bene}\n',
          f'Promedio total sin bene {prom_total}\n',
          f'Aprobacion sin bene {apro_sin}\n',
          f'Reprobado sin bene {repo_sin}\n',
          f'Ceros sin bene {ceros_sin}\n')
    
    grafica.append(['',''])
    grafica.append(['Mas de un intento',''])
    grafica.append(['Descripción de estudiantes','# de estudiantes','# Aprobados','% Aprobados',
                    '# Reprobados','% Reprobados','# Ceros','% Ceros','Promedio de calificación'])
    if total_gen!=[] and total_gen[0]!='**********':
        grafica.append(['Estudiantes de Generación E',len(total_gen),apro_Gen,(apro_Gen/len(total_gen))*100,
                        generacion2,(generacion2/len(total_gen))*100,generacion1,(generacion1/len(total_gen))*100,
                        statistics.mean(total_gen)])
    else:
        grafica.append(['Estudiantes de Generación E','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])
    if total_matri0!=[] and total_matri0[0]!='**********':
        grafica.append(['Estudiantes de matricula cero',len(total_matri0),apro_Matri0,(apro_Matri0/len(total_matri0))*100,
                        matricula2,(matricula2/len(total_matri0))*100,matricula1,(matricula1/len(total_matri0))*100,
                        statistics.mean(total_matri0)])
    else:
        grafica.append(['Estudiantes de matricula cero','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])
    grafica.append(['Estudiantes sin beneficios',total_sin_bene,apro_sin,(apro_sin/total_sin_bene)*100,                    
                    repo_sin,(repo_sin/total_sin_bene)*100,ceros_sin,(ceros_sin/total_sin_bene)*100,
                    statistics.mean(prom_bene)])
    grafica.append(['Total de estudiantes',total_estudiantes,Aprobado,(Aprobado/(total_estudiantes))*100,
                    Reprobaron,(Reprobaron/(total_estudiantes))*100,Ceros,(Ceros/(total_estudiantes))*100,
                    statistics.mean(prom_total[0])])
    
    
    prom_total2.append(total_gen2+total_matri02+prom_bene2)
    total_sin_bene2=total_estudiantes2-len(total_matri02)-len(total_gen2)
    print(f'sin beneficio {total_sin_bene2}\n',
          f'total estudiantes: {total_estudiantes2}')
    
    apro_sin2=Aprobado2-apro_Gen2-apro_Matri02
    repo_sin2=Reprobaron2-generacion22-matricula22
    ceros_sin2=Ceros2-generacion12-matricula12
    
    
    grafica.append(['',''])
    grafica.append(['Descripción de estudiantes','# de estudiantes','# Aprobados','% Aprobados',
                    '# Reprobados','% Reprobados','# Ceros','% Ceros','Promedio de calificación'])
    if total_gen2!=[] and total_gen2[0]!='**********' and total_sin_bene2 != 0:
        grafica.append(['Estudiantes de Generación E',len(total_gen2),apro_Gen2,(apro_Gen2/len(total_gen2))*100,
                        generacion22,(generacion22/len(total_gen2))*100,generacion12,(generacion12/len(total_gen2))*100,
                        statistics.mean(total_gen2)])
    else:
        grafica.append(['Estudiantes de Generación E','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])
    if total_matri02!=[] and total_matri02[0]!='**********':
        grafica.append(['Estudiantes de matricula cero',len(total_matri02),apro_Matri02,(apro_Matri02/len(total_matri02))*100,
                        matricula22,(matricula22/len(total_matri02))*100,matricula12,(matricula12/len(total_matri02))*100,
                        statistics.mean(total_matri02)])
    else:
        grafica.append(['Estudiantes de matricula cero','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])
    if total_sin_bene2 >0:
        grafica.append(['Estudiantes sin beneficios',total_sin_bene2,apro_sin2,(apro_sin2/total_sin_bene2)*100,                    
                        repo_sin2,(repo_sin2/total_sin_bene2)*100,ceros_sin2,(ceros_sin2/total_sin_bene2)*100,
                        statistics.mean(prom_bene2)])
    else:
        grafica.append(['Estudiantes de matricula cero','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])
    if total_sin_bene2 >0:
        grafica.append(['Total de estudiantes',total_estudiantes2,Aprobado2,(Aprobado2/(total_estudiantes2))*100,
                        Reprobaron2,(Reprobaron2/(total_estudiantes2))*100,Ceros2,(Ceros2/(total_estudiantes2))*100,
                        statistics.mean(prom_total2[0])])
    else:
        grafica.append(['Estudiantes de matricula cero','N.A','N.A','N.A','N.A','N.A','N.A','N.A','N.A'])

    
    est = pd.DataFrame (productos)
    po = pd.DataFrame (Program)
    ce = pd.DataFrame (Center)
    gra = pd.DataFrame (grafica)
    zo = pd.DataFrame (Zone)
    te =pd.DataFrame (inten)
    condi =pd.DataFrame (condiciones)
    
    
    try:
        
        with pd.ExcelWriter(nombre.replace(".xlsx", "") +' Reporte '+str(arg[2])+'.xlsx') as writer:  
            est.to_excel(writer, sheet_name='Estudiantes',header=['Cedula','Estudiante','Correo'],index=False)
            gra.to_excel(writer, sheet_name='Grafica',header=False,index=False)
            ce.to_excel(writer, sheet_name='Centros',header=['Centros','Ceros','Reprobaron','Aprobaron'],index=False)
            po.to_excel(writer, sheet_name='Programa',header=['Programa','Ceros','Reprobaron','Aprobaron'],index=False)
            te.to_excel(writer, sheet_name='Intentos',header=['Intentos','Ceros','Reprobaron','Aprobaron','Total'],index=False)
            zo.to_excel(writer, sheet_name='Zonas',header=['Zonas','Ceros','Reprobaron','Aprobaron','Total'],index=False)
            condi.to_excel(writer, sheet_name='Condiciones',header=['Cedula','Estudiante','Correo','Condicion especial'],index=False)
    except:        
        salida = '\nNo se puedo crear el '+'Reporte '+str(arg[2])+'.xlsx\n'
        return salida
    finally:   
        return ('El Reporte '+str(arg[2])+' Se creo Exitosamente')
        
        