# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 15:23:42 2022

@author: Cristian Gonzalez
"""

import smtplib
import credenciales
import pandas as pd

Etapa='1'
periodo='16-01 2022'
curso='Sistemas Dinamicos'
plazo='24 de Marzo de 2022'
docente='Christian Saul Gonzalez Santos'
skype='cristian-saul-66'
rol='Tutor'

archivo = 'Momento 1.xlsx'
df = pd.read_excel(archivo, sheet_name='Estudiantes')
fil ,col=df.shape

asunto=('Entregas pendientes Etapa ', Etapa,' ',curso,' ',periodo)
asunto=''.join(asunto)

servidor= smtplib.SMTP('smtp.gmail.com',587)
servidor.starttls()
servidor.login(credenciales.usuario(),credenciales.clave())
con=1

for X in range(0, fil):
    nombres=[]
    estudiante=df.iloc[X, 2]
    aux=df.iloc[X, 1].split()
    nombres.append(aux[0].capitalize())
    nombres.append(aux[1].capitalize())
    nombres.append(aux[2].capitalize())
    
    mensaje=('Estimado ',nombres[0],' ',nombres[1],' ',nombres[2],'\n\nMe comunico con el fin de invitarle a realizar la entrega correspondiente',
             ' del desarrollo de las actividades planteadas en la Etapa ',Etapa,' del curso ',curso,
             ' del periodo academico ',periodo,' o si deseas mejorar la calificacion que obtuviste. \n\n',
             'El plazo para la entrega es el proximo ',plazo,' por medio del correo interno del',
             ' curso antes de las 23:55 horas directamente conmigo. \n\nQuedo atento a sus comentarios.',
             '\n\nCordialmente. \n\n',docente,'\nskype: ',skype,' \n',rol)
    mensaje=''.join(mensaje)
    mensaje='Subject: {}\n\n{}'.format(asunto,mensaje)
    servidor.sendmail(credenciales.usuario(),estudiante,mensaje)
    print(X+1,'  ',estudiante)
    con+=1

servidor.quit()

print(X,' se enviaron exitosamente')
