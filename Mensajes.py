# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 15:23:42 2022

@author: Cristian Gonzalez
"""

import smtplib
import Credenciales
import pandas as pd

Etapa='1'
periodo='16-02 2022'
curso='Sistemas Dinámicos'
plazo='28 de Abril de 2022'
docente='Christian Saúl González Santos'
skype='cristian-saul-66'
rol='Tutor'

archivo = 'Reporte 1 dinamicos.xlsx'
df = pd.read_excel(archivo, sheet_name='Estudiantes')
fil ,col=df.shape

asunto=('Entrega pendiente Etapa ', Etapa,' ',curso,' ',periodo)
asunto=''.join(asunto)

servidor= smtplib.SMTP('smtp.gmail.com',587)
servidor.starttls()
servidor.login(Credenciales.usuario(),Credenciales.clave())
con=1

for X in range(0, fil):
    nombres=[]
    estudiante=df.iloc[X, 2]
    aux=df.iloc[X, 1].split()
    if len(aux)==2:
        nombres.append(aux[0].capitalize())
        nombres.append(aux[1].capitalize())
        mensaje=(f'Estimado Estudiante. {nombres[0]} {nombres[1]} \n\nMe comunico con el fin de invitarle ',
                 'a realizar la entrega correspondiente del desarrollo de las actividades planteadas ',
                 f'en la Etapa {Etapa} del curso {curso} del periodo académico {periodo} o si deseas ',
                 'mejorar la calificación que obtuviste. \n\nEl plazo para la entrega es el próximo ',
                 f'{plazo} por medio del correo interno del curso antes de las 23:55 horas directamente ',
                 f'conmigo. \n\nQuedo atento a sus comentarios.\n\nCordialmente.\n\n{docente}',
                 f'\nSkype: {skype}\n{rol}')
    else:
        nombres.append(aux[0].capitalize())
        nombres.append(aux[1].capitalize())
        nombres.append(aux[2].capitalize())
        mensaje=(f'Estimado Estudiante. {nombres[0]} {nombres[1]} {nombres[2]}\n\nMe comunico con el fin de invitarle ',
                 'a realizar la entrega correspondiente del desarrollo de las actividades planteadas ',
                 f'en la Etapa {Etapa} del curso {curso} del periodo académico {periodo} o si deseas ',
                 'mejorar la calificación que obtuviste. \n\nEl plazo para la entrega es el próximo ',
                 f'{plazo} por medio del correo interno del curso antes de las 23:55 horas directamente ',
                 f'conmigo. \n\nQuedo atento a sus comentarios.\n\nCordialmente.\n\n{docente}',
                 f'\nSkype: {skype}\n{rol}')
    mensaje=''.join(mensaje)
    mensaje='Subject: {}\n\n{}'.format(asunto,mensaje)
    servidor.sendmail(Credenciales.usuario(),estudiante,mensaje.encode('latin-1'))
    print(X,'  ',estudiante)
    con+=1

servidor.quit()

print('\n',X+1,' mensajes se enviaron exitosamente')
