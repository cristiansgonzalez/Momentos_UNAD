# -*- coding: utf-8 -*-
"""
Created on Sat Mar 26 14:01:34 2022

@author: Cristian Gonzalez
"""

import smtplib
import credenciales
import pandas as pd

Etapa='1'
periodo='16-01 2022'
curso='Sistemas Dinámicos'
plazo='24 de Marzo de 2022'
docente='Christian Saul Gonzalez Santos'
skype='cristian-saul-66'
rol='Tutor'

archivo = 'Sabana_de_notas_243005_1141.csv'

df = pd.read_csv(archivo,encoding="latin-1")
df=df.to_string().split(";")
df=df[16::]
fil=int(len(df)/11)

ini_correo=6
ini_nombres=2
estudiante=[]
con=0

# inicializar correo electronico
asunto=('Alerta Temprana ', curso)
asunto=''.join(asunto)
servidor= smtplib.SMTP('smtp.gmail.com',587)
servidor.starttls()
servidor.login(credenciales.usuario(),credenciales.clave())

for X in range(0,fil):
    #estudiante.append(df[ini_correo+X*11])
    nombres=[]
    estudiante=df[ini_correo+X*11]
    aux=(df[ini_nombres+X*11].split())
    nombres.append(aux[0].capitalize())
    nombres.append(aux[1].capitalize())
    nombres.append(aux[2].capitalize())
    
    mensaje=(f'Estimado Estudiante. {nombres[0]} {nombres[1]} {nombres[2]}\n\nRecibe un cordial ',
             f'saludo del curso "{curso.upper()}". Escribo para motivarte a que ingreses, ',
             'realices la identificación de la agenda y la guía del curso. En este ',
             f'momento estamos ejecutando la Etapa {Etapa} del curso, puedes comunicarte ',
             'conmigo por este medio, por correo interno del aula o a través de skype.\n\n',
             'Agradezco y quedo atenta a cualquier duda.',
             f'\n\nCordialmente.\n\n{docente}',
             f'\nSkype: {skype}\n{rol}')
    
    mensaje=''.join(mensaje)
    mensaje='Subject: {}\n\n{}'.format(asunto,mensaje)
    servidor.sendmail(credenciales.usuario(),estudiante,mensaje.encode('latin-1'))
    print(X,'  ',estudiante)
    con+=1

servidor.quit()

print('\n',X+1,' mensajes se enviaron exitosamente')






    