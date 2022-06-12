# -*- coding: utf-8 -*-
"""
Created on Thu Mar 24 15:23:42 2022

@author: Cristian Gonzalez
"""

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import smtplib
import Credenciales
import pandas as pd
import time
mensaje = MIMEMultipart()

ruta_adjunto = 'imagen.jpg'
nombre_adjunto = 'imagen.jpg'

Etapa='1'
periodo='16-02 2022'
curso='Sistemas Dinámicos'
plazo='28 de Abril de 2022'
docente='Christian Saúl González Santos'
skype='cristian-saul-66'
rol='Tutor'

archivo = 'estudiantes Reporte 2.xlsx'
df = pd.read_excel(archivo, sheet_name='Estudiantes')
fil ,col=df.shape

asunto=('Actividad de Nivelación')
asunto=''.join(asunto)

servidor= smtplib.SMTP('smtp.gmail.com',587)
servidor.starttls()
servidor.login(Credenciales.usuario(),Credenciales.clave())
con=1

for X in range(0, fil):
    nombres=[]
    aux3=[]
    estudiante=df.iloc[X, 2]
    aux=df.iloc[X, 1].split()
    for Y in range(0,len(aux)):
        aux3.append(aux[Y].capitalize())
    nombres=' '.join([str(item) for item in aux3])
    
    mensaje=(f'Estimado Estudiante. {nombres} \n\nEn la red del curso de Sistemas Dinámicos',
             'se habló y se estructuro una actividad de nivelación para todos los estudiantes',
             'que quieran mejorar su calificación, en el documento está escrito todas las',
             'indicaciones para que las tendrán presentes, en donde deben enviar la actividad',
             'y la fecha de entrega la cual es para el 3 de julio de 2022, toda la información',
             'está en el documento por favor léanlo a detalle.',
             
             f'\n\nQuedo atento a sus comentarios.\n\nCordialmente.\n\n{docente}',
             f'\nSkype: {skype}\n{rol}')
    
    mensaje=''.join(mensaje)
    mensaje='Subject: {}\n\n{}'.format(asunto,mensaje)
    
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
     
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    mensaje.attach(adjunto_MIME)
    
    texto = mensaje.as_string()
    
    servidor.sendmail(Credenciales.usuario(),estudiante,texto.encode('utf8'))
    print(X,'  ',estudiante)
    con+=1
    time.sleep(2)

servidor.quit()

print('\n',X+1,' mensajes se enviaron exitosamente')
