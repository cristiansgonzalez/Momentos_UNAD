# -*- coding: utf-8 -*-
"""
Created on Sat Apr  2 08:12:14 2022

@author: Cristian Gonzalez
"""

import Inicializar

'''
archivo= es el nombre del documento en excel 
puntaje_actividad= cuantos puntos vale la actividad
Etapa= Etapa que se va a analizar
Etapas_totales= cantidad total de estapas
'''


archivo= 'estudiantes'
puntaje_actividad= 300
Etapa= 2
Etapas_totales= 5
Generacion_E='si'

#arg=[archivo,puntaje_actividad,Etapa,Etapas_totales,Generacion_E]

Inicializar.inicio(archivo,puntaje_actividad,Etapa,Etapas_totales,Generacion_E)
