# -*- coding: utf-8 -*-
"""
Created on Sun Apr  3 22:02:55 2022

@author: Cristian Gonzalez
"""

import Analisis

def inicio(*arg):
    
    try:
        
        Analisis.Analisis_Curso(arg[0], arg[1], arg[2], arg[3])
    except:
        print('Error de compilacion, Verificar los datos ingresados')
        