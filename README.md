# Universidad Nacional y a Distancia (UNAD)

Este código permite a los docentes de la UNAD obtener características claves de los estudiantes del curso, tales como número de estudiantes que reprobaron y no presentaron la actividad con sus debidos porcentajes    

## Principal.py

Los parámetros de entrada son 4:

1. Descargar el excel desde el centralizador de la UNAD, guardarlo en la misma carpeta en donde descargo el `Principal.py` y colocar el nombre del documento que descargo en la siguiente línea de código.

```python
archivo = 'NOMBRE_DEL_EXCEL'
```

2. Ingresar cuantos puntos vale la actividad que se quiere analizar.

```python
puntaje_actividad = 25
```

3. Ingresar el numero de actividad.

```python
Etapa = 1
```

4. Ingresar el numero total de actividad que tiene el curso.

```python
Etapas_totales = 5
```

Se creara un documento en excel que se llama `Reporte #.xlsx` (# indica que actividad fue analizada) con 2 pestañas.

- Pestaña 1 (Estudiantes): muestra la cédula, el nombre y el correo electrónico de los estudiantes que no presentaron y que la perdieron la actividad.

- Pestaña 2 (Grafica): muestra el total de estudiantes del curso, los estudiantes que aprobaron la etapa, los estudiantes que perdieron la etapa, los estudiantes que no participaron, el porcentaje de aprobación, el porcentaje de los que no participaron, el porcentaje de los estudiantes que perdieron.

- Pestaña 3 (Centros): muestra por cada uno de los centros de la UNAD cuantos estudiantes obtuvieron cero y los que reprobaron la actividad.

- Pestaña 4 (Programa): muestra por cada uno de los programas cuantos estudiantes obtuvieron ceros, cuantos reprobaron y cuantos aprobaron.

- Pestaña 5 (Intentos): muestra los estudiantes que están repitiendo el curso, cuantas veces ha repetido y cuantos obtuvieron cero en la etapa.

- Pestaña 6 (Zonas): muestra por zonas cuantos estudiantes obtuvieron cer, cuantos reprobaron y cuantos aprobaron.

## Mensajes.py

Toma como valor de entrada el documento exportado con el código `Principal.py` y les envía a los estudiantes que perdieron o no entregaron la actividad un mensaje a su correo institucional, donde se le invita a entregar su actividad pendiente.

## Credenciales.py

El correo electrónico del que va a enviar los correos se agrega en la función `usuario()`

```python
def usuario():
    correo='CORREO_ELECTRONICO'
    return correo
```

La contraseña del correo electrónico escogido se ingresa en la función `clave()`

```python
def clave():
    contra='CONTRASEÑA'
    return contra
```


## Autor

- Cristian González (<cristian-saul-66@hotmail.com>)
