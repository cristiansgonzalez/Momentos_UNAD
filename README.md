# Universidad Nacional y a Distancia (UNAD)

Este código permite a los docentes de la UNAD obtener características claves de los estudiantes que están en el curso, tales como la tasa de aprobación y reprobados por zonas y centro que hay en el país.

## Analisis_estudiantes.py

Como parámetro de entrada se debe ingresar el nombre del documento en excel, el cual se descarga desde el centralizador de la UNAD y agregarlo en la siguiente línea de código.

```python
archivo = 'NOMBRE_DEL_EXCEL.xlsx'
```

Se creara un documento en excel que se llama `Momento 1.xlsx`con 2 pestañas.

- Pestaña 1 (Estudiantes): muestra la cédula, el nombre y el correo electrónico de los estudiantes que perdieron la actividad.

- Pestaña 2 (Grafica): muestra el total de estudiantes del curso, los estudiantes que aprobaron la etapa, los estudiantes que perdieron la etapa, los estudiantes que no participaron, el porcentaje de aprobación, el porcentaje de los que no participaron, el porcentaje de los estudiantes que perdieron.

## Mensajes.py

Toma como valor de entrada el documento exportado con el código `Analisis_estudiantes.py` y envía a los estudiantes que perdieron o no entregaron la actividad un mensaje a su correo institucional, donde se le invita a entregar sus actividades pendientes.

## credenciales.py

El correo electrónico del que va a enviar los correos se agrega en la función usuario()

```python
def usuario():
    correo='CORREO_ELECTRONICO'
    return correo
```

La contraseña del correo electrónico escogido se ingresa en la función clave()

```python
def clave():
    contra='CONTRASEÑA'
    return contra
```


## Autor

- Cristian González (<cristian-saul-66@hotmail.com>)

