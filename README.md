# Universidad Nacional y a Distancia (UNAD)

Este código permite a los docentes de la UNAD obtener características claves de los estudiantes que están en el curso, tales como la tasa de aprobación y reprobados por zonas y centro que hay en el país.

### Momentos.py

Crea un documento en excel con 5 pestañas.

- Pestaña 1 (Estudiantes): muestra la cédula, el nombre y el correo electrónico de los estudiantes que perdieron la actividad.

- Pestaña 2 (Grafica): muestra el total de estudiantes del curso, los estudiantes que aprobaron la etapa, los estudiantes que perdieron la etapa, los estudiantes que no participaron, el porcentaje de aprobación, el porcentaje de los que no participaron, el porcentaje de los estudiantes que perdieron, los estudiantes que generación E que no participaron y lo que reprobaron la actividad.

- Pestaña 3 (Centros): muestra por cada uno de los centros de la UNAD cuantos estudiantes obtuvieron cero y los que reprobaron la actividad.

- Pestaña 4 (Programa): muestra por cada uno de los programas cuantos estudiantes obtuvieron ceros y cuantos reprobaron.

- Pestaña 5 (Intentos): muestra los estudiantes que están repitiendo el curso, cuantas veces ha repetido y cuantos obtuvieron cero en la etapa.

- Pestaña 6 (Zonas): muestra por zonas cuantos estudiantes obtuvieron cero y cuantos reprobaron.

### Mensajes.py

Toma como valor de entrada el documento exportado con el código Momentos.py y envía a los estudiantes que perdieron o no entregaron la actividad un mensaje a su correo institucional, donde se le invita a entregar sus actividades pendientes.

### credenciales.py

El correo electrónico del que va a enviar los correos se agrega en la función usuario()

```python
def usuario():
    correo='CORREO ELECTRONICO'
    return correo
```

La contraseña del correo electrónico escogido se ingresa en la función clave()

```python
def clave():
    contra='CONTRASEÑA'
    return contra
```
-------------
<div style="text-align:center"><img src ="https://github.com/lokocristian/Momentos_UNAD/blob/main/icono.webp" /></div>
