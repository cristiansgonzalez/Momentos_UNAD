# UNAD

### Momentos.py

este codigo permite a los docentes de la UNAD obtener caracteristicas claves de los estudiantes que estan en el curso
Se crea una documento en excel con 5 pestañas.

- Pestaña 1 (Estudiantes): muestra la cedula, el nombre y el correo de los estudiantes que perdieron la actividad.

- Pestaña 2 (Grafica): muestra el total de estudiantes del curso, los estudiantes que aprobaron la etapa, los estudiantes que perdieron la estapa, los estudiantes que no participaron, el porcentaje de aprobacion, el porcentaje de los que no participaron, el porcentaje de los estudiantes que perdieron, los estudiantes que generacion E que no participaron y lo que reprobaron la actividad.

- Pestaña 3 (Centros): muestra por cada uno de los centros de la unad cuantos estudiantes sacaron Cero y los que reprobaron la actividad.

- Pestaña 4 (Programa): muestra los programas que tiene el curso junto con los estudiantes que sacaron ceros y reprobaron en cada uno de los programas


- Pestaña 5 (Intentos): muestra los ceros de los estudiantes que estan repitiendo el curso y cuantas veces ha repetido el curso.

- Pestaña 6 (Zonas): muestra por zonas cuantos estudiantes sacaron cero y cuantos reprobaron.

### mensajes.py

Toma como valor de entrada el documento exportado con el condico Momentos.py 
Envia a los estudiantes que perdieron o no entregaron la actividad un mensaje a su correo institucional, donde se le invita a entregar sus actividades pendientes.

La informacion del correo en el que se va a enviar los correos a los estudiantes se colocara en el archivo credenciales.py
