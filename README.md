#Record of neighborhood acts

##Uso de Formato base para registro de actas de vecindad

######Este aplicativo facilita armar un documento en Word para realizar el registro de un acta de vecindad, el cual hace:

1. Armar los títulos del proyecto, sea del cliente y de la obra
2. Armar la descripción del acta NOTA: Esta parte es requerida revisar si el elemento no es un Apartamento, ya se está diseñado únicamente para estos elementos
3. Ingresar tablas con descripción de la particularidad que se desea aclarar
4. Ingresa las imágenes con el tamaño adecuado al caso (en este caso es de 5.18cm x 8.5cm)
5. Por cantidad de caracteres añade un círculo y una flecha para ubicar la particularidad si es el caso
6. Añade en la tabla tiempos para hacer la referencia de la ubicación de la particularidad en el video

Debido al poco apoyo de la empresa con esta iniciativa, no se ha seguido avanzando en el desarrollo de este para volverlo en un aplicativo común en el desarrollo de las actividades de la empresa.

##Para su uso se deben de seguir los siguientes pasos:

1. Personaliza el Word con el formato requerido por la empresa, sea marca de agua, pie de pagina, etc.
2. Ingresa al código desde el apartado de programación -> Visual Basic. Busca la función Generar_documento, en las siguientes variables ingresa el valor según corresponda:

	`conjunto` = "Nombre del conjunto donde se está desarrollando el acta, si no es el caso, dejar vacío"
	`direccion` = "Dirección de donde se encuentra el acta, sea dirección de una casa o del conjunto"
	`municipio` = "Ciudad o municipio donde se realiza el acta"

4. Ingresa a Archivo -> Opciones -> Avanzadas -> Tamaño y calidad de la imgen, dejar las casillas "Descartar datos de edición" y "No comprimir las imagenes del archivo" sin selección, y en "Resolución predeterminada" en 150ppp. Esto se realiza para que el documento final de Word no tenga un tamaño excesivo y sea manejable a nivel de recursos del PC y sea mas facil de transferir
3. Guarda este archivo y utiliza este documento como base para copiar y realizar todas las actas del proyecto
5. Organiza las fotos numeradas desde el 1 hasta la cantidad deseada
5. En el archivo "Plantilla de array.xlsx" ingresa la información de la siguiente manera en las columnas:

	`F_Ini` = Foto inicial que contiene la descripción
	`F_Fin` = Foto final que contiene la descripción
	`Descripción` = Añade la descripción correspondiente de la particularidad o del área a inspeccionar

7. Verificar en la columna Revisión si aparece error, esto verifica que el orden de las fotos sea lógico

   - **NOTA 1:** Tener en cuenta que esta verificación puede aparecer como error en las 2 ultimas filas de la descripción, esto dado a que las validaciones dependen de estas 2 ultimas filas
   - **NOTA 2:** Si la descripción tiene una sola foto, esta provoca error, sin embargo, esto es válido en la ejecución del programa

8. Se debe ir a la última fila, en la celda C19, la cual contiene toda la información concatenada, esta se debe copiar y pegar en una celda cualquiera como TEXTO, de lo contrario el código no funciona

   - **NOTA 1:** Si la cantidad de descripción supera las filas dispuestas para esto, estas se deben concatenar de manera manual, editando la última celda y añadiendo los elementos nuevos dispuestos en la columna E seguido del carácter "&"
   - **NOTA 2:** Si la cantidad de filas es excesiva, se debe añadir más de dos descripciones por línea, para esto se edita la celda E2, y se formula de la siguiente manera: 
	`=+A2&","&B2&","""&C2&""","`
Esto se hace para eliminar el salto de página y el guion bajo que usa la sintaxis de VBA para este uso, esto se debe copiar en líneas intercaladas para hacer esta acción

9. Copiar la INFORMACIÓN del paso 8 se debe copiar y pegar en la variable "registro()" y eliminar el salto de página si se genera y la última coma del arreglo
   - **NOTA 1:** Si copia la celda y no la información, entonces la información se copiará con comillas dobles y generará un error en el código
   - **NOTA 2:** Si genera error por exceso de líneas, ejecutar lo explicado en la NOTA 2 del paso 8
10. Llenar los datos del acta, de la siguiente manera:

	`fecha` = "Fecha en la cual se realizó el acta en campo"
	`apto` = "Numero del apartamento del acta, después de la numeración se puede añadir la torre si es el caso

11. Ejecutar el código
12. Revisar el documento generado, arreglando saltos de página, adaptando tamaños de imágenes y acomodando los círculos y flechas según corresponda
13. Añadir los tiempos correspondientes del video registrado
14. Guardar el archivo generado como un documento en Word, para retirar la macro

Comprendo que son demasiados pasos y que el usuario final no tiene porqué interactuar con el código, pero dado el poco interés de la empresa por desarrollar esta iniciativa, este se realiza de forma independiente, con tal de terminar rápidamente una tarea repetitiva y entregar un producto de calidad a la empresa y al cliente

Este proyecto se desarrolló por interés propio y para procesar grandes cantidades de información rápidamente

##Limitaciones de esta aplicación:

1. Interacción directa del usuario con el código
2. Carpeta llamada "FOTOS" no debe cambiar de nombre o no funcionará
3. Uso de aplicaciones externas (Excel)
4. Asignación de minutos en tiempos de video son aproximadas, rangos dados según experiencia
5. Columna de Revisión limitada

##Función Renombrar_fotos

Esta se usa para renombrar más facilmente las fotos, se diseñó exclusivamente para fotos que son generadas con nombre consecutivo por una camara fotográfica, por ende las fotos de celular no lo soporta. Es recomendable tener conocimientos en programación para utilizar esta herramienta

Para ejecutar se deben seguir los siguientes pasos:

1. Ingresar el nombre de la foto inicial en la variable `Ainicial`
2. Ingresar el nombre de la foto final en la variable `Afinal`
3. Si el inicio de la foto es de caracteres distintos a `DSC`, se debe configurar estos en la linea 15
4. Ejecutar

   - **NOTA:** Esta opción se debe usar con cuidado, ya que mal ejecutada puede cambiar de forma permanente el nombre de sus fotos, se recomienda usar con BackUp de las fotos
