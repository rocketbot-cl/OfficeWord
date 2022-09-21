# Office Word
  
Modulo para trabajar con archivos docx  

*Read this in other languages: [English](Manual_OfficeWord.md), [Portugues](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md).*
  
![banner](imgs/Banner_OfficeWord.png)

## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Nuevo documento
  
Crea un nuevo documento word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
| --- | --- | --- |

### Abrir Documento
  
Abre un documento de Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo|Archivo con formato docx que se abrirá|archivo.docx|

### Leer documento
  
Extrae texto de documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Variable donde se guardará el texto extraído|Variable|

### Guardar documento
  
Guarda el documento Word abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Guardar archivo|Guarda el archivo con el nombre y la ruta especificada|archivo.docx|

### Escribir en documento
  
Escribe en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Escriba texto|Texto que será escrito en el documento|Lorem ipsum |
|Fuente de texto|Fuente de texto que se usará en el documento|Arial |
|Tipo de texto|Tipo de texto que será escrito (Titulo, Header 1, Header 2, etc.)|Title|
|Tamaño de fuente|Tamaño de fuente que tendrá el texto escrito|12|
|Alineación|Alineación que tendrá el texto|left|
|Negrita|Casilla para elegir si el texto escrito estará en negrita|False|
|Cursiva|Casilla para elegir si el texto escrito estará en cursiva|True|
|Subrayar|Casilla para elegir si el texto escrito estará subrayado|True|

### Leer tabla
  
Extrae texto de una tabla de un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Resultado|Variable donde se guardará el texto de la tabla|resultado|

### Agregar texto desde un bookmark
  
Agrega texto desde un bookmark a documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese texto|Texto que se agregará|Lorem ipsum|
|Limpiar|Casilla para elegir si el texto anterior será eliminado|True|
|Ingrese bookmark|Nombre del bookmark|Lorem ipsum|

### Cerrar documento
  
Cierra el documento que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
| --- | --- | --- |

### Insertar página
  
Inserta una nueva página al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
| --- | --- | --- |

### Agregar imagen
  
Agrega una imagen al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la imagen|Ruta de la imagen a agregar en el documento|imagen.jpg|
|Ancho de la imagen|Ancho que tendrá la imagen|600|
|Alto de la imagen|Alto que tendrá la imagen|500|

### Convertir a PDF
  
Convierte documento Word a PDF.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo Word|Archivo de word que se convertirá a PDF|archivo.docx|
|Guardar archivo|Nombre y ruta del archivo donde se guardará el archivo generado|archivo.pdf|

### Buscar Texto en párrafo
  
Busca el párrafo donde se encuentra el texto indicado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que se buscará en el documento|Hola mundo|
|Nombre de la variable|Variable donde se guardará el número de párrafo que contiene el texto buscado|Variable|

### Contar párrafos
  
Cuenta la cantidad de párrafos del documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Nombre de la variable|Variable donde se guardará la cantidad de párrafos del documento|Variable|

### Remplazar texto en párrafo
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que será buscado en el documento|Hola mundo|
|Texto a Reemplazar|Texto a reemplazar en el documento|Hola mundo|
|Lista de párrafo|Lista de párrafos donde se buscará y reemplazará el texto|Separados por comas ',' ejemplo: 1,2|
|Nombre de la variable|Nombre de la variable donde se almacenará el resultado|Variable|
