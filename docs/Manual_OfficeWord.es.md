



# Office Word
  
Modulo para crear, abrir y editar documentos .docx  

*Read this in other languages: [English](Manual_OfficeWord.md), [Português](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md)*
  
![banner](imgs/Banner_OfficeWord.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.  


## Descripción de los comandos

### Nuevo documento
  
Crea un nuevo documento word
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|

### Abrir Documento
  
Abre un documento de Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo|Archivo con formato docx que se abrirá|archivo.docx|
|session||session|

### Leer documento
  
Extrae texto de documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variable donde se guardará el texto extraído|Variable|

### Escribir en documento
  
Escribe en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Escriba texto|Texto que será escrito en el documento|Lorem ipsum |
|Fuente de texto|Fuente de texto que se usará en el documento|Arial|
|Tipo de texto|Tipo de texto que será escrito (Titulo, Header 1, Header 2, etc.)|Title|
|Tamaño de fuente|Tamaño de fuente que tendrá el texto escrito|12|
|Alineación|Alineación que tendrá el texto|left|
|Negrita|Casilla para elegir si el texto escrito estará en negrita|False|
|Cursiva|Casilla para elegir si el texto escrito estará en cursiva|True|
|Subrayar|Casilla para elegir si el texto escrito estará subrayado|True|
|session||session|

### Leer tabla
  
Extrae texto de una tabla de un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variable donde se guardará el texto de la tabla|resultado|

### Agregar texto desde un bookmark
  
Agrega texto desde un bookmark a documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese texto|Texto que se agregará|Lorem ipsum|
|Limpiar|Casilla para elegir si el texto anterior será eliminado|True|
|Ingrese bookmark|Nombre del bookmark|Lorem ipsum|
|session||session|

### Cerrar documento
  
Cierra el documento que se está ejecutando
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|

### Insertar página
  
Inserta una nueva página al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |

### Agregar imagen
  
Agrega una imagen al documento
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ruta de la imagen|Ruta de la imagen a agregar en el documento|imagen.jpg|
|Ancho de la imagen|Ancho que tendrá la imagen|600|
|Alto de la imagen|Alto que tendrá la imagen|500|
|session||session|

### Buscar Texto en párrafo
  
Busca el párrafo donde se encuentra el texto indicado.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que se buscará en el documento|Hola mundo|
|session||session|
|Nombre de la variable|Variable donde se guardará el número de párrafo que contiene el texto buscado|Variable|

### Contar párrafos
  
Cuenta la cantidad de párrafos del documento.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|
|Nombre de la variable|Variable donde se guardará la cantidad de párrafos del documento|Variable|

### Obtener parrafos
  
Obtiene un listado de parrafos en forma de diccionario {numero: texto}.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variable donde se guardará el texto extraído|Variable|

### Limpiar parrafo
  
Limpia el contenido de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Numero de parrafo|Posición del parrafo a borrar.|1|
|session||session|
|Resultado|Variable donde se guardará el texto extraído|Variable|

### Agregar párrafo
  
Agrega un parrafo en la posición deseada en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Escriba texto|Texto que será escrito en el documento|Lorem ipsum |
|Numero de parrafo|Posición del nuevo parrafo.|1|
|Fuente de texto|Fuente de texto que se usará en el documento|Arial|
|Tamaño de fuente|Tamaño de fuente que tendrá el texto escrito|12|
|Alineación|Alineación que tendrá el texto|left|
|Negrita|Casilla para elegir si el texto escrito estará en negrita|False|
|Cursiva|Casilla para elegir si el texto escrito estará en cursiva|True|
|Subrayar|Casilla para elegir si el texto escrito estará subrayado|True|
|session||session|

### Agregar texto a párrafo
  
Agrega texto al final de un parrafo en un documento Word.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Escriba texto|Texto que será escrito en el documento|Lorem ipsum |
|Numero de parrafo|Posición del parrafo.|1|
|Fuente de texto|Fuente de texto que se usará en el documento|Arial|
|Tamaño de fuente|Tamaño de fuente que tendrá el texto escrito|12|
|Alineación|Alineación que tendrá el texto|left|
|Negrita|Casilla para elegir si el texto escrito estará en negrita|False|
|Cursiva|Casilla para elegir si el texto escrito estará en cursiva|True|
|Subrayar|Casilla para elegir si el texto escrito estará subrayado|True|
|session||session|

### Remplazar texto en párrafo
  
Remplaza el texto de un párrafo.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Texto a Buscar|Texto que será buscado en el documento|Hola mundo|
|Texto a Reemplazar|Texto a reemplazar en el documento|Hola mundo|
|Lista de párrafo|Lista de párrafos donde se buscará y reemplazará el texto|Separados por comas ',' ejemplo: 1,2|
|session||session|
|Nombre de la variable|Nombre de la variable donde se almacenará el resultado|Variable|

### Convertir a PDF
  
Convierte documento Word a PDF.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Archivo Word|Archivo de word que se convertirá a PDF|archivo.docx|
|Guardar archivo|Nombre y ruta del archivo donde se guardará el archivo generado|archivo.pdf|

### Guardar documento
  
Guarda el documento Word abierto
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|session||session|
|Guardar archivo|Guarda el archivo con el nombre y la ruta especificada|archivo.docx|
