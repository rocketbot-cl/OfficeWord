# Office Word
  
Module to work with docx files
  
*Read this in other languages: [English](Manual_OfficeWord.md), [Portugues](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md).*
  
![banner](/docs/imgs/Banner_OfficeWord.png)

## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### New Document
  
Create a new word document
|Parameters|Description|example|
| --- | --- | --- |
| --- | --- | --- |

### Open Document
  
Open a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|File|docx format file to open|file.docx|

### Read Document
  
Extract text from a Word document
|Parameters|Description|example|
| --- | --- | --- |
|Result|Variable where the extracted text will be saved|Variable|

### Save document
  
Extract text from file.
|Parameters|Description|example|
| --- | --- | --- |
|Save file|Save the file with the specified name and path|file.docx|

### Write in Document
  
Write in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Write text|Text to be written on the document|Lorem ipsum |
|Text type|Type of text to be written (Title, Header 1, Header 2, etc.)|Title|
|Font size|Font size that the written text will have|12|
|Alignment|Alignment that the text will have|left|
|Bold|Checkbox to choose if the written text will be in bold|False|
|Italic|Checkbox to choose if the written text will be in italics|True|
|Underline|Checkbox to choose if the written text will be underlined|True|

### Read Table
  
Extract table text from a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Result|Variable where the text of the table will be saved|result|

### Add text from bookmark
  
Add text from a bookmark to Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Add text|Text to be added|Lorem ipsum|
|Clean|Checkbox to choose if the previous text will be deleted|True|
|Bookmark|Bookmark name|Lorem ipsum|

### Close document
  
Close the document that is running
|Parameters|Description|example|
| --- | --- | --- |
| --- | --- | --- |

### Add Page
  
Add a new page to the document
|Parameters|Description|example|
| --- | --- | --- |
| --- | --- | --- |

### Add Picture
  
Add an image to the document.
|Parameters|Description|example|
| --- | --- | --- |
|Image path|Path of the image to add in the document|image.jpg|
|Image width|Width that the image will have|600|
|Image height|Height that the image will have|500|

### Convert to PDF
  
Convert Word document to PDF.
|Parameters|Description|example|
| --- | --- | --- |
|Word file|Word file to be converted to PDF|file.docx|
|Save file|Name and path of the file where the generated file will be saved|file.pdf|

### Locate Text in Paragraph
  
Locate in which paragraph there is an indicated text.
|Parameters|Description|example|
| --- | --- | --- |
|Text to Search|Text to search for in the document|Hello Word|
|variable name|Variable where the paragraph number containing the searched text will be saved|Variable|

### Count Paragraphs
  
Count the number of paragraphs in the document.
|Parameters|Description|example|
| --- | --- | --- |
|Variable name|Variable where the number of paragraphs of the document will be saved|Variable|

### Replace text in paragraph
  
Replace the text of a paragraph.
|Parameters|Description|example|
| --- | --- | --- |
|Text to Search|Text to be searched for in the document|Hello Word|
|Text to replace|Text to replace in the document|Hello Word|
|Paragraph numbers|List of paragraphs where text will be found and replaced|Comma separated ',' example: 1,2|
|Variable name|Name of the variable where the result will be stored|Variable|
