# Office Word
  
Module to create, open and edit .docx documents  

*Read this in other languages: [English](Manual_OfficeWord.md), [Português](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md)*
  
![banner](imgs/Banner_OfficeWord.png)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


## Description of the commands

### New Document
  
Create a new word document
|Parameters|Description|example|
| --- | --- | --- |
|session||session|

### Open Document
  
Open a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|File|docx format file to open|file.docx|
|session||session|

### Read Document
  
Extract text from a Word document
|Parameters|Description|example|
| --- | --- | --- |
|session||session|
|Result|Variable where the extracted text will be saved|Variable|

### Write in Document
  
Write in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Write text|Text to be written on the document|Lorem ipsum |
|Text font|Text font that will be used in the document|Arial|
|Text type|Type of text to be written (Title, Header 1, Header 2, etc.)|Title|
|Font size|Font size that the written text will have|12|
|Alignment|Alignment that the text will have|left|
|Bold|Checkbox to choose if the written text will be in bold|False|
|Italic|Checkbox to choose if the written text will be in italics|True|
|Underline|Checkbox to choose if the written text will be underlined|True|
|session||session|

### Read Table
  
Extract table text from a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|session||session|
|Result|Variable where the text of the table will be saved|result|

### Add text from bookmark
  
Add text from a bookmark to Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Add text|Text to be added|Lorem ipsum|
|Clean|Checkbox to choose if the previous text will be deleted|True|
|Bookmark|Bookmark name|Lorem ipsum|
|session||session|

### Close document
  
Close the document that is running
|Parameters|Description|example|
| --- | --- | --- |
|session||session|

### Add Page
  
Add a new page to the document
|Parameters|Description|example|
| --- | --- | --- |

### Add Picture
  
Add an image to the document.
|Parameters|Description|example|
| --- | --- | --- |
|Image path|Path of the image to add in the document|image.jpg|
|Image width|Width that the image will have|600|
|Image height|Height that the image will have|500|
|session||session|

### Locate Text in Paragraph
  
Locate in which paragraph there is an indicated text.
|Parameters|Description|example|
| --- | --- | --- |
|Text to Search|Text to search for in the document|Hello Word|
|session||session|
|variable name|Variable where the paragraph number containing the searched text will be saved|Variable|

### Count Paragraphs
  
Count the number of paragraphs in the document.
|Parameters|Description|example|
| --- | --- | --- |
|session||session|
|Variable name|Variable where the number of paragraphs of the document will be saved|Variable|

### Get Paragraphs
  
Gets a list of paragraphs in the form of a dictionary {number: text}.
|Parameters|Description|example|
| --- | --- | --- |
|session||session|
|Result|Variable where the extracted text will be saved|Variable|

### Clear Paragraph
  
Clears the content of a paragraph.
|Parameters|Description|example|
| --- | --- | --- |
|Paragraph number|Position of the paragraph to delete.|1|
|session||session|
|Result|Variable where the extracted text will be saved|Variable|

### Add paragraph
  
Add a paragraph at the desired position in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Write text|Text to be written on the document|Lorem ipsum |
|Paragraph number|Position of the new paragraph.|1|
|Text font|Text font that will be used in the document|Arial|
|Font size|Font size that the written text will have|12|
|Alignment|Alignment that the text will have|left|
|Bold|Checkbox to choose if the written text will be in bold|False|
|Italic|Checkbox to choose if the written text will be in italics|True|
|Underline|Checkbox to choose if the written text will be underlined|True|
|session||session|

### Add text to paragraph
  
Add text to the end of a paragraph in a Word document.
|Parameters|Description|example|
| --- | --- | --- |
|Write text|Text to be written on the document|Lorem ipsum |
|Paragraph number|Position of the paragraph.|1|
|Text font|Text font that will be used in the document|Arial|
|Font size|Font size that the written text will have|12|
|Alignment|Alignment that the text will have|left|
|Bold|Checkbox to choose if the written text will be in bold|False|
|Italic|Checkbox to choose if the written text will be in italics|True|
|Underline|Checkbox to choose if the written text will be underlined|True|
|session||session|

### Replace text in paragraph
  
Replace the text of a paragraph.
|Parameters|Description|example|
| --- | --- | --- |
|Text to Search|Text to be searched for in the document|Hello Word|
|Text to replace|Text to replace in the document|Hello Word|
|Paragraph numbers|List of paragraphs where text will be found and replaced|Comma separated ',' example: 1,2|
|session||session|
|Variable name|Name of the variable where the result will be stored|Variable|

### Convert to PDF
  
Convert Word document to PDF.
|Parameters|Description|example|
| --- | --- | --- |
|Word file|Word file to be converted to PDF|file.docx|
|Save file|Name and path of the file where the generated file will be saved|file.pdf|

### Save document
  
Extract text from file.
|Parameters|Description|example|
| --- | --- | --- |
|session||session|
|Save file|Save the file with the specified name and path|file.docx|
