# Office Word
  
Módulo para trabalhar com arquivos docx  

*Read this in other languages: [English](Manual_OfficeWord.md), [Portugues](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md).*
  
![banner](/docs/imgs/Banner_OfficeWord.png)

## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  



## Descrição do comando

### Novo documento
  
Criar um novo documento do Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
| --- | --- | --- |

### Abrir Documento
  
Abra um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo|arquivo de formato docx para abrir|arquivo.docx|

### Ler documento
  
Extrair texto de um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado|Variável onde o texto extraído será salvo|Variável|

### Salvar documento
  
Extraia o texto do arquivo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Salvar arquivo|Salve o arquivo com o nome e caminho especificados|arquivo.docx|

### Escrever no documento
  
Escreva em um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Escreva texto|Texto a ser escrito no documento|Lorem ipsum |
|Tipo de texto|Tipo de texto a ser escrito (Título, Header 1, Header 2, etc.)|Title|
|Tamanho da fonte|Tamanho da fonte que o texto escrito terá|12|
|Alinhamento|Alinhamento que o texto terá|left|
|Negrito|Caixa de seleção para escolher se o texto escrito ficará em negrito|False|
|Itálico|Caixa de seleção para escolher se o texto escrito ficará em itálico|True|
|Sublinhado|Caixa de seleção para escolher se o texto escrito será sublinhado|True|

### Ler Tabela
  
Extrair o texto da tabela de um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado|Variável onde o texto da tabela será salvo|resultado|

### Adicionar texto do marcador
  
Adicione texto de um marcador a um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Adicione texto|Texto a ser adicionado|Lorem ipsum|
|Limpar|Caixa de seleção para escolher se o texto anterior será excluído|True|
|Digite marcador|Nome do marcador|Lorem ipsum|

### Fechar documento
  
Feche o documento que está sendo executado
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
| --- | --- | --- |

### Adicionar Página
  
Adicionar uma nova página ao documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
| --- | --- | --- |

### Adicionar imagem
  
Adicione uma imagem ao documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho da imagem|Caminho da imagem a ser adicionada ao documento|imagem.jpg|
|Anchura da imagem|Anchura que a imagem terá|600|
|Altura da imagem|Altura que a imagem terá|500|

### Converter para PDF
  
Converter documento do Word para PDF.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo word|Arquivo do Word a ser convertido para PDF|arquivo.docx|
|Salvar arquivo|Nome e caminho do arquivo onde o arquivo gerado será salvo|arquivo.pdf|

### Localizar texto no parágrafo
  
Localize em qual parágrafo há um texto indicado.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Texto para pesquisar|Texto a pesquisar no documento|Olá Mundo|
|Nome variável|Variável onde será salvo o número do parágrafo que contém o texto pesquisado|Variável|

### Contar parágrafos
  
Contar o número de parágrafos no documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Nome variável|Variável onde será salvo o número de parágrafos do documento|Variável|

### Substituir texto no parágrafo
  
Substituir o texto de um parágrafo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Texto para pesquisar|Texto a ser pesquisado no documento|Olá mundo|
|Texto a substituir|Texto a substituir no documento|Olá mundo|
|Números de parágrafos|Lista de parágrafos onde o texto será encontrado e substituído|Separados por vírgulas ',' exemplo: 1,2|
|Nome Varíavel|Nome da variável onde o resultado será armazenado|Variável|
