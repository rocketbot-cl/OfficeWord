



# Office Word
  
Módulo para criar, abrir e editar documentos .docx  

*Read this in other languages: [English](Manual_OfficeWord.md), [Português](Manual_OfficeWord.pr.md), [Español](Manual_OfficeWord.es.md)*
  
![banner](imgs/Banner_OfficeWord.png)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Descrição do comando

### Novo documento
  
Criar um novo documento do Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|

### Abrir Documento
  
Abra um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo|arquivo de formato docx para abrir|arquivo.docx|
|session||session|

### Ler documento
  
Extrair texto de um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variável onde o texto extraído será salvo|Variável|

### Escrever no documento
  
Escreva em um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Escreva texto|Texto a ser escrito no documento|Lorem ipsum |
|Fonte de texto|Fonte de texto que será usada no documento|Arial|
|Tipo de texto|Tipo de texto a ser escrito (Título, Header 1, Header 2, etc.)|Title|
|Tamanho da fonte|Tamanho da fonte que o texto escrito terá|12|
|Alinhamento|Alinhamento que o texto terá|left|
|Negrito|Caixa de seleção para escolher se o texto escrito ficará em negrito|False|
|Itálico|Caixa de seleção para escolher se o texto escrito ficará em itálico|True|
|Sublinhado|Caixa de seleção para escolher se o texto escrito será sublinhado|True|
|session||session|

### Ler Tabela
  
Extrair o texto da tabela de um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variável onde o texto da tabela será salvo|resultado|

### Adicionar texto do marcador
  
Adicione texto de um marcador a um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Adicione texto|Texto a ser adicionado|Lorem ipsum|
|Limpar|Caixa de seleção para escolher se o texto anterior será excluído|True|
|Digite marcador|Nome do marcador|Lorem ipsum|
|session||session|

### Fechar documento
  
Feche o documento que está sendo executado
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|

### Adicionar Página
  
Adicionar uma nova página ao documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |

### Adicionar imagem
  
Adicione uma imagem ao documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Caminho da imagem|Caminho da imagem a ser adicionada ao documento|imagem.jpg|
|Anchura da imagem|Anchura que a imagem terá|600|
|Altura da imagem|Altura que a imagem terá|500|
|session||session|

### Localizar texto no parágrafo
  
Localize em qual parágrafo há um texto indicado.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Texto para pesquisar|Texto a pesquisar no documento|Olá Mundo|
|session||session|
|Nome variável|Variável onde será salvo o número do parágrafo que contém o texto pesquisado|Variável|

### Contar parágrafos
  
Contar o número de parágrafos no documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|
|Nome variável|Variável onde será salvo o número de parágrafos do documento|Variável|

### Obter parágrafos
  
Obtém uma lista de parágrafos na forma de um dicionário {number: text}.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|
|Resultado|Variável onde o texto extraído será salvo|Variável|

### Limpar parágrafo
  
Limpa o conteúdo de um parágrafo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número do parágrafo|Posição do parágrafo a excluir.|1|
|session||session|
|Resultado|Variável onde o texto extraído será salvo|Variável|

### Adicionar parágrafo
  
Adicione um parágrafo na posição desejada em um documento do Word..
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Escreva texto|Texto a ser escrito no documento|Lorem ipsum |
|Número do parágrafo|Posição do novo parágrafo.|1|
|Fonte de texto|Fonte de texto que será usada no documento|Arial|
|Tamanho da fonte|Tamanho da fonte que o texto escrito terá|12|
|Alinhamento|Alinhamento que o texto terá|left|
|Negrito|Caixa de seleção para escolher se o texto escrito ficará em negrito|False|
|Itálico|Caixa de seleção para escolher se o texto escrito ficará em itálico|True|
|Sublinhado|Caixa de seleção para escolher se o texto escrito será sublinhado|True|
|session||session|

### Adicionar texto ao parágrafo
  
Adicione texto ao final de um parágrafo em um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Escreva texto|Texto a ser escrito no documento|Lorem ipsum |
|Número do parágrafo|Posição do parágrafo.|1|
|Fonte de texto|Fonte de texto que será usada no documento|Arial|
|Tamanho da fonte|Tamanho da fonte que o texto escrito terá|12|
|Alinhamento|Alinhamento que o texto terá|left|
|Negrito|Caixa de seleção para escolher se o texto escrito ficará em negrito|False|
|Itálico|Caixa de seleção para escolher se o texto escrito ficará em itálico|True|
|Sublinhado|Caixa de seleção para escolher se o texto escrito será sublinhado|True|
|session||session|

### Substituir texto no parágrafo
  
Substituir o texto de um parágrafo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Texto para pesquisar|Texto a ser pesquisado no documento|Olá mundo|
|Texto a substituir|Texto a substituir no documento|Olá mundo|
|Números de parágrafos|Lista de parágrafos onde o texto será encontrado e substituído|Separados por vírgulas ',' exemplo: 1,2|
|session||session|
|Nome Varíavel|Nome da variável onde o resultado será armazenado|Variável|

### Converter para PDF
  
Converter documento do Word para PDF.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo word|Arquivo do Word a ser convertido para PDF|arquivo.docx|
|Salvar arquivo|Nome e caminho do arquivo onde o arquivo gerado será salvo|arquivo.pdf|

### Salvar documento
  
Extraia o texto do arquivo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|session||session|
|Salvar arquivo|Salve o arquivo com o nome e caminho especificados|arquivo.docx|
