# Criar-Word-em-Massa
Código com interface gráfica simples e interativa, focado em criar arquivos DOCX e PDF em massa com dados extraídos do Excel

**Criador de Arquivos em Massa**

Este é um programa em Python que usa as bibliotecas openpyxl e docxtpl para gerar automaticamente words com base em uma planilha de dados. O programa permite que o usuário selecione a planilha de dados e a pasta de destino, e forneça os dados do documento. Ele então usa os dados da planilha para preencher um modelo de documento do Word e salvar o documento preenchido na pasta de destino.

**Como usar o programa**

Para usar o programa, execute o script "wordt.py" no Python. Isso abrirá uma janela do PySimpleGUI com campos para selecionar a planilha de dados, a pasta de destino e os dados do documento. Depois de preencher esses campos, clique no botão "Iniciar" para gerar os words.

**Arquivos do programa**

O programa é composto pelos seguintes arquivos:

- wordt.py: O script principal que contém o código Python para o programa.
- modelo.docx: O modelo de documento do Word que é preenchido com os dados da planilha de dados.
- README.md: Este arquivo de documentos.

**Bibliotecas usadas**

O programa usa as seguintes bibliotecas Python:

- openpyxl: Biblioteca para ler e gravar arquivos Excel.
- docxtpl: Biblioteca para preencher modelos de documento do Word.
- PySimpleGUI: Biblioteca para criar interfaces gráficas do usuário.

**Como o programa funciona**

O programa segue os seguintes passos:

1. Cria uma janela do PySimpleGUI com campos para selecionar a planilha de dados, a pasta de destino e os dados do documento.
1. Quando o usuário clica no botão "Iniciar", o programa carrega uma planilha de dados selecionada usando a biblioteca openpyxl.
1. O programa percorre cada linha da planilha de dados e preenche o modelo de documento do Word usando os dados da linha atual.
1. O documento preenchido é salvo na pasta de destino especificada pelo usuário.
1. O programa exibe uma mensagem de conclusão quando todos os words foram gerados com.

**Possíveis melhorias**

Algumas melhorias possíveis para o programa incluem:

- Adicionando uma opção de conversor de documentos gerados em PDF.
- Implementar cache de memorização para melhorar o desempenho do programa.
- Permitir que o usuário selecione o modelo de documento do Word em vez de usar um modelo fixo.
- Adicionando mais campos ao modelo de documento do Word para incluir mais informações dos dados da planilha.
- Adicionando opções de configuração para o programa, como idioma e formato de dados.

**Últimas melhorias**

Algumas melhorias realziadas:

- Botão de criar um documento word e converter em PDF.
