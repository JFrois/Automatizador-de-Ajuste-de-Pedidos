# Automatizador de Processamento de Pedidos PDF

Este projeto é uma aplicação de desktop desenvolvida para automatizar o processo de download e processamento de documentos (como Pedidos de Compra) em formato PDF a partir de um portal B2B. A ferramenta extrai informações cruciais dos documentos, organiza os dados e gera um relatório consolidado em Excel, notificando o usuário por e-mail ao final do processo.

## Índice

  - [Funcionalidades](https://www.google.com/search?q=%23funcionalidades)
  - [Como Funciona](https://www.google.com/search?q=%23como-funciona)
  - [Pré-requisitos](https://www.google.com/search?q=%23pr%C3%A9-requisitos)
  - [Como Usar](https://www.google.com/search?q=%23como-usar)
  - [Estrutura do Projeto](https://www.google.com/search?q=%23estrutura-do-projeto)
  - [Bibliotecas Utilizadas](https://www.google.com/search?q=%23bibliotecas-utilizadas)

## Funcionalidades

  - **Interface Gráfica Amigável**: Interface intuitiva construída com `CustomTkinter` para facilitar a interação do usuário.
  - **Automação de Login**: Realiza o login seguro em portais web, incluindo suporte para autenticação de dois fatores (TOTP).
  - **Download e Renomeação Automática**: Navega até a seção de documentos, baixa os novos PDFs e os renomeia de forma inteligente com base em um código identificador (ex: código da peça), para fácil organização.
  - **Extração Inteligente de Dados**: Utiliza `PyPDF2` e expressões regulares (`re`) para ler o conteúdo dos PDFs e extrair dados como:
      - Número do Pedido/Documento
      - Código de Produto/Peça
      - Identificadores regionais (ex: cidade)
      - Datas de alteração e validade
      - Valores, preços e quantidades
  - **Categorização de Documentos**: Identifica e classifica o tipo de cada documento/alteração em categorias predefinidas, como:
      - Novo Pedido
      - Alteração de Preço
      - Custo Logístico
      - Prazo de Pagamento
      - Alteração de Validade
  - **Geração de Relatório Excel**: Consolida todos os dados extraídos em um único arquivo Excel, organizado em abas separadas para cada categoria de documento.
  - **Organização de Arquivos**: Move os arquivos PDFs processados para subpastas de `Processados` e `Nao Processados`, mantendo o diretório de trabalho limpo.
  - **Notificação por E-mail**: Envia um e-mail de status via Outlook ao final da execução, informando a quantidade de arquivos processados com sucesso e com falha, anexando o relatório Excel gerado.

## Como Funciona

O fluxo de trabalho da aplicação é dividido em duas etapas principais, controladas por botões na interface:

1.  **Baixar e Renomear PDFs (Botão 1)**:

      - O usuário insere suas credenciais (login, senha, código TOTP), e-mail e seleciona uma pasta de destino.
      - O robô (Selenium) abre o navegador, acessa o portal B2B e realiza o login.
      - Navega até a seção de documentos relevantes.
      - Identifica todos os "Novos Pedidos" ou documentos, faz o download do PDF de cada um e o renomeia com o código identificador correspondente.

2.  **Processar PDFs Baixados (Botão 2)**:

      - O script lê todos os arquivos `.pdf` na pasta selecionada pelo usuário.
      - Para cada arquivo, o texto é extraído e analisado para identificar o tipo de documento e os dados relevantes.
      - Os dados extraídos são armazenados em memória.
      - Após processar todos os arquivos, os dados são exportados para um arquivo Excel.
      - Os arquivos PDFs são movidos para as pastas `Processados` ou `Nao Processados` (dentro de subpastas baseadas em identificadores regionais, como a cidade).
      - Por fim, um e-mail de resumo é enviado para o usuário e para o endereço de e-mail em cópia.

A execução pode ser interrompida a qualquer momento pelo botão **Parar Execução**.

## Pré-requisitos

Antes de executar o projeto, certifique-se de que você tem o seguinte instalado:

  - Python 3.8 ou superior.
  - Navegador Google Chrome.
  - Microsoft Outlook instalado e configurado no seu computador.
  - As bibliotecas Python listadas na seção [Bibliotecas Utilizadas](https://www.google.com/search?q=%23bibliotecas-utilizadas).

Para instalar as dependências, execute o seguinte comando no seu terminal:

```bash
pip install pandas openpyxl pypdf2 pywin32 customtkinter selenium
```

## Como Usar

1.  Clone ou faça o download deste repositório.
2.  Instale todas as bibliotecas necessárias conforme a seção de [Pré-requisitos](https://www.google.com/search?q=%23pr%C3%A9-requisitos).
3.  Execute o arquivo principal:
    ```bash
    python Main.py
    ```
4.  Na janela da aplicação, preencha todos os campos:
      - **Login**: Seu nome de usuário para o portal.
      - **Senha**: Sua senha.
      - **Código TOTP**: O código de 6 dígitos do seu aplicativo autenticador.
      - **E-mail para Cópia**: O endereço de e-mail que receberá uma cópia do relatório.
      - **Pasta dos PDFs**: Clique em "Selecionar Pasta" para escolher onde os PDFs serão salvos e processados.
5.  Clique no botão **"1. Baixar e Renomear PDFs"** e aguarde a finalização do download.
6.  Após o download, clique no botão **"2. Processar PDFs Baixados"** para iniciar a extração de dados e a geração do relatório.
7.  Verifique a pasta selecionada para encontrar o arquivo Excel e as subpastas com os PDFs organizados.
8.  Verifique sua caixa de entrada do Outlook para o e-mail de confirmação.

## Estrutura do Projeto

  - `Main.py`: Contém o código da interface gráfica (classe `App`) e a lógica principal para o tratamento e processamento dos dados dos PDFs (classe `tratamentoDados`). É o ponto de entrada da aplicação.
  - `acessar_pedidos.py`: Contém a classe `AutomacaoPedidos`, responsável por toda a interação com o navegador web usando Selenium, incluindo login, navegação, download e renomeação dos arquivos.

## Bibliotecas Utilizadas

  - **customtkinter**: Para a criação da interface gráfica moderna.
  - **pandas**: Para manipulação de dados e exportação para o formato Excel.
  - **openpyxl**: Motor utilizado pelo pandas para escrever arquivos `.xlsx`.
  - **PyPDF2**: Para extrair texto de arquivos PDF.
  - **selenium**: Para automação do navegador web.
  - **pywin32** (`win3com.client`): Para interagir com a aplicação Outlook e enviar e-mails.
  - **os, re, shutil, threading, traceback, datetime**: Bibliotecas padrão do Python para manipulação de arquivos, expressões regulares, multithreading, tratamento de erros e datas.
