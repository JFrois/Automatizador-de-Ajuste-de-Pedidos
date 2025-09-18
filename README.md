# Automatizador de Processamento de Pedidos PDF

Este projeto é uma aplicação de desktop desenvolvida para automatizar o processo de download e processamento de documentos (como Pedidos de vendas) em formato PDF a partir de um portal B2B. A ferramenta extrai informações cruciais dos documentos, organiza os dados e gera um relatório consolidado em Excel, notificando o usuário por e-mail ao final do processo.

> **Aviso Importante:** Este repositório serve como um **template genérico**. Para que a automação funcione, você **precisará adaptar** partes do código, como URLs e seletores de elementos web, para o portal específico que você deseja automatizar. Consulte a seção de [Configuração Essencial](https://www.google.com/search?q=%23configura%C3%A7%C3%A3o-essencial).

## Índice

  - [Funcionalidades](https://www.google.com/search?q=%23funcionalidades)
  - [Como Funciona](https://www.google.com/search?q=%23como-funciona)
  - [Pré-requisitos](https://www.google.com/search?q=%23pr%C3%A9-requisitos)
  - [Configuração Essencial](https://www.google.com/search?q=%23configura%C3%A7%C3%A3o-essencial)
  - [Como Usar](https://www.google.com/search?q=%23como-usar)
  - [Estrutura do Projeto](https://www.google.com/search?q=%23estrutura-do-projeto)
  - [Bibliotecas Utilizadas](https://www.google.com/search?q=%23bibliotecas-utilizadas)
  - [Contribuições](https://www.google.com/search?q=%23contribui%C3%A7%C3%B5es)
  - [Licença](https://www.google.com/search?q=%23licen%C3%A7a)

## Funcionalidades

  - **Interface Gráfica Amigável**: Interface intuitiva construída com `CustomTkinter` para facilitar a interação do usuário.
  - **Automação de Login**: Realiza o login seguro em portais web, incluindo suporte para autenticação de dois fatores (TOTP).
  - **Download e Renomeação Automática**: Navega até a seção de documentos, baixa os PDFs e os renomeia de forma inteligente com base em um código identificador (ex: código da peça).
  - **Extração Inteligente de Dados**: Utiliza `PyPDF2` e expressões regulares (`re`) para ler o conteúdo dos PDFs e extrair dados como:
      - Número do Pedido/Documento
      - Código de Produto/Peça
      - Identificadores regionais (ex: cidade)
      - Datas de alteração e validade
      - Valores, preços e quantidades
  - **Categorização de Documentos**: Identifica e classifica o tipo de cada documento em categorias predefinidas, como:
      - Novo Pedido
      - Alteração de Preço
      - Custo Logístico
      - Prazo de Pagamento
      - Alteração de Validade
  - **Geração de Relatório Excel**: Consolida todos os dados extraídos em um único arquivo Excel, organizado em abas separadas para cada categoria.
  - **Organização de Arquivos**: Move os arquivos PDFs processados para subpastas de `Processados` e `Nao Processados`, mantendo o diretório de trabalho limpo.
  - **Notificação por E-mail**: Envia um e-mail de status via Outlook ao final da execução, informando o sucesso da operação e anexando o relatório Excel gerado.

## Como Funciona

O fluxo de trabalho da aplicação é dividido em duas etapas principais, controladas pela interface:

1.  **Baixar e Renomear PDFs (Botão 1)**:

      - O usuário insere suas credenciais (login, senha, código TOTP), e-mail e seleciona uma pasta de destino.
      - O robô (`Selenium`) abre o navegador, acessa o portal B2B e realiza o login.
      - Navega até a seção de documentos, identifica os pedidos, faz o download do PDF de cada um e o renomeia com um código identificador correspondente.

2.  **Processar PDFs Baixados (Botão 2)**:

      - O script lê todos os arquivos `.pdf` na pasta selecionada.
      - Para cada arquivo, o texto é extraído e analisado para identificar o tipo de documento e os dados relevantes.
      - Após processar todos os arquivos, os dados são exportados para um arquivo Excel.
      - Os PDFs são movidos para as pastas `Processados` ou `Nao Processados`.
      - Por fim, um e-mail de resumo é enviado para o usuário.

A execução pode ser interrompida a qualquer momento pelo botão **Parar Execução**.

## Pré-requisitos

Antes de executar o projeto, certifique-se de que você tem o seguinte instalado:

  - Python 3.8 ou superior.
  - Navegador Google Chrome.
  - Microsoft Outlook instalado e configurado no seu computador.
  - As bibliotecas Python listadas abaixo.

Para instalar as dependências, execute o seguinte comando no seu terminal:

```bash
pip install pandas openpyxl pypdf2 pywin32 customtkinter selenium selenium-wire
```

**Nota:** O `Selenium` moderno geralmente gerencia o `chromedriver` automaticamente. Se você encontrar problemas, pode ser necessário baixar o `chromedriver` correspondente à sua versão do Chrome e garantir que ele esteja no `PATH` do seu sistema.

## Configuração Essencial

Antes de usar a aplicação, você **deve** ajustar o código-fonte para o portal web que deseja automatizar. Procure pelos comentários `# SENSÍVEL:`, `# TODO:` e `# NOTE:` nos arquivos.

Os principais pontos a serem configurados estão no arquivo `acessar_site_pedidos.py`:

1.  **URL de Login**:

      - Altere a variável `login_url` com a URL correta da página de login do portal.

    <!-- end list -->

    ```python
    # TODO: Substitua pela URL de login correta do portal que você está automatizando.
    login_url = "https://portal-do-seu-cliente.com/login"
    ```

2.  **Seletores de Elementos Web (IDs, XPaths)**:

      - Inspecione a página do portal e encontre os seletores corretos para os campos de login, senha, botões e links de navegação. Atualize os métodos `_fazer_login` e `_navegar_e_baixar_pdfs`.

    <!-- end list -->

    ```python
    # Exemplo de seletor que precisa ser adaptado:
    campo_usuario = self.wait.until(
        EC.presence_of_element_located((By.ID, "contentForm:profileIdInput"))
    )
    ```

3.  **Padrões de Documentos e Dados**:

      - No arquivo `main.py`, ajuste as palavras-chave (ex: `Cidade A`, `Cidade B`) e expressões regulares para corresponderem ao conteúdo dos seus PDFs.

## Como Usar

1.  Clone ou faça o download deste repositório.
2.  Instale todas as bibliotecas necessárias (veja [Pré-requisitos](https://www.google.com/search?q=%23pr%C3%A9-requisitos)).
3.  **Realize a [Configuração Essencial](https://www.google.com/search?q=%23configura%C3%A7%C3%A3o-essencial)** nos arquivos `.py`.
4.  Execute o arquivo principal:
    ```bash
    python main.py
    ```
5.  Na janela da aplicação, preencha todos os campos:
      - **Login**: Seu nome de usuário para o portal.
      - **Senha**: Sua senha.
      - **Código TOTP**: O código de 6 dígitos do seu aplicativo autenticador.
      - **E-mail (Cópia)**: Endereço de e-mail que receberá uma cópia do relatório.
      - **Pasta dos PDFs**: Clique em "Selecionar Pasta" para escolher onde os PDFs serão salvos e processados.
6.  Clique no botão **"1. Baixar e Renomear PDFs"** e aguarde a finalização.
7.  Após o download, clique em **"2. Processar PDFs Baixados"** para gerar o relatório.
8.  Verifique a pasta selecionada para encontrar o arquivo Excel e os PDFs organizados.
9.  Verifique sua caixa de entrada do Outlook para o e-mail de confirmação.

## Estrutura do Projeto

  - `main.py`: Contém a interface gráfica (classe `App`) e a lógica para o processamento dos dados dos PDFs (classe `tratamentoDados`). É o ponto de entrada da aplicação.
  - `acessar_site_pedidos.py`: Contém a classe `AutomacaoPedidos`, responsável por toda a interação com o navegador (login, navegação e download dos arquivos).

## Bibliotecas Utilizadas

  - **customtkinter**: Para a criação da interface gráfica moderna.
  - **pandas**: Para manipulação de dados e exportação para o formato Excel.
  - **openpyxl**: Motor utilizado pelo pandas para escrever arquivos `.xlsx`.
  - **PyPDF2**: Para extrair texto de arquivos PDF.
  - **selenium**: Para automação do navegador web.
  - **selenium-wire**: Uma extensão do Selenium que permite inspecionar as requisições de rede, usada aqui para capturar e salvar os arquivos PDF diretamente.
  - **pywin32** (`win3com.client`): Para interagir com a aplicação Microsoft Outlook e enviar e-mails.
  - **os, re, shutil, threading, traceback, datetime**: Bibliotecas padrão do Python usadas para diversas funcionalidades.

## Contribuições

Contribuições são bem-vindas\! Se você tiver sugestões para melhorar este projeto, sinta-se à vontade para abrir uma *issue* ou enviar um *pull request*.
