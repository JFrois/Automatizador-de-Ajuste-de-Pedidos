# ============================================================ #
# ------------------------------------------------------------ #
# ===> PASSO A PASSO DO NOSSO SCRIPT PARA AJUSTE DE PREÇO <=== #
# ------------------------------------------------------------ #
# ============================================================ #

# ---> 0. Importar as bibliotecas necessárias
from datetime import datetime
import traceback
import pandas as pd
import PyPDF2
import os
import re
from win32com.client import Dispatch
import customtkinter as ctk
from tkinter import messagebox
import threading
import pythoncom
import shutil
from tkinter import filedialog
from acessar_site_pedidos import AutomacaoPedidos


# ---> Função para identificar o último log
def encontra_ultimo_arquivo(folder_path, base_name):
    # ---> Encontra o arquivo de log com o maior número sequencial.
    try:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        arquivos = os.listdir(folder_path)
        arquivos_data = [
            file for file in arquivos if re.match(f"{base_name}(\\d+)\\.csv", file)
        ]
        if not arquivos_data:
            return None
        ultimo_arquivo = max(
            arquivos_data, key=lambda x: int(re.search(r"(\d+)", x).group())
        )
        return ultimo_arquivo
    except Exception as e:
        print(f"Erro ao encontrar último arquivo de log: {e}")
        return None


# ---> Função para criar o próximo arquivo de log
def cria_proximo_arquivo(folder_path, base_name):
    # ---> Gera o nome para o próximo arquivo de log sequencial.
    try:
        ultimo_arquivo = encontra_ultimo_arquivo(folder_path, base_name)
        if ultimo_arquivo:
            ultimo_numero_match = re.search(r"(\d+)", ultimo_arquivo)
            ultimo_numero = (
                int(ultimo_numero_match.group()) if ultimo_numero_match else 0
            )
            proximo_numero = ultimo_numero + 1
        else:
            proximo_numero = 1
        nome_proximo_arquivo = f"{base_name}{proximo_numero}.csv"
        proximo_caminho = os.path.join(folder_path, nome_proximo_arquivo)
        return proximo_caminho
    except Exception as e:
        print(f"Erro ao criar próximo arquivo de log: {e}")
        return os.path.join(folder_path, f"{base_name}1.csv")


# ===> Classe para o tratamento dos dados dos arquivos PDF
class tratamentoDados:
    def __init__(self):
        self.user = os.getlogin()
        # MODIFICADO: E-mail padrão agora usa um placeholder de domínio.
        self.user_mail = f"{self.user}@suaempresa.com"
        self.cc_mail = ""
        self.caminho_pasta_pdf = ""
        self.caminho_saida_excel = ""
        self.caminho_log = r""
        self.dados_extraidos = []
        self.pedidos_sucesso = set()
        self.pedidos_falha = set()

    # Processamento dos arquivos que já estão na pasta
    def processar_arquivos_baixados(self):
        self.dados_extraidos.clear()
        self.pedidos_sucesso.clear()
        self.pedidos_falha.clear()

        arquivos_pdf = [
            f for f in os.listdir(self.caminho_pasta_pdf) if f.lower().endswith(".pdf")
        ]
        if not arquivos_pdf:
            return False, "Nenhum arquivo PDF encontrado na pasta para processar."

        for nome_arquivo in arquivos_pdf:
            caminho_completo = os.path.join(self.caminho_pasta_pdf, nome_arquivo)
            conteudo_linhas = self.extrair_texto_pdf(caminho_completo)
            if conteudo_linhas:
                self.processar_conteudo(conteudo_linhas, nome_arquivo)
            else:
                self.pedidos_falha.add(nome_arquivo)
                self._mover_arquivo_processado(
                    nome_arquivo, "Falha_Extracao", sucesso=False
                )

        self.exportar_para_excel()
        self.enviar_email()

        if self.pedidos_falha:
            return (
                False,
                f"{len(self.pedidos_falha)} arquivo(s) falharam no processamento.",
            )
        return True, "Todos os arquivos foram processados com sucesso."

    # ---> 2. Extrai o texto de UM arquivo PDF
    def extrair_texto_pdf(self, caminho_arquivo):
        conteudo_linhas = []
        try:
            with open(caminho_arquivo, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        conteudo_linhas.extend(page_text.split("\n"))
            return conteudo_linhas
        except Exception as e:
            print(f"Erro ao ler o arquivo {caminho_arquivo}: {e}")
            return None

    # ---> 3. Organizar arquivos em pastas de processados e não processados
    def _mover_arquivo_processado(self, nome_arquivo, cidade, sucesso):
        try:
            pasta_status = "Processados" if sucesso else "Nao Processados"
            caminho_origem = os.path.join(self.caminho_pasta_pdf, nome_arquivo)
            pasta_destino = os.path.join(self.caminho_pasta_pdf, cidade, pasta_status)
            os.makedirs(pasta_destino, exist_ok=True)
            shutil.move(caminho_origem, pasta_destino)
            print(f"Arquivo '{nome_arquivo}' movido para a pasta '{pasta_status}'.")
        except Exception as e:
            print(f"ERRO ao tentar mover o arquivo '{nome_arquivo}': {e}")
            if nome_arquivo in self.pedidos_sucesso:
                self.pedidos_sucesso.remove(nome_arquivo)
            self.pedidos_falha.add(nome_arquivo)

    # ---> 4. Processa o conteúdo extraído de UM PDF
    def processar_conteudo(self, conteudo_linhas, nome_arquivo):
        try:
            texto_completo = "\n".join(conteudo_linhas)

            # ATENÇÃO: Os nomes das cidades são específicos para o seu cliente/empresa.
            # Altere "CidadeA" e "CidadeB" para os nomes corretos encontrados nos seus PDFs.
            linha_cidade = next(
                (l for l in conteudo_linhas if "CidadeA" in l or "CidadeB" in l), ""
            )
            cidade = "CidadeA" if "CidadeA" in linha_cidade else "CidadeB"
            print(f"PDF da loja de {cidade}\n")

            alteracao_pedido = any("Amendment" in linha for linha in conteudo_linhas)
            is_price_update = (
                "PRICE CHANGE" in texto_completo or "PRICE ADJUSTMENT" in texto_completo
            )

            # ---> Função: extrair_preco
            def extrair_preco(texto):
                if not texto:
                    return 0.0
                match = re.search(r"(\d{1,4}(?:[ .]\d{3})*[,.]\d{2})", texto)
                if not match:
                    match = re.search(r"(\d+[,.]\d{2})", texto)
                    if not match:
                        return 0.0
                numero_str = match.group(1)
                num_limpo = re.sub(r"[. ]", "", numero_str[:-3]) + numero_str[
                    -3:
                ].replace(",", ".")
                try:
                    return float(num_limpo)
                except ValueError:
                    return 0.0

            # ---> Função: extrair_quantidade
            def extrair_quantidade(texto):
                if not texto:
                    return 1
                match = re.search(r"per\s*(\d+)", texto, re.I)
                if match:
                    return int(match.group(1))
                partes = texto.split()
                for i, parte in enumerate(partes):
                    if re.search(r"\d+[,.]\d{2}$", parte):
                        if (i + 1) < len(partes) and partes[i + 1].isdigit():
                            return int(partes[i + 1])
                match = re.search(r"(\d+)\s*BRL", texto, re.I)
                if match:
                    return int(match.group(1))
                return 1

            dados = {"Arquivo": nome_arquivo, "Cidade": cidade}

            match_pedido = re.search(
                r"Purchase Order No\.\s*\n\s*[SE]\s*(\d+\s*\d+)",
                texto_completo,
                re.DOTALL,
            )
            if not match_pedido:
                match_pedido = re.search(r"[SE]\s*(\d{2}\s*\d{6})", texto_completo)
            dados["Codigo pedido"] = (
                re.sub(r"\s", "", match_pedido.group(1))
                if match_pedido
                else "Não Encontrado"
            )
            dados["Codigo peça"] = (
                nome_arquivo.split("_")[0] if nome_arquivo else "Não Encontrado"
            )

            dados_foram_extraidos = False

            if alteracao_pedido:
                ancora_datas = next(
                    (l for l in conteudo_linhas if "END-DATE" in l), None
                )
                if ancora_datas:
                    datas_encontradas = re.findall(
                        r"\d{2}\.\d{2}\.\d{2,4}", ancora_datas
                    )
                    dados["Data da Alteração"] = (
                        datas_encontradas[0].replace(".", "/")
                        if len(datas_encontradas) > 0
                        else "Não Encontrado"
                    )
                    dados["Data validade antiga"] = (
                        datas_encontradas[1].replace(".", "/")
                        if len(datas_encontradas) > 1
                        else "Não Encontrado"
                    )
                    dados["Data validade nova"] = (
                        datas_encontradas[2].replace(".", "/")
                        if len(datas_encontradas) > 2
                        else "Não Encontrado"
                    )

                linha_antigo, linha_novo = None, None

                idx_antigo = -1
                for i, linha in enumerate(conteudo_linhas):
                    if "OLD" in linha and ("BRL" in linha or "****" in linha):
                        linha_antigo = linha
                        idx_antigo = i
                        break
                if idx_antigo != -1:
                    for i in range(
                        idx_antigo + 1, min(idx_antigo + 5, len(conteudo_linhas))
                    ):
                        if "NEW" in conteudo_linhas[i] and (
                            "BRL" in conteudo_linhas[i] or "****" in conteudo_linhas[i]
                        ):
                            linha_novo = conteudo_linhas[i]
                            break

                if not linha_antigo:
                    linha_antigo = next(
                        (
                            l
                            for l in conteudo_linhas
                            if "OLD" in l and re.search(r"\d+[,.]\d{2}", l)
                        ),
                        None,
                    )
                if not linha_novo:
                    linha_novo = next(
                        (
                            l
                            for l in conteudo_linhas
                            if "NEW" in l and re.search(r"\d+[,.]\d{2}", l)
                        ),
                        None,
                    )

                # ---> Custo Logístico
                if "LOGISTIC COSTS" in texto_completo:
                    dados["Tipo de Alteração"] = "CUSTO LOGISTICO"
                    linha_total = next(
                        (l for l in conteudo_linhas if "TOTAL" in l), None
                    )
                    linha_transporte = next(
                        (l for l in conteudo_linhas if "Transport" in l), None
                    )
                    linha_embalagem = next(
                        (l for l in conteudo_linhas if "Packaging" in l), None
                    )
                    valor_transporte = extrair_preco(linha_transporte)
                    valor_embalagem = extrair_preco(linha_embalagem)
                    valor_logistico_total = valor_transporte + valor_embalagem
                    dados["Valor transporte"] = valor_transporte
                    dados["Valor embalagem"] = valor_embalagem
                    dados["Valor final"] = extrair_preco(linha_total)
                    if is_price_update and linha_antigo and linha_novo:
                        dados["Valor antigo"] = extrair_preco(linha_antigo)
                        dados["Valor novo"] = extrair_preco(linha_novo)
                    else:
                        valor_base_peca = (
                            dados.get("Valor final", 0.0) - valor_logistico_total
                        )
                        dados["Valor antigo"] = valor_base_peca
                        dados["Valor novo"] = valor_base_peca
                    dados["Valor código"] = (
                        dados.get("Valor novo", 0.0) + valor_logistico_total
                    )
                    dados["Documento X Codigo é Válido"] = (
                        "Correto"
                        if round(dados["Valor código"], 2)
                        == round(dados.get("Valor final", 0.0), 2)
                        else "Diferente"
                    )
                    dados["Preço por peça"] = extrair_quantidade(
                        linha_total or linha_antigo
                    )
                    dados_foram_extraidos = True

                # ---> Prazo de Pagamento
                elif "TERMS OF PAYMENT" in texto_completo:
                    dados["Tipo de Alteração"] = "PRAZO PAGAMENTO"
                    linha_prazo_antigo = next(
                        (
                            l
                            for l in conteudo_linhas
                            if "OLD" in l and "DAYS" in l.upper()
                        ),
                        None,
                    )
                    linha_prazo_novo = next(
                        (
                            l
                            for l in conteudo_linhas
                            if "NEW" in l and "DAYS" in l.upper()
                        ),
                        None,
                    )
                    prazo_antigo_match = (
                        re.search(r"(\d+)\s*DAYS", linha_prazo_antigo, re.IGNORECASE)
                        if linha_prazo_antigo
                        else None
                    )
                    prazo_novo_match = (
                        re.search(r"(\d+)\s*DAYS", linha_prazo_novo, re.IGNORECASE)
                        if linha_prazo_novo
                        else None
                    )
                    dados["Prazo antigo"] = (
                        prazo_antigo_match.group(1).strip()
                        if prazo_antigo_match
                        else "Não Encontrado"
                    )
                    dados["Prazo novo"] = (
                        prazo_novo_match.group(1).strip()
                        if prazo_novo_match
                        else "Não Encontrado"
                    )
                    dados["Preço Antigo"] = extrair_preco(linha_antigo)
                    dados["Preço Novo"] = extrair_preco(linha_novo)
                    dados["Preço por peça"] = extrair_quantidade(linha_antigo)

                    dados_foram_extraidos = True

                # ---> Alteração de Preço
                elif is_price_update:
                    dados["Tipo de Alteração"] = "ALTERAÇÃO DE PREÇO"
                    dados["Preço Antigo"] = extrair_preco(linha_antigo)
                    dados["Preço Novo"] = extrair_preco(linha_novo)
                    dados["Preço por peça"] = extrair_quantidade(linha_antigo)
                    dados_foram_extraidos = True

                # ---> Apenas alteração validade
                elif ancora_datas:
                    dados["Tipo de Alteração"] = "ALTERAÇÃO VALIDADE"
                    dados_foram_extraidos = True

            # ---> Pedido novo
            else:
                dados["Tipo de Alteração"] = "PEDIDO NOVO"
                linha_pecas = next(
                    (
                        l
                        for l in conteudo_linhas
                        if "OPEN" in l and ("BRL" in l or re.search(r"\d+[,.]\d{2}", l))
                    ),
                    None,
                )
                if linha_pecas:
                    valor_nova_peca = extrair_preco(linha_pecas)
                    quantidade_por_preco = extrair_quantidade(linha_pecas)
                    dados["Preço Peça"] = valor_nova_peca
                    dados["Preço por Peça"] = quantidade_por_preco
                    dados_foram_extraidos = True
                else:
                    dados_foram_extraidos = False

            if dados_foram_extraidos:
                self.dados_extraidos.append(dados)
                self.pedidos_sucesso.add(nome_arquivo)
                self._mover_arquivo_processado(nome_arquivo, cidade, sucesso=True)
            else:
                print(
                    f"Tipo de alteração não identificado para o arquivo {nome_arquivo}."
                )
                self.pedidos_falha.add(nome_arquivo)
                self._mover_arquivo_processado(nome_arquivo, cidade, sucesso=False)

        except Exception as e:
            print(f"Erro de processamento no arquivo {nome_arquivo}. Erro: {e}")
            self.pedidos_falha.add(nome_arquivo)
            # ATENÇÃO: Lógica de fallback com nomes de cidades
            linha_cidade_fallback = next(
                (l for l in conteudo_linhas if "CidadeA" in l or "CidadeB" in l), ""
            )
            cidade_extraida = (
                "CidadeA" if "CidadeA" in linha_cidade_fallback else "CidadeB"
            )
            self._mover_arquivo_processado(nome_arquivo, cidade_extraida, sucesso=False)
            traceback.print_exc()

    # ---> 5. Exporta a lista de dados para um único arquivo Excel com cabeçalhos específicos
    def exportar_para_excel(self):
        if not self.dados_extraidos:
            print("\nNenhum dado foi extraído para ser exportado.")
            return
        try:
            cabecalhos_especificos = {
                "PRAZO PAGAMENTO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Cidade",
                    "Tipo de Alteração",
                    "Data da Alteração",
                    "Data validade antiga",
                    "Data validade nova",
                    "Prazo antigo",
                    "Prazo novo",
                    "Preço Antigo",
                    "Preço Novo",
                    "Preço por peça",
                ],
                "ALTERAÇÃO DE PREÇO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Cidade",
                    "Tipo de Alteração",
                    "Data da Alteração",
                    "Data validade antiga",
                    "Data validade nova",
                    "Preço Antigo",
                    "Preço Novo",
                    "Preço por peça",
                ],
                "CUSTO LOGISTICO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Cidade",
                    "Tipo de Alteração",
                    "Data da Alteração",
                    "Data validade antiga",
                    "Data validade nova",
                    "Valor antigo",
                    "Valor novo",
                    "Valor embalagem",
                    "Valor transporte",
                    "Valor final",
                    "Valor código",
                    "Documento X Codigo é Válido",
                    "Preço por peça",
                ],
                "PEDIDO NOVO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Cidade",
                    "Tipo de Alteração",
                    "Preço Peça",
                    "Preço por Peça",
                ],
                "ALTERAÇÃO VALIDADE": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Cidade",
                    "Tipo de Alteração",
                    "Data da Alteração",
                    "Data validade antiga",
                    "Data validade nova",
                ],
            }
            df = pd.DataFrame(self.dados_extraidos)
            with pd.ExcelWriter(self.caminho_saida_excel, engine="openpyxl") as writer:
                ordem_abas = [
                    "PEDIDO NOVO",
                    "ALTERAÇÃO DE PREÇO",
                    "CUSTO LOGISTICO",
                    "PRAZO PAGAMENTO",
                    "ALTERAÇÃO VALIDADE",
                ]
                categorias_presentes = [
                    cat for cat in ordem_abas if cat in df["Tipo de Alteração"].unique()
                ]

                if not categorias_presentes:
                    print(
                        "\nALERTA: Nenhum dado com 'Tipo de Alteração' válido foi encontrado para exportar."
                    )
                    return
                for tipo_alteracao in categorias_presentes:
                    df_filtrado = df[df["Tipo de Alteração"] == tipo_alteracao].copy()
                    colunas_desejadas = cabecalhos_especificos.get(
                        tipo_alteracao, df_filtrado.columns.tolist()
                    )
                    for col in colunas_desejadas:
                        if col not in df_filtrado.columns:
                            df_filtrado[col] = None
                    df_final_aba = df_filtrado[colunas_desejadas]
                    df_final_aba.to_excel(
                        writer, sheet_name=str(tipo_alteracao), index=False
                    )
            print(
                f"\n---> DataFrame exportado com sucesso para: {self.caminho_saida_excel}"
            )
        except Exception as e:
            print(f"\n---> Ocorreu um erro ao exportar para o Excel: {e}")
            traceback.print_exc()

    # ---> 6. Envia e-mail com o resultado
    def enviar_email(self):
        if not self.pedidos_sucesso and not self.pedidos_falha:
            print("Nenhum processamento realizado, e-mail não enviado.")
            return
        try:
            outlook = Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = self.user_mail
            mail.cc = self.cc_mail
            subject_parts = []
            if self.pedidos_sucesso:
                subject_parts.append("SUCESSO")
            if self.pedidos_falha:
                subject_parts.append("FALHA")
            mail.Subject = f"Ajuste de Preços - {'/'.join(subject_parts)} - {datetime.now().strftime('%d/%m/%Y')}"
            qtde_sucesso = len(self.pedidos_sucesso)
            qtde_falha = len(self.pedidos_falha)
            total_pedidos = qtde_falha + qtde_sucesso
            texto_total = (
                f"<b>{total_pedidos} pedido processado.</b>"
                if total_pedidos == 1
                else f"<b>{total_pedidos} pedidos processados.</b>"
            )
            html_body = f"<p>Olá,</p><p>A rotina de ajuste de preços foi finalizada.</p><p>{texto_total}</p>"
            if self.pedidos_sucesso:
                texto_sucesso = (
                    f"{qtde_sucesso} arquivo processado com sucesso."
                    if qtde_sucesso == 1
                    else f"{qtde_sucesso} arquivos processados com sucesso."
                )
                html_body += f"<p><b>{texto_sucesso}</b></p>"
                if os.path.exists(self.caminho_saida_excel):
                    mail.Attachments.Add(os.path.abspath(self.caminho_saida_excel))
            if self.pedidos_falha:
                texto_falha = (
                    f"{qtde_falha} arquivo falhou."
                    if qtde_falha == 1
                    else f"{qtde_falha} arquivos falharam."
                )
                html_body += f"<p><b>{texto_falha}</b></p>"
            # MODIFICADO: Assinatura de e-mail genérica
            html_body += "<p>Att,<br>Bot de Automação</p>"
            mail.HTMLBody = html_body
            mail.Send()
            print("Email de status enviado com sucesso!")
        except Exception as e:
            print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")
            messagebox.showerror(
                "Erro de Email", f"Não foi possível enviar o e-mail: {e}"
            )


# ===> Classe da Interface Gráfica
class App(ctk.CTk):
    # ---> Inicialização da interface
    def __init__(self):
        super().__init__()
        self.tratamento = tratamentoDados()
        self.automacao_thread = None
        self.evento_parar = threading.Event()
        self.title("AJUSTADOR DE PEDIDOS")
        self.minsize(700, 600)
        ctk.set_appearance_mode("dark")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.criar_widgets()

    # ---> Função para criar os widgets da interface
    def criar_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        status_font = ctk.CTkFont(family="Arial", size=12)
        self.btn_font = ctk.CTkFont(family="Arial", size=14, weight="bold")

        # --- Frame para os campos de entrada ---
        inputs_frame = ctk.CTkFrame(main_frame)
        inputs_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        inputs_frame.grid_columnconfigure(1, weight=1)

        # ---> Campo de entrada para Login
        ctk.CTkLabel(inputs_frame, text="Login:", font=status_font).grid(
            row=0, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.entry_login = ctk.CTkEntry(inputs_frame, font=status_font)
        self.entry_login.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="ew")
        self.entry_login.insert(0, "")

        # ---> Campo de entrada para senha
        ctk.CTkLabel(inputs_frame, text="Senha:", font=status_font).grid(
            row=1, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.entry_senha = ctk.CTkEntry(inputs_frame, show="*", font=status_font)
        self.entry_senha.grid(row=1, column=1, padx=10, pady=(10, 5), sticky="ew")
        self.entry_senha.insert(0, "")

        # ---> Mostrar senha
        self.check_mostrar_pwd = ctk.CTkCheckBox(
            inputs_frame,
            text="Mostrar senha",
            font=status_font,
            command=self.mostrar_senha,
        )
        self.check_mostrar_pwd.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # ---> Campo de entrada para Autenticação de dois fatores
        ctk.CTkLabel(inputs_frame, text="Código TOTP", font=status_font).grid(
            row=2, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.entry_codigo = ctk.CTkEntry(inputs_frame, font=status_font)
        self.entry_codigo.grid(row=2, column=1, padx=10, pady=(10, 5), sticky="ew")
        self.entry_codigo.insert(0, "")

        # MODIFICADO: Rótulo do campo de e-mail genérico
        ctk.CTkLabel(inputs_frame, text="E-mail (em cópia):", font=status_font).grid(
            row=3, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.entry_email = ctk.CTkEntry(inputs_frame, font=status_font)
        self.entry_email.grid(row=3, column=1, padx=10, pady=(10, 5), sticky="ew")
        self.entry_email.insert(0, "")

        # ---> Campo de entrada para pasta de PDFs
        ctk.CTkLabel(inputs_frame, text="Pasta dos PDFs:", font=status_font).grid(
            row=4, column=0, padx=10, pady=(5, 10), sticky="w"
        )
        self.entry_pdf_path = ctk.CTkEntry(inputs_frame, font=status_font)
        self.entry_pdf_path.grid(row=4, column=1, padx=10, pady=(5, 10), sticky="ew")
        self.entry_pdf_path.insert(0, r"")

        # ---> Botão para selecionar pasta
        ctk.CTkButton(
            inputs_frame,
            text="Selecionar Pasta",
            font=status_font,
            command=self.selecionar_pasta,
        ).grid(row=4, column=2, padx=10, pady=(5, 10))

        # --- Frame para os botões ---
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        action_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # ---> Botão 1: Baixar PDFs
        self.btn_baixar = ctk.CTkButton(
            action_frame,
            text="1. Baixar e Renomear PDFs",
            font=self.btn_font,
            height=50,
            fg_color="#007BFF",  # Azul
            hover_color="#0056b3",
            command=self.iniciar_download,
        )
        self.btn_baixar.grid(row=0, column=0, padx=10, pady=20, sticky="ew")

        # ---> Botão 2: Processar PDFs
        self.btn_processar = ctk.CTkButton(
            action_frame,
            text="2. Processar PDFs Baixados",
            font=self.btn_font,
            height=50,
            fg_color="#4CAF50",
            hover_color="#45a049",
            command=self.iniciar_processamento,
        )
        self.btn_processar.grid(row=0, column=1, padx=10, pady=20, sticky="ew")

        self.btn_parar = ctk.CTkButton(
            action_frame,
            text="Parar Execução",
            font=self.btn_font,
            height=50,
            fg_color="#f44336",
            hover_color="#da190b",
            command=self.parar_automacao,
            state="disabled",
        )
        self.btn_parar.grid(row=0, column=2, padx=20, pady=20, sticky="w")
        self.return_user = ctk.CTkLabel(
            main_frame, text="Pronto para iniciar.", font=status_font
        )
        self.return_user.grid(row=2, column=0, pady=10, sticky="ew")

    # ---> Função para validar os campos de entrada
    def _validar_campos(self, verificar_login=True):
        erros = []
        if verificar_login:
            if not self.entry_login.get():
                erros.append("Login é obrigatório.")
            if not self.entry_senha.get():
                erros.append("Senha é obrigatória.")
            if not self.entry_codigo.get():
                erros.append("Código autenticador é obrigatório.")

        if not self.entry_pdf_path.get():
            erros.append("Caminho da pasta é obrigatório.")
        if not self.entry_email.get():
            erros.append("E-mail para cópia é obrigatório.")

        if erros:
            mensagem = "Erros de Validação:\n- " + "\n- ".join(erros)
            messagebox.showerror("Campos Inválidos", mensagem)
            return False

        # Atualiza os dados na classe de tratamento
        self.tratamento.cc_mail = self.entry_email.get()
        self.tratamento.caminho_pasta_pdf = self.entry_pdf_path.get()
        self.tratamento.caminho_saida_excel = os.path.join(
            self.tratamento.caminho_pasta_pdf, "Ajuste Pedidos.xlsx"
        )
        return True

    # ---> Função para iniciar o download dos PDFs
    def iniciar_download(self):
        if not self._validar_campos(verificar_login=True):
            return

        self._configurar_botoes_para_rodar(True)

        automator = AutomacaoPedidos(
            tratamento_dados=self.tratamento,
            evento_parar=self.evento_parar,
            login=self.entry_login.get(),
            senha=self.entry_senha.get(),
            autenticador=self.entry_codigo.get(),
        )
        self.automation_thread = threading.Thread(
            target=self.executar_e_atualizar_ui, args=(automator.executar,), daemon=True
        )
        self.automation_thread.start()

    # ---> Função para iniciar o processamento dos PDFs baixados
    def iniciar_processamento(self):
        if not self._validar_campos(verificar_login=False):
            return

        self._configurar_botoes_para_rodar(True)
        self.automation_thread = threading.Thread(
            target=self.executar_e_atualizar_ui,
            args=(self.tratamento.processar_arquivos_baixados,),
            daemon=True,
        )
        self.automation_thread.start()

    # ---> Função que executa a automação em uma thread separada
    def executar_e_atualizar_ui(self, funcao_alvo):
        self.atualizar_status("Iniciando processo...", "cyan")
        try:
            # A função alvo (seja download ou processamento) é executada aqui
            sucesso, mensagem = funcao_alvo()
            resultado = {"sucesso": sucesso, "mensagem": mensagem}
        except Exception as e:
            traceback.print_exc()
            resultado = {"sucesso": False, "mensagem": f"Erro crítico na thread: {e}"}

        self.after(0, self.finalizar_automacao, resultado)

    # ---> Função para finalizar a automação e atualizar a UI
    def finalizar_automacao(self, resultado):
        mensagem = resultado.get("mensagem", "Ocorreu um erro desconhecido.")
        cor = "green" if resultado.get("sucesso") else "red"

        if self.evento_parar.is_set():
            mensagem = "Automação interrompida pelo usuário."
            cor = "orange"

        self.atualizar_status(mensagem, cor)

        if resultado.get("sucesso"):
            messagebox.showinfo("Sucesso", mensagem)
        elif not self.evento_parar.is_set():
            messagebox.showerror("Erro", mensagem)

        self._configurar_botoes_para_rodar(False)

    # ---> Função para parar a automação
    def parar_automacao(self):
        if self.automation_thread and self.automation_thread.is_alive():
            self.evento_parar.set()
            self.atualizar_status("Sinal de parada enviado...", "orange")

    # ---> Função para configurar o estado dos botões
    def _configurar_botoes_para_rodar(self, rodando=True):
        if rodando:
            self.btn_baixar.configure(state="disabled")
            self.btn_processar.configure(state="disabled")
            self.btn_parar.configure(state="normal")
            self.evento_parar.clear()
        else:
            self.btn_baixar.configure(state="normal")
            self.btn_processar.configure(state="normal")
            self.btn_parar.configure(state="disabled")

    # ---> Função para atualizar o status na interface
    def atualizar_status(self, texto, cor):
        self.return_user.configure(text=texto, text_color=cor)

    # ---> Função para mostrar/ocultar senha
    def mostrar_senha(self):
        self.entry_senha.configure(
            show="" if self.entry_senha.cget("show") == "*" else "*"
        )

    # ---> Função para selecionar pasta de PDFs
    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory()
        if pasta_selecionada:
            self.entry_pdf_path.delete(0, ctk.END)
            self.entry_pdf_path.insert(0, pasta_selecionada)


# ---> Looping aplicativo
if __name__ == "__main__":
    app = App()
    app.mainloop()
