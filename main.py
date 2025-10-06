# ---> 0. Importar as bibliotecas necessárias
from datetime import datetime
from time import strftime
import traceback
import pandas as pd
import PyPDF2
import os
import re
from win32com.client import Dispatch
import customtkinter as ctk
from CTkToolTip import CTkToolTip
from tkinter import messagebox
import threading
import pythoncom
import shutil
from tkinter import filedialog
from web_automation import AutomacaoPedidos


# ---> Funções de Log (mantidas)
def encontra_ultimo_arquivo(folder_path, base_name):
    try:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        arquivos = os.listdir(folder_path)
        arquivos_data = [
            file for file in arquivos if re.match(f"{base_name}(\\d+)\\.csv", file)
        ]
        if not arquivos_data:
            return None
        return max(arquivos_data, key=lambda x: int(re.search(r"(\\d+)", x).group()))
    except Exception as e:
        print(f"Erro ao encontrar último arquivo de log: {e}")
        return None


def cria_proximo_arquivo(folder_path, base_name):
    try:
        ultimo_arquivo = encontra_ultimo_arquivo(folder_path, base_name)
        ultimo_numero = (
            int(re.search(r"(\\d+)", ultimo_arquivo).group()) if ultimo_arquivo else 0
        )
        proximo_numero = ultimo_numero + 1
        return os.path.join(folder_path, f"{base_name}{proximo_numero}.csv")
    except Exception as e:
        print(f"Erro ao criar próximo arquivo de log: {e}")
        return os.path.join(folder_path, f"{base_name}1.csv")


# ===> Classe para o tratamento dos dados dos arquivos PDF
class tratamentoDados:
    def __init__(self):
        self.user = os.getlogin()
        # TODO: Altere para um domínio de e-mail genérico
        self.user_mail = f"{self.user}@yourcompany.com"
        self.cc_mail = ""
        self.caminho_pasta_pdf = ""
        self.caminho_saida_excel = ""
        # TODO: Substitua o caminho de rede por um caminho local ou relativo para o portfólio
        self.caminho_log = r".\Logs" # Exemplo: "C:\Automation\Logs"
        self.dados_extraidos = []
        self.pedidos_sucesso = set()
        self.pedidos_falha = set()
        self.pedidos_existentes = set()

    def processar_arquivos_baixados(self, evento_parar):
        pythoncom.CoInitialize()
        try:
            self.dados_extraidos.clear()
            self.pedidos_sucesso.clear()
            self.pedidos_falha.clear()
            self.pedidos_existentes.clear()

            if not os.path.isdir(self.caminho_pasta_pdf):
                return {
                    "sucesso": False,
                    "mensagem": "O caminho especificado para os PDFs não é uma pasta válida.",
                    "quantidade": 0,
                    "tipo_email": "processamento",
                }
            arquivos_pdf = [
                f
                for f in os.listdir(self.caminho_pasta_pdf)
                if f.lower().endswith(".pdf")
            ]
            if not arquivos_pdf:
                return {
                    "sucesso": True,
                    "mensagem": "Nenhum arquivo PDF encontrado para processar.",
                    "quantidade": 0,
                    "tipo_email": "processamento",
                }
            for nome_arquivo in arquivos_pdf:
                if evento_parar.is_set():
                    print("Processamento interrompido pelo usuário.")
                    break
                caminho_completo = os.path.join(self.caminho_pasta_pdf, nome_arquivo)
                conteudo_linhas = self.extrair_texto_pdf(caminho_completo)
                if conteudo_linhas:
                    self.processar_conteudo(conteudo_linhas, nome_arquivo)
                else:
                    self.pedidos_falha.add(nome_arquivo)
                    self._mover_arquivo_processado(
                        nome_arquivo, "Falha_Extracao", sucesso=False
                    )

            total_pedidos = (
                len(self.pedidos_sucesso)
                + len(self.pedidos_falha)
                + len(self.pedidos_existentes)
            )
            if self.dados_extraidos:
                self.exportar_para_excel()

            sucesso_geral = not bool(self.pedidos_falha)
            mensagem = f"Processamento concluído. Sucesso: {len(self.pedidos_sucesso)}, Falha: {len(self.pedidos_falha)}, Existentes: {len(self.pedidos_existentes)}."

            return {
                "sucesso": sucesso_geral,
                "mensagem": mensagem,
                "quantidade": total_pedidos,
                "pedidos_sucesso": self.pedidos_sucesso,
                "pedidos_falha": self.pedidos_falha,
                "pedidos_existentes": self.pedidos_existentes,
                "tipo_email": "processamento",
            }
        finally:
            pythoncom.CoUninitialize()

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

    def _mover_arquivo_processado(self, nome_arquivo, location, sucesso):
        try:
            pasta_status = "Processados" if sucesso else "Nao Processados"
            caminho_origem = os.path.join(self.caminho_pasta_pdf, nome_arquivo)
            pasta_destino = os.path.join(self.caminho_pasta_pdf, location, pasta_status)
            os.makedirs(pasta_destino, exist_ok=True)
            caminho_final_arquivo = os.path.join(pasta_destino, nome_arquivo)
            if os.path.exists(caminho_final_arquivo):
                print(
                    f"AVISO: Arquivo '{nome_arquivo}' já existe no destino. O original será removido."
                )
                if nome_arquivo in self.pedidos_sucesso:
                    self.pedidos_sucesso.remove(nome_arquivo)
                self.pedidos_existentes.add(nome_arquivo)
                os.remove(caminho_origem)
                return
            shutil.move(caminho_origem, pasta_destino)
            print(f"Arquivo '{nome_arquivo}' movido para a pasta '{pasta_destino}'.")
        except Exception as e:
            print(f"ERRO ao tentar mover o arquivo '{nome_arquivo}': {e}")
            if nome_arquivo in self.pedidos_sucesso:
                self.pedidos_sucesso.remove(nome_arquivo)
            self.pedidos_falha.add(nome_arquivo)

    @staticmethod
    def extrair_preco(texto):
        if not texto:
            return 0.0
        match = re.search(r"(\d{1,4}(?:[ .]\d{3})*[,.]\d{2})", texto)
        if not match:
            match = re.search(r"(\d+[,.]\d{2})", texto)
            if not match:
                return 0.0
        numero_str = match.group(1)
        num_limpo = re.sub(r"[. ]", "", numero_str[:-3]) + numero_str[-3:].replace(
            ",", "."
        )
        try:
            return float(num_limpo)
        except ValueError:
            return 0.0

    @staticmethod
    def extrair_quantidade_principal(texto):
        if not texto:
            return 0
        match = re.search(r"(\d+)\s*PCS", texto, re.I)
        if match:
            return int(match.group(1))
        return 0

    @staticmethod
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

    def processar_conteudo(self, conteudo_linhas, nome_arquivo):
        try:
            texto_completo = "\n".join(conteudo_linhas)
            # Termos genéricos para localização
            linha_location = next(
                (l for l in conteudo_linhas if "LocationA" in l or "LocationB" in l), ""
            )
            location = "LocationA" if "LocationA" in linha_location else "LocationB"
            print(f"PDF da unidade de {location}\n")
            dados = {"Arquivo": nome_arquivo, "Location": location}
            codigo_pedido_completo = "Não Encontrado"
            
            # Regex genérico para número de pedido
            match_pedido_obj = re.search(
                r"Order No\.\s*\n\s*([A-Z]\s*\d+\s*\d+)",
                texto_completo,
                re.DOTALL,
            )
            if not match_pedido_obj:
                match_pedido_obj = re.search(r"([A-Z]\s*\d{2}\s*\d{6})", texto_completo)
            
            if match_pedido_obj:
                codigo_pedido_completo = re.sub(r"\s", "", match_pedido_obj.group(1))
            dados["Codigo pedido"] = codigo_pedido_completo
            dados["Codigo peça"] = (
                nome_arquivo.split("_")[0] if nome_arquivo else "Não Encontrado"
            )
            dados["Código pedido formatado"] = nome_arquivo.split("_")[0].replace(
                " ", ""
            )

            dados_foram_extraidos = False
            # Termos genéricos para tipo de alteração
            alteracao_pedido = any("Amendment" in linha for linha in conteudo_linhas)
            is_price_update = (
                "PRICE CHANGE" in texto_completo or "PRICE ADJUSTMENT" in texto_completo
            )
            linha_cancelamento = next(
                (l for l in conteudo_linhas if "CANCELLATION" in l.upper()), None
            )

            if linha_cancelamento is None:
                if codigo_pedido_completo != "Não Encontrado" and (
                    codigo_pedido_completo.startswith("P")
                    or codigo_pedido_completo.startswith("F")
                ):
                    dados["Tipo de Alteração"] = "FECHADO"
                    linha_pecas = next(
                        (
                            l
                            for l in conteudo_linhas
                            if "PCS" in l
                            and ("****" in l or re.search(r"\d+[,.]\d{2}", l))
                        ),
                        None,
                    )
                    if linha_pecas:
                        valor_nova_peca = self.extrair_preco(linha_pecas)
                        quantidade_por_preco = self.extrair_quantidade(linha_pecas)
                        quantidade = self.extrair_quantidade_principal(linha_pecas)
                        dados["Quantidade"] = quantidade
                        dados["Preço Peça"] = valor_nova_peca
                        dados["Preço por Peça"] = quantidade_por_preco
                    dados_foram_extraidos = True
                else:
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
                                idx_antigo + 1,
                                min(idx_antigo + 5, len(conteudo_linhas)),
                            ):
                                if "NEW" in conteudo_linhas[i] and (
                                    "BRL" in conteudo_linhas[i]
                                    or "****" in conteudo_linhas[i]
                                ):
                                    linha_novo = conteudo_linhas[i]
                                    break
                        if not linha_antigo:
                            linha_antigo = next(
                                (
                                    l
                                    for l in conteudo_linhas
                                    if "OLD" in l
                                    and re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}", l)
                                ),
                                None,
                            )
                        if not linha_novo:
                            linha_novo = next(
                                (
                                    l
                                    for l in conteudo_linhas
                                    if "NEW" in l
                                    and re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}", l)
                                ),
                                None,
                            )
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
                            valor_transporte = self.extrair_preco(linha_transporte)
                            valor_embalagem = self.extrair_preco(linha_embalagem)
                            valor_logistico_total = valor_transporte + valor_embalagem
                            dados["Valor transporte"] = valor_transporte
                            dados["Valor embalagem"] = valor_embalagem
                            dados["Valor final"] = self.extrair_preco(linha_total)
                            if is_price_update and linha_antigo and linha_novo:
                                dados["Valor antigo"] = self.extrair_preco(linha_antigo)
                                dados["Valor novo"] = self.extrair_preco(linha_novo)
                            else:
                                valor_base_peca = (
                                    dados.get("Valor final", 0.0)
                                    - valor_logistico_total
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
                            dados["Preço por peça"] = self.extrair_quantidade(
                                linha_total or linha_antigo
                            )
                            dados_foram_extraidos = True
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
                                re.search(
                                    r"(\d+)\s*DAYS", linha_prazo_antigo, re.IGNORECASE
                                )
                                if linha_prazo_antigo
                                else None
                            )
                            prazo_novo_match = (
                                re.search(
                                    r"(\d+)\s*DAYS", linha_prazo_novo, re.IGNORECASE
                                )
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
                            dados["Preço Antigo"] = self.extrair_preco(linha_antigo)
                            dados["Preço Novo"] = self.extrair_preco(linha_novo)
                            dados["Preço por peça"] = self.extrair_quantidade(
                                linha_antigo
                            )
                            dados_foram_extraidos = True
                        elif is_price_update:
                            dados["Tipo de Alteração"] = "ALTERAÇÃO DE PREÇO"
                            dados["Preço Antigo"] = self.extrair_preco(linha_antigo)
                            dados["Preço Novo"] = self.extrair_preco(linha_novo)
                            dados["Preço por peça"] = self.extrair_quantidade(
                                linha_antigo
                            )
                            dados_foram_extraidos = True
                        elif ancora_datas:
                            dados["Tipo de Alteração"] = "ALTERAÇÃO VALIDADE"
                            dados_foram_extraidos = True
                    else:
                        dados["Tipo de Alteração"] = "PEDIDO NOVO"
                        linha_pecas = next(
                            (
                                l
                                for l in conteudo_linhas
                                if "OPEN" in l
                                and ("BRL" in l or re.search(r"\d+[,.]\d{2}", l))
                            ),
                            None,
                        )
                        if linha_pecas:
                            valor_nova_peca = self.extrair_preco(linha_pecas)
                            quantidade_por_preco = self.extrair_quantidade(linha_pecas)
                            dados["Preço Peça"] = valor_nova_peca
                            dados["Preço por Peça"] = quantidade_por_preco
                            dados_foram_extraidos = True
                        else:
                            dados_foram_extraidos = False
            else:
                dados["Tipo de Alteração"] = "CANCELAMENTO"
                dados_foram_extraidos = True

            if dados_foram_extraidos:
                self.dados_extraidos.append(dados)
                self.pedidos_sucesso.add(nome_arquivo)
                self._mover_arquivo_processado(nome_arquivo, location, sucesso=True)
            else:
                print(
                    f"Tipo de alteração não identificado para o arquivo {nome_arquivo}."
                )
                self.pedidos_falha.add(nome_arquivo)
                self._mover_arquivo_processado(nome_arquivo, location, sucesso=False)
        except Exception as e:
            print(f"Erro de processamento no arquivo {nome_arquivo}. Erro: {e}")
            self.pedidos_falha.add(nome_arquivo)
            linha_cidade_fallback = next(
                (l for l in conteudo_linhas if "LocationA" in l or "LocationB" in l), ""
            )
            location_extraida = (
                "LocationA" if "LocationA" in linha_cidade_fallback else "LocationB"
            )
            self._mover_arquivo_processado(nome_arquivo, location_extraida, sucesso=False)
            traceback.print_exc()

    def exportar_para_excel(self):
        if not self.dados_extraidos:
            print("\nNenhum dado foi extraído para ser exportado.")
            return
        try:
            cabecalhos_especificos = {
                "RESUMO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Código pedido formatado",
                    "Valor novo",
                    "Valor embalagem",
                    "Valor transporte",
                    "Valor final",
                    "Valor código",
                    "Documento X Codigo é Válido",
                    "Preço Peça",
                    "Preço por Peça",
                ],
                "PRAZO PAGAMENTO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Location",
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
                    "Location",
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
                    "Location",
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
                    "Location",
                    "Tipo de Alteração",
                    "Preço Peça",
                    "Preço por Peça",
                ],
                "ALTERAÇÃO VALIDADE": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Location",
                    "Tipo de Alteração",
                    "Data da Alteração",
                    "Data validade antiga",
                    "Data validade nova",
                ],
                "CANCELAMENTO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Location",
                    "Tipo de Alteração",
                ],
                "FECHADO": [
                    "Arquivo",
                    "Codigo pedido",
                    "Codigo peça",
                    "Location",
                    "Tipo de Alteração",
                    "Quantidade",
                    "Preço Peça",
                    "Preço por Peça",
                ],
            }
            df = pd.DataFrame(self.dados_extraidos)

            with pd.ExcelWriter(self.caminho_saida_excel, engine="openpyxl") as writer:
                df_resumo = df.copy()
                colunas_resumo = cabecalhos_especificos["RESUMO"]

                for col in colunas_resumo:
                    if col not in df_resumo.columns:
                        df_resumo[col] = None

                df_resumo_final = df_resumo[colunas_resumo]
                df_resumo_final.to_excel(writer, sheet_name="RESUMO", index=False)

                ordem_abas_especificas = [
                    "PEDIDO NOVO",
                    "ALTERAÇÃO DE PREÇO",
                    "CUSTO LOGISTICO",
                    "PRAZO PAGAMENTO",
                    "ALTERAÇÃO VALIDADE",
                    "CANCELAMENTO",
                    "FECHADO",
                ]

                categorias_presentes = [
                    cat
                    for cat in ordem_abas_especificas
                    if cat in df["Tipo de Alteração"].unique()
                ]

                if not categorias_presentes and df.empty:
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


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.tratamento = tratamentoDados()
        self.automacao_thread = None
        self.evento_parar = threading.Event()
        self.title("RPA PDF Processor - Order Automation")
        self.minsize(850, 600)
        ctk.set_appearance_mode("dark")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.criar_widgets()

    def criar_widgets(self):
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)
        status_font = ctk.CTkFont(family="Arial", size=12)
        self.btn_font = ctk.CTkFont(family="Arial", size=14, weight="bold")
        inputs_frame = ctk.CTkFrame(main_frame)
        inputs_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        inputs_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(inputs_frame, text="Login:", font=status_font).grid(
            row=0, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.entry_login = ctk.CTkEntry(
            inputs_frame,
            font=status_font,
            placeholder_text="Informe o login do portal",
        )
        self.entry_login.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="ew")
        ctk.CTkLabel(inputs_frame, text="Senha:", font=status_font).grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_senha = ctk.CTkEntry(
            inputs_frame,
            show="*",
            font=status_font,
            placeholder_text="Informe a senha do portal",
        )
        self.entry_senha.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.check_mostrar_pwd = ctk.CTkCheckBox(
            inputs_frame,
            text="Mostrar senha",
            font=status_font,
            command=self.mostrar_senha,
        )
        self.check_mostrar_pwd.grid(row=1, column=2, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(inputs_frame, text="Código TOTP:", font=status_font).grid(
            row=2, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_codigo = ctk.CTkEntry(
            inputs_frame, font=status_font, placeholder_text="Informe o código 2FA"
        )
        self.entry_codigo.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(inputs_frame, text="Part Number(s):", font=status_font).grid(
            row=3, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_pedido = ctk.CTkEntry(
            inputs_frame,
            font=status_font,
            placeholder_text="Informe os Part Numbers separados por vírgula ou ponto e vírgula",
        )
        self.entry_pedido.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.btn_carregar_planilha = ctk.CTkButton(
            inputs_frame,
            text="Carregar Planilha",
            font=status_font,
            command=self.carregar_planilha_pedidos,
        )
        self.btn_carregar_planilha.grid(row=3, column=2, padx=10, pady=5)
        ctk.CTkLabel(inputs_frame, text="E-mail (cópia):", font=status_font).grid(
            row=4, column=0, padx=10, pady=5, sticky="w"
        )
        self.entry_email = ctk.CTkEntry(
            inputs_frame,
            font=status_font,
            placeholder_text="Informe e-mail adicional para receber os resultados",
        )
        self.entry_email.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkLabel(inputs_frame, text="Pasta dos PDFs:", font=status_font).grid(
            row=5, column=0, padx=10, pady=(5, 10), sticky="w"
        )
        self.entry_pdf_path = ctk.CTkEntry(
            inputs_frame,
            font=status_font,
            placeholder_text="Selecione a pasta para salvar e processar os PDFs",
        )
        self.entry_pdf_path.grid(row=5, column=1, padx=10, pady=(5, 10), sticky="ew")
        self.btn_pasta = ctk.CTkButton(
            inputs_frame,
            text="Selecionar Pasta",
            font=status_font,
            command=self.selecionar_pasta,
        )
        self.btn_pasta.grid(row=5, column=2, padx=10, pady=(5, 10))
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        action_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_baixar_todos = ctk.CTkButton(
            action_frame,
            text="1. Baixar Todos",
            font=self.btn_font,
            height=50,
            fg_color="#FFAE00",
            hover_color="#9B6A00",
            command=lambda: self.iniciar_download(modo="todos"),
        )
        self.btn_baixar_todos.grid(row=0, column=0, padx=5, pady=10, sticky="ew")

        self.btn_baixar_especifico = ctk.CTkButton(
            action_frame,
            text="2. Baixar Específicos",
            font=self.btn_font,
            height=50,
            fg_color="#007BFF",
            hover_color="#0056b3",
            command=lambda: self.iniciar_download(modo="especificos"),
        )
        self.btn_baixar_especifico.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        self.btn_processar = ctk.CTkButton(
            action_frame,
            text="3. Processar PDFs",
            font=self.btn_font,
            height=50,
            fg_color="#4CAF50",
            hover_color="#45a049",
            command=self.iniciar_processamento,
        )
        self.btn_processar.grid(row=0, column=2, padx=5, pady=10, sticky="ew")
        self.btn_parar = ctk.CTkButton(
            action_frame,
            text="Parar",
            font=self.btn_font,
            height=50,
            fg_color="#f44336",
            hover_color="#da190b",
            command=self.parar_automacao,
            state="disabled",
        )
        self.btn_parar.grid(row=0, column=3, padx=10, pady=10, sticky="ew")
        status_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_frame.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        status_frame.grid_columnconfigure(0, weight=1)
        self.return_user = ctk.CTkLabel(
            status_frame, text="Pronto para iniciar.", font=status_font
        )
        self.return_user.grid(row=0, column=0, sticky="ew")

        CTkToolTip(
            self.btn_carregar_planilha,
            message="Selecione um arquivo Excel (.xlsx).\n A planilha deve ter os Part Numbers na coluna A, começando da segunda linha.",
        )
        CTkToolTip(
            self.btn_pasta,
            message="Selecione a pasta que desejar para download ou processamento dos PDFs.",
        )
        CTkToolTip(
            self.btn_baixar_todos,
            message="Essa opção baixa todos os pedidos do portal do cliente.",
        )
        CTkToolTip(
            self.btn_baixar_especifico,
            message="Essa opção baixa apenas os pedidos específicos informados no campo 'Part Number(s)'.\nSe este campo estiver vazio, o aplicativo realizará o processo de download de todos os pedidos novos da plataforma.",
        )
        CTkToolTip(
            self.btn_processar,
            message="Essa opção processa os arquivos PDF que já foram baixados e estão na pasta selecionada.",
        )
        CTkToolTip(
            self.btn_parar,
            message="Essa opção interrompe qualquer processo em andamento.",
        )

    def carregar_planilha_pedidos(self):
        tipos_arquivo = [("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de Part Numbers", filetypes=tipos_arquivo
        )
        if not caminho_arquivo:
            return
        try:
            df = pd.read_excel(
                caminho_arquivo, usecols=[0], header=None, skiprows=1, engine="openpyxl"
            )
            pedidos = df.iloc[:, 0].dropna().astype(str).tolist()
            if not pedidos:
                messagebox.showwarning(
                    "Planilha Vazia",
                    "Nenhum Part Number encontrado na coluna A da planilha.",
                )
                return
            self.entry_pedido.delete(0, ctk.END)
            self.entry_pedido.insert(0, "; ".join(pedidos))
            self.atualizar_status(
                f"{len(pedidos)} Part Numbers carregados da planilha.", "green"
            )
        except Exception as e:
            messagebox.showerror(
                "Erro ao Ler Planilha", f"Não foi possível ler o arquivo:\n{e}"
            )
            traceback.print_exc()

    def _obter_lista_de_pedidos(self):
        pedido_texto = self.entry_pedido.get().strip()
        if not pedido_texto:
            return []
        return [p.strip() for p in re.split(r"[,;]", pedido_texto) if p.strip()]

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
            messagebox.showerror(
                "Campos Inválidos", "Erros de Validação:\n- " + "\n- ".join(erros)
            )
            return False
        self.tratamento.cc_mail = self.entry_email.get()
        self.tratamento.caminho_pasta_pdf = self.entry_pdf_path.get()
        self.tratamento.caminho_saida_excel = os.path.join(
            self.tratamento.caminho_pasta_pdf, "Order_Adjustments.xlsx"
        )
        return True

    def iniciar_download(self, modo):
        if not self._validar_campos(verificar_login=True):
            return
        lista_pedidos, log_rotina = [], ""
        if modo == "todos":
            log_rotina = "RPA Orders - Download de todos os pedidos"
        elif modo == "especificos":
            lista_pedidos = self._obter_lista_de_pedidos()
            log_rotina = (
                "RPA Orders - Download de pedidos novos"
                if not lista_pedidos
                else "RPA Orders - Download de pedidos específicos"
            )
        self._configurar_botoes_para_rodar(True)
        automator = AutomacaoPedidos(
            tratamento_dados=self.tratamento,
            evento_parar=self.evento_parar,
            login=self.entry_login.get(),
            senha=self.entry_senha.get(),
            autenticador=self.entry_codigo.get(),
            pedidos_especificos=lista_pedidos,
            modo=modo,
        )
        self.automation_thread = threading.Thread(
            target=self.executar_e_atualizar_ui,
            args=(automator.executar, log_rotina),
            daemon=True,
        )
        self.automation_thread.start()

    def iniciar_processamento(self):
        if not self._validar_campos(verificar_login=False):
            return
        self._configurar_botoes_para_rodar(True)
        self.automation_thread = threading.Thread(
            target=self.executar_e_atualizar_ui,
            args=(
                self.tratamento.processar_arquivos_baixados,
                "RPA Orders - Processamento de PDFs",
                self.evento_parar,
            ),
            daemon=True,
        )
        self.automation_thread.start()

    def _criar_log(self, rotina, quantidade_itens, base_name):
        try:
            total_tempo_humano_por_item = 0
            tempo_download_humano_por_item, tempo_preenchimento_humano_por_item = 90, 90
            if "Processamento" in rotina:
                total_tempo_humano_por_item = tempo_preenchimento_humano_por_item
            elif "Download" in rotina:
                total_tempo_humano_por_item = tempo_download_humano_por_item
            tempo_bot_fixo_segundos = 45
            df_log = pd.DataFrame(
                {
                    "Usuario": [self.tratamento.user],
                    "Rotina": [rotina],
                    "Data/Hora": [strftime("%Y-%m-%d %H:%M:%S")],
                    "Quantidade de itens": [quantidade_itens],
                    "Tempo_humano(segundos)": [
                        quantidade_itens * total_tempo_humano_por_item
                    ],
                    "Tempo_bot(segundos)": [tempo_bot_fixo_segundos],
                }
            )
            caminho_do_log = cria_proximo_arquivo(
                folder_path=self.tratamento.caminho_log, base_name=base_name
            )
            if not caminho_do_log:
                print("Erro: Não foi possível gerar um nome de arquivo de log.")
                return
            df_log.to_csv(caminho_do_log, index=False, sep=";", encoding="utf-8-sig")
            print(f"Log salvo com sucesso em: {caminho_do_log}")
        except Exception as e:
            print(f"Erro ao salvar log: {e}")
            traceback.print_exc()

    def executar_e_atualizar_ui(self, funcao_alvo, rotina, *args_para_funcao):
        self.atualizar_status(f"Executando: {rotina}...", "#3498DB")
        start_time = datetime.now()
        resultado_completo = {}
        try:
            resultado_completo = funcao_alvo(*args_para_funcao)
        except Exception as e:
            traceback.print_exc()
            resultado_completo = {
                "sucesso": False,
                "mensagem": f"Erro crítico na thread: {e}",
                "quantidade": 0,
            }
        finally:
            end_time = datetime.now()
            tempo_total = (end_time - start_time).total_seconds()
            print(f"INFO: Tempo real de execução foi de {tempo_total:.2f} segundos.")
            contagem_final = resultado_completo.get("quantidade", 0)
            if contagem_final > 0:
                self._criar_log(rotina, contagem_final, "data")
            else:
                print(f"Nenhum item foi processado (Itens: {contagem_final}).")
            self.after(0, self.finalizar_automacao, resultado_completo)

    def finalizar_automacao(self, resultado):
        mensagem, cor = resultado.get("mensagem", "Ocorreu um erro desconhecido."), (
            "green" if resultado.get("sucesso") else "red"
        )
        if self.evento_parar.is_set():
            mensagem, cor = "Automação interrompida pelo usuário.", "orange"
        self.atualizar_status(mensagem, cor)
        if not self.evento_parar.is_set():
            self._enviar_email_notificacao(resultado)
            if resultado.get("sucesso"):
                messagebox.showinfo("Sucesso", mensagem)
            else:
                if (
                    "O caminho especificado para os PDFs não é uma pasta válida"
                    not in mensagem
                ):
                    messagebox.showerror("Erro", mensagem)
        self._configurar_botoes_para_rodar(False)

    def _enviar_email_notificacao(self, resultado):
        tipo_email = resultado.get("tipo_email")
        if not tipo_email:
            return

        pedidos_sucesso = resultado.get("pedidos_sucesso", [])
        pedidos_falha = resultado.get("pedidos_falha", [])

        if (
            not pedidos_sucesso
            and not pedidos_falha
            and not resultado.get("pedidos_existentes")
        ):
            print("Nenhum item para relatar, e-mail não enviado.")
            return

        try:
            outlook = Dispatch("outlook.application")
            mail = outlook.CreateItem(0)

            mail.To = self.tratamento.user_mail
            mail.CC = self.tratamento.cc_mail

            subject_parts = []
            if pedidos_sucesso:
                subject_parts.append("SUCESSO")
            if tipo_email == "processamento" and resultado.get("pedidos_existentes"):
                subject_parts.append("AVISO")
            if pedidos_falha:
                subject_parts.append("FALHA")

            status_subject = "/".join(subject_parts) if subject_parts else "INFO"
            data_hoje = datetime.now().strftime("%d/%m/%Y")
            html_body = ["<p>Olá,</p>"]

            if tipo_email == "processamento":
                mail.Subject = f"Processamento de PDFs - {status_subject} - {data_hoje}"
                qtde_sucesso = len(pedidos_sucesso)
                qtde_falha = len(pedidos_falha)
                qtde_existentes = len(resultado.get("pedidos_existentes", set()))
                total = qtde_sucesso + qtde_falha + qtde_existentes
                html_body.append(
                    f"<p>A rotina de <b>processamento de PDFs</b> foi finalizada. {total} arquivo(s) verificado(s).</p>"
                )
                if qtde_sucesso > 0:
                    html_body.append(
                        f'<p><b style="color:green;">{qtde_sucesso} arquivo(s) processado(s) com sucesso.</b></p>'
                    )
                if qtde_existentes > 0:
                    html_body.append(
                        f'<p><b style="color:orange;">{qtde_existentes} arquivo(s) não foram movidos pois já existem no destino.</b></p>'
                    )
                if qtde_falha > 0:
                    html_body.append(
                        f'<p><b style="color:red;">{qtde_falha} arquivo(s) falharam no processamento.</b></p>'
                    )
                if qtde_sucesso > 0 and os.path.exists(
                    self.tratamento.caminho_saida_excel
                ):
                    mail.Attachments.Add(
                        os.path.abspath(self.tratamento.caminho_saida_excel)
                    )

            elif tipo_email in ["download_especifico", "download_todos"]:
                titulo_rotina = (
                    "Download de Pedidos Específicos"
                    if tipo_email == "download_especifico"
                    else "Download de Todos os Pedidos"
                )
                mail.Subject = f"{titulo_rotina} - {status_subject} - {data_hoje}"
                qtde_sucesso = len(pedidos_sucesso)
                qtde_falha = len(pedidos_falha)
                total = qtde_sucesso + qtde_falha
                html_body.append(
                    f"<p>A rotina de <b>{titulo_rotina}</b> foi finalizada. {total} download(s) tentado(s).</p>"
                )
                if qtde_sucesso > 0:
                    html_body.append(
                        f'<p><b style="color:green;">{qtde_sucesso} arquivo(s) baixado(s) com sucesso.</b></p>'
                    )
                if qtde_falha > 0:
                    html_body.append(
                        f'<p><b style="color:red;">{qtde_falha} item/itens não puderam ser baixados.</b> Consulte o relatório em anexo.</p>'
                    )
                caminho_relatorio = resultado.get("caminho_relatorio")
                if caminho_relatorio and os.path.exists(caminho_relatorio):
                    mail.Attachments.Add(caminho_relatorio)

            html_body.append("<br><p>Atenciosamente,<br>Automation Bot</p>")

            mail.HTMLBody = "".join(html_body)
            mail.Send()

            print("Email de status enviado com sucesso!")

        except Exception as e:
            print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")
            messagebox.showerror(
                "Erro de Email", f"Não foi possível enviar o e-mail de notificação: {e}"
            )

    def parar_automacao(self):
        if self.automation_thread and self.automation_thread.is_alive():
            self.evento_parar.set()
            self.atualizar_status(
                "Sinal de parada enviado... Aguardando finalização.", "orange"
            )

    def _configurar_botoes_para_rodar(self, rodando=True):
        state = "disabled" if rodando else "normal"
        self.btn_baixar_todos.configure(state=state)
        self.btn_baixar_especifico.configure(state=state)
        self.btn_processar.configure(state=state)
        self.btn_parar.configure(state="normal" if rodando else "disabled")
        if rodando:
            self.evento_parar.clear()

    def atualizar_status(self, texto, cor):
        self.return_user.configure(text=texto, text_color=cor)

    def mostrar_senha(self):
        self.entry_senha.configure(
            show="" if self.entry_senha.cget("show") == "*" else "*"
        )

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory()
        if pasta_selecionada:
            self.entry_pdf_path.delete(0, ctk.END)
            self.entry_pdf_path.insert(0, pasta_selecionada)


if __name__ == "__main__":
    app = App()
    app.mainloop()
