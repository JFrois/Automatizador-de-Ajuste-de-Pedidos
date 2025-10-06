import time
import traceback
import pandas as pd
from seleniumwire import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    NoSuchElementException,
    StaleElementReferenceException,
)
import pythoncom
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import re


class AutomacaoPedidos:
    def __init__(
        self,
        tratamento_dados,
        evento_parar,
        login,
        senha,
        autenticador,
        pedidos_especificos,
        modo,
    ):
        self.tratamento = tratamento_dados
        self.evento_parar = evento_parar
        self.login = login
        self.senha = senha
        self.autenticador = autenticador
        self.pedidos_especificos = pedidos_especificos
        self.modo = modo
        self.driver = None
        self.pedidos_sucesso = []
        self.pedidos_falha = []

        print("Iniciando o navegador com captura de rede...")
        chrome_options = Options()
        arguments = [
            "--lang=pt-BR",
            "--start-maximized",
            "--disable-notifications",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-extensions",
            "--log-level=3",
            "--ignore-certificate-errors",
            "--disable-popup-blocking",
        ]
        for argument in arguments:
            chrome_options.add_argument(argument)

        service = ChromeService()
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.short_wait = WebDriverWait(self.driver, 5)
        self.wait = WebDriverWait(self.driver, 20)

    def _verificar_parada(self):
        if self.evento_parar.is_set():
            raise InterruptedError("Automação interrompida pelo usuário.")

    def _fazer_login(self):
        print("\n---> Acessando a página de login.")
        # TODO: Substitua pela URL genérica do portal que você está automatizando
        self.driver.get("https://supplier-portal.example.com/login")
        self._verificar_parada()
        try:
            # TODO: Ajuste os seletores (ID, XPATH, etc.) para os campos da página de login real
            campo_usuario = self.wait.until(
                EC.presence_of_element_located((By.ID, "contentForm:profileIdInput"))
            )
            campo_usuario.send_keys(self.login)
            campo_senha = self.driver.find_element(By.ID, "contentForm:passwordInput")
            campo_senha.send_keys(self.senha)
            self.driver.find_element(By.ID, "contentForm:passwordLoginAction").click()

            campo_totp = self.wait.until(EC.presence_of_element_located((By.ID, "otp")))
            campo_totp.send_keys(self.autenticador)
            self.driver.find_element(By.XPATH, "//button[text()='Login']").click()
            print("Login realizado com sucesso.")
        except TimeoutException:
            raise TimeoutException(
                "Elemento não encontrado ou tempo excedido durante o login. Verifique as credenciais ou o TOTP."
            )
        except Exception as e:
            raise RuntimeError(f"Ocorreu um erro inesperado durante o login: {e}")

    def _processar_todos_arquivos(self, aba_pedidos_handle, url_busca):
        print("\n--- Fluxo 1: Iniciando busca por todos os pedidos ---")
        try:
            self.driver.switch_to.window(aba_pedidos_handle)
            self.driver.get(url_busca)
            del self.driver.requests
            # TODO: Ajuste o texto do filtro para ser o que aparece no portal real
            Select(self.driver.find_element(By.NAME, "filter")).select_by_visible_text(
                "todos os processos"
            )
            ActionChains(self.driver).send_keys(
                Keys.TAB, Keys.TAB, Keys.ENTER
            ).perform()
            self._iterar_entre_paginas(aba_pedidos_handle, url_busca)
        except TimeoutException:
            print("AVISO: A busca geral não retornou resultados.")
            self.pedidos_falha.append(
                {
                    "Arquivo": "FALHA NA BUSCA GERAL",
                    "Part Number": "N/A",
                    "Status": "Nenhum resultado encontrado",
                }
            )

    def _processar_fila_sequencial(self, aba_pedidos_handle):
        print("\n--- Fluxo 2: Iniciando Processamento Padrão (Fila Sequencial) ---")
        contador_pedidos = 1
        while True:
            self._verificar_parada()
            try:
                self.driver.refresh()
                self.wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//table[contains(@class, 'table-results')]//tr[td/a]") # Seletor genérico
                    )
                )

                print(f"\nBuscando pedido #{contador_pedidos} da fila...")
                primeira_linha = self.driver.find_element(
                    By.XPATH, "(//table[contains(@class, 'table-results')]//tr[td/a])[1]"
                )
                link_primeiro_pedido = primeira_linha.find_element(
                    By.XPATH, ".//td[3]/a"
                )
                codigo_pedido_completo = link_primeiro_pedido.text.strip().split(
                    " de "
                )[0]
                codigo_para_arquivo = primeira_linha.find_element(
                    By.XPATH, ".//td[4]"
                ).text.strip()

                print(
                    f"Processando primeiro item da lista: Pedido {codigo_pedido_completo}"
                )
                link_primeiro_pedido.click()

                self._baixar_pdf_aberto(
                    aba_pedidos_handle, codigo_pedido_completo, codigo_para_arquivo
                )
                contador_pedidos += 1
            except (NoSuchElementException, TimeoutException):
                print("\nFila de pedidos vazia. Processamento sequencial concluído.")
                break
            except Exception as e:
                print(f"ERRO inesperado ao processar a fila sequencial: {e}")
                traceback.print_exc()
                break

    def _buscar_e_processar_pedidos(
        self, pedido_para_buscar, aba_pedidos_handle, url_busca
    ):
        print(
            f"\n--- Fluxo 3: Iniciando busca por Part Number: '{pedido_para_buscar}' ---"
        )
        try:
            self.driver.switch_to.window(aba_pedidos_handle)
            self.driver.get(url_busca)
            del self.driver.requests
            elemento_busca = self.wait.until(
                EC.presence_of_element_located((By.NAME, "searchString"))
            )
            elemento_busca.clear()
            elemento_busca.send_keys(pedido_para_buscar)
            Select(self.driver.find_element(By.NAME, "filter")).select_by_visible_text(
                "todos os processos"
            )
            ActionChains(self.driver).send_keys(
                Keys.TAB, Keys.TAB, Keys.ENTER
            ).perform()
            print(
                f"Busca por '{pedido_para_buscar}' realizada. Aguardando resultados..."
            )
            self._iterar_entre_paginas(aba_pedidos_handle, url_busca)
        except TimeoutException:
            print(f"AVISO: A busca por '{pedido_para_buscar}' não retornou resultados.")
            self.pedidos_falha.append(
                {
                    "Arquivo": "FALHA NA BUSCA",
                    "Part Number": pedido_para_buscar,
                    "Status": "Nenhum resultado encontrado",
                }
            )
            return

    def _iterar_entre_paginas(self, aba_pedidos_handle, url_busca_retorno):
        try:
            self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//table[contains(@class, 'table-results')]//tr[td/a]")
                )
            )
        except TimeoutException:
            print("Nenhum pedido encontrado nos resultados da busca.")
            return

        pagina_atual = 1
        while True:
            self._verificar_parada()
            print(f"\nProcessando página de resultados {pagina_atual}...")

            try:
                linhas_de_pedido = self.driver.find_elements(
                    By.XPATH, "//table[contains(@class, 'table-results')]//tr[td/a]"
                )
                num_links = len(linhas_de_pedido)

                if num_links == 0:
                    print("Nenhum link de pedido encontrado nesta página.")
                    break

                print(f"Encontrados {num_links} pedido(s) na página atual.")

                for i in range(num_links):
                    self._verificar_parada()
                    try:
                        linha_atual = self.driver.find_elements(
                            By.XPATH, "//table[contains(@class, 'table-results')]//tr[td/a]"
                        )[i]

                        link = linha_atual.find_element(By.XPATH, ".//td[3]/a")
                        pedido_completo = link.text.strip().split(" de ")[0]
                        codigo_arquivo = linha_atual.find_element(
                            By.XPATH, ".//td[4]"
                        ).text.strip()

                        print(
                            f"Processando item {i+1}/{num_links}: Pedido {pedido_completo}"
                        )

                        link.click()

                        self.wait.until(EC.staleness_of(linha_atual))

                        self._baixar_pdf_aberto(
                            aba_pedidos_handle, pedido_completo, codigo_arquivo
                        )

                    except StaleElementReferenceException:
                        print(
                            "Erro de referência obsoleta, tentando recarregar e continuar do início da página."
                        )
                        self.driver.get(url_busca_retorno)
                        break
                    except Exception as e:
                        print(
                            f"ERRO ao processar um item. Pulando para o próximo. Erro: {e}"
                        )
                        self.driver.get(url_busca_retorno)
                        continue

                try:
                    botao_proxima = self.driver.find_element(
                        By.XPATH, "//a[text()='>']"
                    )
                    botao_proxima_ref = botao_proxima
                    self.driver.execute_script("arguments[0].click();", botao_proxima)
                    pagina_atual += 1
                    self.wait.until(EC.staleness_of(botao_proxima_ref))
                except (NoSuchElementException, TimeoutException):
                    print("Não há mais páginas de resultados. Fim da busca.")
                    break

            except (NoSuchElementException, TimeoutException):
                print(
                    "Não há mais páginas de resultados ou erro ao processar a página. Fim da busca."
                )
                break

    def _baixar_pdf_aberto(
        self, aba_pedidos_handle, codigo_pedido_completo, codigo_para_arquivo
    ):
        nome_provisorio_falha = f"{codigo_para_arquivo.replace(' ', '_')}_{codigo_pedido_completo}.pdf (FALHA)"
        try:
            self._verificar_parada()
            aba_detalhes_handle = self.driver.current_window_handle
            links_pdf = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.PARTIAL_LINK_TEXT, "Pedido em PDF") # Termo genérico
                )
            )
            del self.driver.requests
            handles_originais = set(self.driver.window_handles)
            self.driver.execute_script("arguments[0].click();", links_pdf[-1])
            self.wait.until(EC.new_window_is_opened(handles_originais))
            handles_novos = set(self.driver.window_handles)
            aba_download_handle = (handles_novos - handles_originais).pop()
            self.driver.switch_to.window(aba_download_handle)
            
            # Regex genérico para capturar a requisição do PDF
            request_pdf = self.driver.wait_for_request(r".*order-download.*\.pdf", timeout=20)
            
            if request_pdf.response and 200 <= request_pdf.response.status_code < 300:
                pdf_content = request_pdf.response.body
                prefixo = re.sub(r"[_/.]", " ", codigo_para_arquivo)
                sufixo = codigo_pedido_completo.replace("_", "")
                nome_final = f"{prefixo}_{sufixo}.pdf"
                caminho_completo = os.path.join(
                    self.tratamento.caminho_pasta_pdf, nome_final
                )
                with open(caminho_completo, "wb") as f:
                    f.write(pdf_content)
                print(f"Arquivo '{nome_final}' salvo com sucesso!")
                self.pedidos_sucesso.append(
                    {
                        "Arquivo": nome_final,
                        "Part Number": codigo_para_arquivo,
                        "Status": "Sucesso",
                    }
                )
            else:
                status_code = (
                    request_pdf.response.status_code if request_pdf.response else "N/A"
                )
                raise Exception(f"Requisição do PDF falhou. Status Code: {status_code}")
        except Exception as e:
            print(
                f"\nERRO: Não foi possível baixar o PDF para o pedido {codigo_pedido_completo}."
            )
            traceback.print_exc()
            self.pedidos_falha.append(
                {
                    "Arquivo": nome_provisorio_falha,
                    "Part Number": codigo_para_arquivo,
                    "Status": f"Falha no download: {e}",
                }
            )
        finally:
            current_handles = self.driver.window_handles
            for handle in current_handles:
                if handle != aba_pedidos_handle:
                    try:
                        self.driver.switch_to.window(handle)
                        self.driver.close()
                    except WebDriverException:
                        pass
            self.driver.switch_to.window(aba_pedidos_handle)

    def _navegar_e_baixar_pdfs(self):
        print("\n---> Navegando para a tela de pedidos.")
        aba_principal_handle = self.driver.current_window_handle
        # TODO: Ajuste o XPATH para o menu principal do portal real
        self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//div[contains(text(), 'Main Menu Item')]",
                )
            )
        ).click()

        self.wait.until(EC.number_of_windows_to_be(2))
        aba_pedidos_handle = [
            h for h in self.driver.window_handles if h != aba_principal_handle
        ][0]
        self.driver.switch_to.window(aba_pedidos_handle)
        url_da_pagina_de_busca = self.driver.current_url
        print(f"URL da página de pedidos definida: {url_da_pagina_de_busca}")

        if self.modo == "todos":
            self._processar_todos_arquivos(aba_pedidos_handle, url_da_pagina_de_busca)

        elif self.modo == "especificos":
            if self.pedidos_especificos:
                for pedido in self.pedidos_especificos:
                    self._verificar_parada()
                    self._buscar_e_processar_pedidos(
                        pedido.strip(), aba_pedidos_handle, url_da_pagina_de_busca
                    )
            else:
                self._processar_fila_sequencial(aba_pedidos_handle)

        print("\nProcesso de download finalizado.")

    def _exportar_relatorio_download(self):
        if not self.pedidos_sucesso and not self.pedidos_falha:
            print("Nenhum download foi realizado ou falhou. Relatório não gerado.")
            return None
        caminho_relatorio = os.path.join(
            self.tratamento.caminho_pasta_pdf, "Download_Report.xlsx"
        )
        try:
            with pd.ExcelWriter(caminho_relatorio, engine="openpyxl") as writer:
                if self.pedidos_sucesso:
                    pd.DataFrame(self.pedidos_sucesso).to_excel(
                        writer, sheet_name="Sucesso", index=False
                    )
                if self.pedidos_falha:
                    pd.DataFrame(self.pedidos_falha).to_excel(
                        writer, sheet_name="Falhas", index=False
                    )
            print(
                f"\n---> Relatório de downloads exportado com sucesso para: {caminho_relatorio}"
            )
            return caminho_relatorio
        except Exception as e:
            print(f"\n---> ERRO ao tentar gerar o relatório de downloads: {e}")
            traceback.print_exc()
            return None

    def executar(self):
        pythoncom.CoInitialize()
        try:
            self._verificar_parada()
            self._fazer_login()
            self._verificar_parada()
            self._navegar_e_baixar_pdfs()
            caminho_relatorio = self._exportar_relatorio_download()
            quantidade_final = len(self.pedidos_sucesso) + len(self.pedidos_falha)
            mensagem_final = "Download de arquivos concluído!"
            if quantidade_final == 0:
                mensagem_final = "Nenhum pedido encontrado para download com os critérios informados."
            tipo_email_final = (
                "download_especifico" if self.pedidos_especificos else "download_todos"
            )
            return {
                "sucesso": True,
                "mensagem": mensagem_final,
                "quantidade": quantidade_final,
                "pedidos_sucesso": self.pedidos_sucesso,
                "pedidos_falha": self.pedidos_falha,
                "tipo_email": tipo_email_final,
                "caminho_relatorio": caminho_relatorio,
            }
        except InterruptedError as e:
            return {"sucesso": False, "mensagem": str(e), "quantidade": 0}
        except (TimeoutException, RuntimeError, WebDriverException) as e:
            traceback.print_exc()
            return {"sucesso": False, "mensagem": str(e), "quantidade": 0}
        except Exception as e:
            traceback.print_exc()
            return {
                "sucesso": False,
                "mensagem": f"Erro fatal na automação: {e}",
                "quantidade": 0,
            }
        finally:
            if self.driver:
                self.driver.quit()
                print("Driver finalizado.")
            pythoncom.CoUninitialize()
