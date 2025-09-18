import datetime
import time
from tkinter import messagebox
import traceback
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
)
import pythoncom
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import re
from win32com.client import (
    Dispatch,
)


class AutomacaoPedidos:
    def __init__(
        self,
        tratamento_dados,
        evento_parar,
        login,
        senha,
        autenticador,
        pedido_especifico,
    ):
        self.user = os.getlogin()
        self.tratamento = tratamento_dados
        self.evento_parar = evento_parar
        self.login = login
        self.senha = senha
        # SENSÍVEL: Domínio de e-mail genérico.
        self.user_mail = f"{self.user}@sua_empresa.com"
        self.cc_mail = ""
        self.autenticador = autenticador
        self.pedidos_especificos = pedido_especifico
        self.driver = None
        self.pedidos_sucesso = set()
        self.pedidos_falha = set()

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
        # SENSÍVEL: URL de login substituída por uma genérica.
        # TODO: Substitua pela URL de login correta do portal que você está automatizando.
        login_url = "https://portal-do-seu-cliente.com/login"
        self.driver.get(login_url)
        self._verificar_parada()
        try:
            # NOTE: Os IDs dos campos ('contentForm:profileIdInput', 'contentForm:passwordInput', etc.)
            # são específicos do site que você está automatizando. Você precisará inspecionar
            # a página de login do seu portal e atualizar esses seletores.
            campo_usuario = self.wait.until(
                EC.presence_of_element_located((By.ID, "contentForm:profileIdInput"))
            )
            campo_usuario.clear()
            campo_usuario.send_keys(self.login)
            campo_senha = self.wait.until(
                EC.presence_of_element_located((By.ID, "contentForm:passwordInput"))
            )
            campo_senha.clear()
            campo_senha.send_keys(self.senha)
            try:
                checkbox = self.short_wait.until(
                    EC.presence_of_element_located((By.ID, "contentForm:j_idt25_input"))
                )
                if not checkbox.is_selected():
                    checkbox.click()
            except TimeoutException:
                pass  # Checkbox não encontrado, ignora
            botao_login = self.wait.until(
                EC.element_to_be_clickable((By.ID, "contentForm:passwordLoginAction"))
            )
            botao_login.click()
            campo_totp = self.wait.until(EC.presence_of_element_located((By.ID, "otp")))
            campo_totp.clear()
            campo_totp.send_keys(self.autenticador)
            botao_confirmar_totp = self.wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Login']"))
            )
            botao_confirmar_totp.click()
            print("Login realizado com sucesso.")
        except TimeoutException as e:
            raise TimeoutException(
                f"Elemento não encontrado ou tempo de espera excedido durante o login: {e}"
            )
        except Exception as e:
            raise RuntimeError(f"Ocorreu um erro inesperado durante o login: {e}")

    def _ler_todos_pedidos_do_site(self):
        print("Lendo a lista de todos os pedidos disponíveis no site...")
        try:
            # NOTE: Este seletor XPath é específico do site original.
            # Você precisará atualizá-lo para corresponder à estrutura da tabela de pedidos do seu portal.
            self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//table[contains(@class, 'TD01b')]")
                )
            )
            elementos_pedidos = self.driver.find_elements(
                By.XPATH, "//table[contains(@class, 'TD01b')]//tr/td[4]"
            )
            if elementos_pedidos:
                pedidos_string = elementos_pedidos[0].text.strip()
                lista_de_pedidos = [
                    p.strip() for p in re.split(r"[,;]", pedidos_string) if p.strip()
                ]
            else:
                lista_de_pedidos = []
            print(f"Encontrados {len(lista_de_pedidos)} pedidos na página.")
            return lista_de_pedidos
        except Exception as e:
            print(f"Erro ao ler la lista de pedidos do site: {e}")
            return []

    def _buscar_e_processar_pedidos(
        self, pedido_para_buscar, aba_pedidos_handle, url_busca
    ):
        print(
            f"\n--- Iniciando processamento para o termo de busca: '{pedido_para_buscar}' ---"
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
            print(f"Busca por '{pedido_para_buscar}' realizada.")
            self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//table[contains(@class, 'TD01b')]")
                )
            )
        except TimeoutException:
            print(
                f"AVISO: A busca por '{pedido_para_buscar}' não retornou resultados."
            )
            self.pedidos_falha.add(pedido_para_buscar)
            return

        pagina_atual = 1
        while True:
            self._verificar_parada()
            print(f"\nProcessando página {pagina_atual} de resultados...")
            
            # NOTE: O seletor XPath abaixo é específico. Adapte-o para o seu site.
            xpath_links_resultados = "//table[contains(@class, 'TD01b')]//tr/td[3]/a"
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, xpath_links_resultados))
                )
                links_encontrados = self.driver.find_elements(
                    By.XPATH, xpath_links_resultados
                )
                num_links_na_pagina = len(links_encontrados)
                if num_links_na_pagina == 0:
                    print("Nenhum resultado de pedido encontrado nesta página.")
                    break
                print(
                    f"Encontrados {num_links_na_pagina} pedido(s) na página {pagina_atual}."
                )
            except TimeoutException:
                print(
                    f"Nenhum resultado de pedido encontrado na página {pagina_atual}."
                )
                break

            for i in range(num_links_na_pagina):
                try:
                    linha_resultado = self.driver.find_elements(
                        By.XPATH, "//table[contains(@class, 'TD01b')]//tr[td[3]/a]"
                    )[i]
                    link_para_clicar = linha_resultado.find_element(
                        By.XPATH, ".//td[3]/a"
                    )

                    codigo_pedido_completo = link_para_clicar.text.strip().split(
                        " de "
                    )[0]
                    codigo_para_arquivo = linha_resultado.find_element(
                        By.XPATH, ".//td[4]"
                    ).text.strip()

                    print(
                        f"Processando item {i+1}/{num_links_na_pagina}: Pedido {codigo_pedido_completo}"
                    )
                    link_para_clicar.click()
                    self._baixar_pdf_aberto(
                        aba_pedidos_handle, codigo_pedido_completo, codigo_para_arquivo
                    )

                    print("Retornando à página de resultados...")
                    self.driver.switch_to.window(aba_pedidos_handle)
                    self.driver.get(url_busca)

                    elemento_busca = self.wait.until(
                        EC.presence_of_element_located((By.NAME, "searchString"))
                    )
                    elemento_busca.clear()
                    elemento_busca.send_keys(pedido_para_buscar)
                    Select(
                        self.driver.find_element(By.NAME, "filter")
                    ).select_by_visible_text("todos os processos")
                    ActionChains(self.driver).send_keys(
                        Keys.TAB, Keys.TAB, Keys.ENTER
                    ).perform()
                    self.wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//table[contains(@class, 'TD01b')]")
                        )
                    )

                    if pagina_atual > 1:
                        print(f"Navegando de volta para a página {pagina_atual}...")
                        for _ in range(pagina_atual - 1):
                            botao_proxima_pagina = self.wait.until(
                                EC.element_to_be_clickable(
                                    (
                                        By.XPATH,
                                        "//a[contains(@href, 'pager.offset')]//img[contains(@alt, 'próxima página')]",
                                    )
                                )
                            )
                            botao_proxima_pagina.click()
                            self.wait.until(EC.staleness_of(botao_proxima_pagina))
                except Exception as e:
                    print(
                        f"ERRO ao processar um item na lista. Pulando. Erro: {e}"
                    )
                    self.driver.get(url_busca)
                    continue
            try:
                # NOTE: Seletor para o botão "próxima página". Adapte se necessário.
                botao_proxima_pagina = self.driver.find_element(
                    By.XPATH,
                    "//a[contains(@href, 'pager.offset')]//img[contains(@alt, 'próxima página')]",
                )
                print("Botão 'próxima página' encontrado. Clicando...")
                botao_proxima_pagina.click()
                pagina_atual += 1
                self.wait.until(EC.staleness_of(botao_proxima_pagina))
            except NoSuchElementException:
                print("Não há mais páginas de resultados.")
                break

    def _baixar_pdf_aberto(
        self, aba_pedidos_handle, codigo_pedido_completo, codigo_para_arquivo
    ):
        try:
            self._verificar_parada()
            links_pdf = self.wait.until(
                EC.presence_of_all_elements_located(
                    (By.PARTIAL_LINK_TEXT, "Pedido em PDF")
                )
            )
            if not links_pdf:
                print("AVISO: Nenhum link de PDF encontrado na página de detalhes.")
                self.pedidos_falha.add(codigo_pedido_completo)
                return

            aba_detalhes_handle = self.driver.current_window_handle
            links_pdf[-1].click()

            self.wait.until(EC.number_of_windows_to_be(3))
            aba_download_handle = [
                h
                for h in self.driver.window_handles
                if h not in [aba_pedidos_handle, aba_detalhes_handle]
            ][0]
            self.driver.switch_to.window(aba_download_handle)

            # SENSÍVEL: Padrão de URL de download generalizado.
            # O original era '.*bestellung.*'. Adapte para um padrão que capture a URL do PDF no seu portal.
            request_pdf = self.driver.wait_for_request(r".*order.*\.pdf", timeout=20)
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
                self.pedidos_sucesso.add(prefixo)

            self.driver.close()
            self.driver.switch_to.window(aba_detalhes_handle)
        except (TimeoutException, IndexError) as e:
            print(
                f"ERRO: Não foi possível baixar o PDF para {codigo_pedido_completo}. Erro: {e}"
            )
            self.pedidos_falha.add(codigo_para_arquivo)
        finally:
            handles = self.driver.window_handles
            if len(handles) > 2 and self.driver.current_window_handle not in [
                aba_pedidos_handle
            ]:
                self.driver.close()
            self.driver.switch_to.window(aba_pedidos_handle)

    def _navegar_e_baixar_pdfs(self):
        print("\n---> Navegando para a tela de pedidos.")
        aba_principal_handle = self.driver.current_window_handle
        
        # NOTE: Este seletor é altamente específico. Você precisará encontrar o
        # seletor correto para o link ou botão que leva à área de pedidos do seu portal.
        self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//div[normalize-space()='Online Orders Series Material (ONB)']",
                )
            )
        ).click()

        self.wait.until(EC.number_of_windows_to_be(2))
        aba_pedidos_handle = [
            h for h in self.driver.window_handles if h != aba_principal_handle
        ][0]
        self.driver.switch_to.window(aba_pedidos_handle)

        url_da_pagina_de_busca = self.driver.current_url
        print(f"URL da página de busca definida: {url_da_pagina_de_busca}")

        lista_para_processar = (
            self.pedidos_especificos or self._ler_todos_pedidos_do_site()
        )

        if not lista_para_processar:
            print("Nenhuma lista de pedidos para processar.")
            return

        for pedido in lista_para_processar:
            self._buscar_e_processar_pedidos(
                pedido.strip(), aba_pedidos_handle, url_da_pagina_de_busca
            )
        print("\nProcesso finalizado para todos os pedidos da lista.")

    def enviar_email(self):
        if not self.pedidos_sucesso and not self.pedidos_falha:
            print("Nenhum processamento realizado, e-mail não será enviado.")
            return

        print("\n---> Preparando para enviar e-mail de status...")
        try:
            outlook = Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = self.user_mail
            mail.CC = self.cc_mail

            subject_parts = []
            if self.pedidos_sucesso:
                subject_parts.append("SUCESSO")
            if self.pedidos_falha:
                subject_parts.append("FALHA")
            status_subject = "/".join(subject_parts)
            data_hoje = datetime.datetime.now().strftime("%d/%m/%Y")
            mail.Subject = f"Download de Pedidos - {status_subject} - {data_hoje}"

            qtde_sucesso = len(self.pedidos_sucesso)
            qtde_falha = len(self.pedidos_falha)
            total_pedidos = qtde_sucesso + qtde_falha

            texto_total = (
                f"<b>{total_pedidos} pedido processado.</b>"
                if total_pedidos == 1
                else f"<b>{total_pedidos} pedidos processados.</b>"
            )

            html_body = f"""
            <p>Olá,</p>
            <p>A rotina de download de pedidos foi finalizada.</p>
            <p>{texto_total}</p>
            """

            if self.pedidos_sucesso:
                texto_sucesso = (
                    f"{qtde_sucesso} arquivo processado com sucesso."
                    if qtde_sucesso == 1
                    else f"{qtde_sucesso} arquivos processados com sucesso."
                )
                html_body += f'<p><b style="color:green;">{texto_sucesso}</b></p>'

            if self.pedidos_falha:
                texto_falha_header = (
                    f"{qtde_falha} arquivo falhou ao ser processado:"
                    if qtde_falha == 1
                    else f"{qtde_falha} arquivos falharam ao ser processados:"
                )
                lista_pedidos_falha_html = "<br>".join(sorted(list(self.pedidos_falha)))
                html_body += f"""
                <p><b style="color:red;">{texto_falha_header}</b></p>
                <p>{lista_pedidos_falha_html}</p>
                """

            # SENSÍVEL: Assinatura do bot generalizada.
            html_body += "<p>Att,<br>Bot de Automação</p>"
            mail.HTMLBody = html_body
            mail.Send()
            print("Email de status enviado com sucesso!")

        except Exception as e:
            print(f"Ocorreu um erro ao tentar enviar o e-mail: {e}")
            messagebox.showerror(
                "Erro de Email", f"Não foi possível enviar o e-mail de status: {e}"
            )

    def executar(self):
        self.pedidos_sucesso.clear()
        self.pedidos_falha.clear()
        pythoncom.CoInitialize()
        try:
            self._verificar_parada()
            self._fazer_login()
            self._verificar_parada()
            self._navegar_e_baixar_pdfs()
            quantidade_final = len(self.pedidos_sucesso)
            return True, "Download de arquivos concluído com sucesso!", quantidade_final
        except InterruptedError as e:
            return False, str(e), 0
        except (TimeoutException, RuntimeError, WebDriverException) as e:
            messagebox.showerror("Erro de Automação", str(e))
            return False, str(e), 0
        except Exception as e:
            traceback.print_exc()
            return False, f"Erro fatal na automação: {e}", 0
        finally:
            self.enviar_email()
            if self.driver:
                self.driver.quit()
                print("Driver finalizado.")
            pythoncom.CoUninitialize()
