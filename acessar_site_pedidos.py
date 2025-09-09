import time
from tkinter import messagebox
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    StaleElementReferenceException,
)
import pythoncom
import os
from selenium.webdriver.common.action_chains import ActionChains


class AutomacaoPedidos:
    def __init__(self, tratamento_dados, evento_parar, login, senha, autenticador):
        self.tratamento = tratamento_dados
        self.evento_parar = evento_parar
        self.login = login
        self.senha = senha
        self.autenticador = autenticador
        self.driver = None
        self.actions = None
        self.wait = None
        self.short_wait = None

    def _verificar_parada(self):
        if self.evento_parar.is_set():
            raise InterruptedError("Automação interrompida pelo usuário.")

    def _iniciar_driver(self):
        print("Iniciando o navegador...")
        chrome_options = Options()
        pasta_download = self.tratamento.caminho_pasta_pdf
        prefs = {
            "download.default_directory": pasta_download,
            "safebrowsing.enabled": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        arguments = [
            "--lang=pt-BR",
            "--start-maximized",
            "--disable-notifications",
            "--disable-gpu",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-extensions",
            "--log-level=3",
        ]
        for argument in arguments:
            chrome_options.add_argument(argument)
        service = ChromeService()
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.short_wait = WebDriverWait(self.driver, 10)
        self.wait = WebDriverWait(self.driver, 20)
        self.actions = ActionChains(self.driver)

    # ===> Função: Realiza a busca do primeiro elemento que corresponde ao caminho usando querySelector, retorna um único WebElement ou None
    def _find_first_element(self, path_selectors):
        script = """
            let element = document;
            const selectors = arguments[0];
            for (let i = 0; i < selectors.length; i++) {
                if (element.shadowRoot) { element = element.shadowRoot; }
                element = element.querySelector(selectors[i]);
                if (!element) { return null; }
            }
            return element;
        """
        return self.wait.until(
            lambda driver: driver.execute_script(script, path_selectors),
            f"Timeout ao tentar encontrar o elemento no caminho Shadow DOM: {path_selectors}",
        )

    def _fazer_login(self):
        print("\n---> Acessando a página de login.")
        # ATENÇÃO: Substitua pela URL correta do portal do cliente.
        self.driver.get("https://portal.cliente.exemplo.com/login")
        self._verificar_parada()
        try:
            # --- Preenchendo o formulário de login ---
            campo_usuario = self.wait.until(
                EC.presence_of_element_located((By.ID, "contentForm:profileIdInput"))
            )
            campo_usuario.clear()
            campo_usuario.send_keys(self.login)
            print("Usuário inserido.")

            # --- Preenchendo o campo de senha ---
            campo_senha = self.wait.until(
                EC.presence_of_element_located((By.ID, "contentForm:passwordInput"))
            )
            campo_senha.clear()
            campo_senha.send_keys(self.senha)
            print("Senha inserida.")

            # --- Ativando o checkbox de autenticação forte, se necessário ---
            try:
                checkbox = self.short_wait.until(
                    EC.presence_of_element_located((By.ID, "contentForm:j_idt25_input"))
                )
                if not checkbox.is_selected():
                    checkbox.click()
                    print("Checkbox 'Strong authentication' ativado.")
            except TimeoutException:
                print("AVISO: Checkbox 'Strong authentication' não encontrado.")

            # --- Clicando no botão de login ---
            botao_login = self.wait.until(
                EC.element_to_be_clickable((By.ID, "contentForm:passwordLoginAction"))
            )
            botao_login.click()
            print("Botão de login clicado.")

            # --- Preenchendo o campo TOTP ---
            self._verificar_parada()
            campo_totp = self.wait.until(EC.presence_of_element_located((By.ID, "otp")))
            campo_totp.clear()
            campo_totp.send_keys(self.autenticador)
            print("Código TOTP inserido.")

            # --- Confirmando o login ---
            self._verificar_parada()
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

    def _renomear_ultimo_arquivo_baixado(self, pasta_download, novo_prefixo):
        time.sleep(3)
        arquivos = sorted(
            [os.path.join(pasta_download, f) for f in os.listdir(pasta_download)],
            key=os.path.getmtime,
        )
        if not arquivos:
            print(
                "AVISO: Nenhum arquivo encontrado na pasta de download para renomear."
            )
            return

        ultimo_arquivo = arquivos[-1]
        if ultimo_arquivo.endswith(".crdownload"):
            print("Aguardando finalização do download...")
            time.sleep(5)
            arquivos = sorted(
                [os.path.join(pasta_download, f) for f in os.listdir(pasta_download)],
                key=os.path.getmtime,
            )
            ultimo_arquivo = arquivos[-1]

        nome_original_sem_ext, extensao = os.path.splitext(
            os.path.basename(ultimo_arquivo)
        )
        partes = nome_original_sem_ext.split("_")

        if len(partes) > 2:
            sufixo_original = "_".join(partes[2:])
            nome_final = f"{novo_prefixo}_{sufixo_original}"
        else:
            nome_final = novo_prefixo

        novo_nome_completo = f"{nome_final}{extensao}"
        novo_caminho = os.path.join(pasta_download, novo_nome_completo)

        try:
            print(
                f"Renomeando '{os.path.basename(ultimo_arquivo)}' para '{novo_nome_completo}'..."
            )
            os.rename(ultimo_arquivo, novo_caminho)
            print("Arquivo renomeado com sucesso.")
        except Exception as e:
            print(f"ERRO ao renomear o arquivo: {e}")

    def _navegar_e_baixar_pdfs(self):
        print("\n---> Navegando para a tela de pedidos.")
        self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//div[normalize-space()='Online Orders Series Material (ONB)']",
                )
            )
        ).click()
        print("Acessada a seção 'Online Orders Series Material'.")
        time.sleep(2)

        # Abrindo a seção de pedidos em uma nova janela
        janela_principal = self.driver.window_handles[0]
        janela_pedidos = self.driver.window_handles[-1]
        self.driver.switch_to.window(janela_pedidos)
        print("Mudando para a nova janela de pedidos.")

        while True:
            self._verificar_parada()
            try:
                xpath_novo_pedido = "//tr[@title='Novo Pedido']/td[3]/a"
                primeiro_pedido_link = self.short_wait.until(
                    EC.presence_of_element_located((By.XPATH, xpath_novo_pedido))
                )

                linha_do_pedido = primeiro_pedido_link.find_element(
                    By.XPATH, "./ancestor::tr"
                )
                novo_codigo_elemento = linha_do_pedido.find_element(
                    By.XPATH, ".//td[4]"
                )
                novo_codigo_texto = novo_codigo_elemento.text.strip()
                novo_prefixo_formatado = (
                    novo_codigo_texto.replace(" ", "_")
                    .replace("/", "-")
                    .replace(".", "")
                )

                # --- Inciando processo de download do PDF ---
                time.sleep(0.5)
                print(
                    f"\nProcessando pedido: {primeiro_pedido_link.text.strip()} | Código para nome: {novo_codigo_texto}"
                )
                primeiro_pedido_link.click()
                time.sleep(2)

                links_pdf = self.driver.find_elements(
                    By.PARTIAL_LINK_TEXT, "Pedido em PDF"
                )
                if not links_pdf:
                    print("AVISO: Nenhum link de PDF encontrado. Pulando.")
                    self.driver.back()
                    time.sleep(2)
                    continue

                links_pdf[-1].click()
                print("Clicado no link do PDF.")
                time.sleep(2)

                janela_pdf = self.driver.window_handles[-1]
                self.driver.switch_to.window(janela_pdf)
                print("Acessada a nova janela do PDF.")

                # ---> Iniciando clique no botão de download dentro do PDF
                # ATENÇÃO: O seletor 'baseSvg' pode ser específico. Mantenha se funcionar.
                campo_download = "baseSvg"
                botao_download = self._find_first_element(campo_download)
                botao_download.click()
                print("Clicado no botão download do PDF.")

                self._renomear_ultimo_arquivo_baixado(
                    self.tratamento.caminho_pasta_pdf, novo_prefixo_formatado
                )

                self.driver.close()
                self.driver.switch_to.window(janela_pedidos)
                print("Fechada a janela do PDF e retornado para a lista de pedidos.")

            except TimeoutException:
                print("\nNão há mais 'Novos Pedidos' para processar.")
                break
            except StaleElementReferenceException:
                print("A página foi atualizada. Tentando encontrar o próximo pedido...")
                continue
            except Exception as e:
                print(f"Ocorreu um erro inesperado no loop: {e}. Tentando novamente...")
                self.driver.refresh()
                time.sleep(3)
                continue
        self.driver.close()

    def executar(self):
        pythoncom.CoInitialize()
        try:
            self._iniciar_driver()
            self._verificar_parada()
            self._fazer_login()
            self._verificar_parada()
            print("\nLogin concluído. Navegando para baixar os arquivos...")
            self._navegar_e_baixar_pdfs()
            print("\nDownloads finalizados.")

            return {
                "sucesso": True,
                "mensagem": "Download e renomeação de arquivos concluídos com sucesso!",
            }

        except InterruptedError as e:
            return {"sucesso": False, "mensagem": str(e)}
        except (TimeoutException, RuntimeError, WebDriverException) as e:
            messagebox.showerror("Erro de Automação", str(e))
            return {"sucesso": False, "mensagem": str(e)}
        except Exception as e:
            traceback.print_exc()
            return {"sucesso": False, "mensagem": f"Erro fatal na automação: {e}"}
        finally:
            if self.driver:
                self.driver.quit()
                print("Driver finalizado.")
            pythoncom.CoUninitialize()
