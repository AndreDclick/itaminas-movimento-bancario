from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook
import time
from pathlib import Path
from .modelo_1 import Modelo_1
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import (BrowserClosedError, DownloadFailed,
                        FormSubmitFailed, InvalidDataFormat, ResultsSaveError)

logger = configure_logger()

class ProtheusScraper:
    def __init__(self, settings=Settings()):
        self.settings = settings
        self.playwright = sync_playwright().start()
        self._initialize_browser()
        logger.info("Browser initialized")

    def _initialize_browser(self):
        self.browser = self.playwright.chromium.launch(
            headless=self.settings.HEADLESS,
            args=["--start-maximized"],
            channel="msedge"
        )
        self.context = self.browser.new_context(no_viewport=True)
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.settings.TIMEOUT)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        time.sleep(self.settings.SHUTDOWN_DELAY)
        self.context.close()
        self.browser.close()
        self.playwright.stop()
        logger.info("Browser closed")

    def start_scraper(self):
        try:
            logger.info(f"Navigating to URL: {self.settings.BASE_URL}")
            self.page.goto(self.settings.BASE_URL)
            butao_ok = self.page.locator('button:has-text("Ok")')
            logger.info("Clicking OK button...")
            butao_ok.click()
        except Exception as e:
            logger.error(f"Failed to start scraper: {e}")
            raise FormSubmitFailed(f"Failed to start scraper: {e}")
        

    def login(self):
        try:
            # Definição dos locators (variáveis)
            iframe_locator = "iframe"
            campo_usuario_locator = self.page.frame_locator(iframe_locator).get_by_placeholder("Ex. sp01\\nome.sobrenome")
            campo_senha_locator = self.page.frame_locator(iframe_locator).get_by_label("Insira sua senha")
            botao_entrar_locator = self.page.frame_locator(iframe_locator).get_by_role("button", name="Entrar")
            # botao_fechar_locator = self.page.frame_locator(iframe_locator).get_by_role("button", name="Fechar")
            logger.info("Aguardando página de login...")
            self.page.wait_for_selector(iframe_locator, state="visible")
            
            logger.info("Preenchendo usuário...")
            campo_usuario_locator.fill(self.settings.USUARIO)
            
            logger.info("Preenchendo senha...")
            campo_senha_locator.fill(self.settings.SENHA)
            
            logger.info("Clicando em 'Entrar'...")
            botao_entrar_locator.click()

            time.sleep(5)
            self.page.wait_for_selector(iframe_locator, state="visible")
            logger.info("Clicando em 'Entrar'...")
            botao_entrar_locator.click()
            logger.info("Login concluído com sucesso.")

            time.sleep(5)
            self.page.get_by_role("button", name="Fechar").click()
            logger.info("popup fechado.")
            
        except Exception as e:
            logger.error(f"Falha no login: {e}")
            raise FormSubmitFailed(f"Falha no login: {e}")
        
    def run(self):
        results = []
        try:
            self.start_scraper()
            self.login()

            modelo_1 = Modelo_1(self, page)
            modelo_1 = execucao()

            results.append({
                'status': 'success',
                'message': 'Scraper executed successfully'
            })
        except Exception as e:
            results.append({
                'status': 'error',
                'message': str(e)
            })
        return results
