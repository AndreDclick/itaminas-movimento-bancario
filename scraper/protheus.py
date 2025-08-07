from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook
import time
from pathlib import Path
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
            time.sleep(15)
            usuario_input = self.page.locator('xpath=//*[@id="po-login[d932dfe5-40c2-f4b1-448b-a48bf77285dd]"]').click()
            usuario_input.fill(self.settings.USUARIO)
            senha_input = self.page.locator('xpath=//*[@id="po-password[830cc1d2-8645-9ae9-7efc-f54003d447c6]"]')
            senha_input.fill(self.settings.SENHA)
            butao_entrar = self.page.locator('button:has-text("Entrar")')
            logger.info("Clicking OK button...")
            butao_entrar.click()
        except Exception as e:
            logger.error(f"Failed to login: {e}")
            raise FormSubmitFailed(f"Failed to login: {e}")
    def run(self):
        results = []
        try:
            self.start_scraper()
            self.login()
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
