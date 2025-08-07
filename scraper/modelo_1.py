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

class Modelo_1:
    def __init__(self, page):
        self.page = page


    def execucao(self):
        logger.info('Iniciando Modelo 1')
        self.page.get_by_text("Relatorios (9)").click()
        self.page.get_by_text("Balancetes (34)").click()
        self.page.get_by_text("Modelo 1", exact=True).click()
        time.sleep()
        self.page.get_by_role("button", name="Confirmar").click()
        self.page.get_by_role("button", name="Fechar").click()