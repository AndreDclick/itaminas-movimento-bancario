from playwright.sync_api import Page
from config.logger import configure_logger

import time

logger = configure_logger()

class UtilsScraper:
    def __init__(self, page: Page):
        self.page = page
        self._definir_locators()

    def _definir_locators(self):
            """Centraliza todos os locators como variáveis"""
            self.locators = {
                'popup_fechar': self.page.get_by_role("button", name="Fechar"),
                'botao_confirmar': self.page.get_by_role("button", name="Confirmar"), 
                'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>")
            }

    def _fechar_popup_se_existir(self):
        """Método reutilizável para fechar popups"""
        try:
            time.sleep(3)
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
        except Exception as e:
            logger.warning(f" Erro ao verificar popup: {e}")

    def _confirmar_operacao(self):
        """confirmação da operação"""
        try:
            time.sleep(3)
            self.locators['botao_confirmar'].click()
            logger.info("operação confirmada")
            self._fechar_popup_se_existir()
        except Exception as e:
            logger.error(f"Falha na confirmação: {e}")
            raise  

    def _selecionar_filiais(self):
        """Seleção Filiais"""
        try: 
            time.sleep(3)
            if self.locators['botao_marcar_filiais'].is_visible():
                self.locators['botao_marcar_filiais'].click()
                time.sleep(1)
                self.locators['botao_confirmar'].click()
        except Exception as e:
            logger.error(f"Falha na escolha de filiais {e}")
            raise