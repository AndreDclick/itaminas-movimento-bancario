from playwright.sync_api import sync_playwright, TimeoutError 
from config.logger import configure_logger
from .utils import UtilsScraper

import time

logger = configure_logger()

class Modelo_1(UtilsScraper):
    def __init__(self, page):  
        """Inicializa o Modelo 1 com a página do navegador"""
        self.page = page
        self._definir_locators()
        logger.info("Modelo_1 inicializado")

    def _definir_locators(self):
        """Centraliza apenas os locators específicos do Modelo 1"""
        self.locators = {
            # submenu
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),
            'submenu_balancetes': self.page.get_by_text("Balancetes (34)"),
            'opcao_modelo1': self.page.get_by_text("Modelo 1", exact=True),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),

            # parametros
            'data_inicial': self.page.locator("#COMP4512").get_by_role("textbox"),
            'data_final': self.page.locator("#COMP4514").get_by_role("textbox"),
            'conta_inicial':self.page.locator("#COMP4516").get_by_role("textbox"),
            'conta_final':self.page.locator("#COMP4518").get_by_role("textbox"),
            'data_lucros_perdas': self.page.locator("#COMP4556").get_by_role("textbox"),
            'grupos_receitas_despesas': self.page.locator("#COMP4562").get_by_role("textbox"),
            'data_sid_art': self.page.locator("#COMP4564").get_by_role("textbox"),
            'num_linha_balancete': self.page.locator("#COMP4566").get_by_role("textbox"),
            'desc_moeda': self.page.locator("#COMP4568").get_by_role("textbox"),
            'selec_filiais': self.page.locator("#COMP4570").get_by_role("combobox"),
            'botao_ok': self.page.locator('button:has-text("Ok")'),

            # Gerar planilha
            'aba_planilha': self.page.get_by_role("button", name="Planilha"),
            'formato': self.page.locator("#COMP4547").get_by_role("combobox"),
            'botao_imprimir': self.page.get_by_role("button", name="Imprimir"),
            'botao_sim': self.page.get_by_role("button", name="Sim")
        }
        logger.info("Seledores definidos")


    def _navegar_menu(self):
        """navegação no menu"""
        try:
            logger.info("Iniciando navegação no menu...")
            # Espera o menu principal estar disponível            
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=10000)
            self.locators['menu_relatorios'].click()
            logger.info("Menu Relatórios clicado")
            time.sleep(1)  
            if not self.locators['submenu_balancetes'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            self.locators['submenu_balancetes'].click()
            logger.info("Submenu Balancetes clicado")
            time.sleep(1)
            self.locators['opcao_modelo1'].wait_for(state="visible")
            self.locators['opcao_modelo1'].click()
            logger.info("Modelo 1 selecionada")
            
        except Exception as e:
            logger.error(f"Falha na navegação do menu: {e}")
            
            raise

    def _preencher_parametros(self):
        input_data_inicial = '01/04/2025'
        input_data_final = '30/04/2025'
        input_conta_inicial = ''
        input_conta_final = 'ZZZZZZZZZZZZZZZZZZZZ'
        input_data_lucros_perdas = ''
        input_grupos_receitas_despesas = '3456'
        input_data_sid_art = '31/12/2024'
        input_num_linha_balancete = '99'
        input_desc_moeda = '01'
        try:
            # parâmetros
            self.locators['data_inicial'].wait_for(state="visible")
            self.locators['data_inicial'].click()
            self.locators['data_inicial'].fill(input_data_inicial)
            time.sleep(0.5) 
            self.locators['data_final'].click()
            self.locators['data_final'].fill(input_data_final)
            time.sleep(0.5) 
            self.locators['conta_inicial'].click()
            self.locators['conta_inicial'].fill(input_conta_inicial)
            time.sleep(0.5) 
            self.locators['conta_final'].click()
            self.locators['conta_final'].fill(input_conta_final)
            time.sleep(0.5) 
            self.locators['data_lucros_perdas'].click()
            self.locators['data_lucros_perdas'].fill(input_data_lucros_perdas)
            time.sleep(0.5) 
            self.locators['grupos_receitas_despesas'].click()
            self.locators['grupos_receitas_despesas'].fill(input_grupos_receitas_despesas)
            time.sleep(0.5) 
            self.locators['data_sid_art'].click()
            self.locators['data_sid_art'].fill(input_data_sid_art)
            time.sleep(0.5) 
            self.locators['num_linha_balancete'].click()
            self.locators['num_linha_balancete'].fill(input_num_linha_balancete)
            time.sleep(0.5) 
            self.locators['desc_moeda'].click()
            self.locators['desc_moeda'].fill(input_desc_moeda)
            time.sleep(0.5)
            self.locators['selec_filiais'].click()
            time.sleep(0.5)
            self.locators['selec_filiais'].select_option("0")
            time.sleep(0.5)
            self.locators['botao_ok'].click()
        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")
            raise

    def _gerar_planilha (self):
        try: 
            self.locators['aba_planilha'].wait_for(state="visible")
            self.locators['aba_planilha'].click()
            time.sleep(1) 
            self.locators['formato'].select_option("3")
            time.sleep(1) 
            self.locators['botao_imprimir'].click()
            time.sleep(5)
            if self.locators['botao_sim'].is_visible():
                self.locators['botao_sim'].click()
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=100000)
        except Exception as e:
            logger.error(f"Falha na escolha impressão de planilha {e}")
            raise

    def execucao(self):
        """Fluxo principal de execução"""
        try:
            logger.info('Iniciando execução do Modelo 1')
            self._navegar_menu()
            self._confirmar_operacao()
            self._fechar_popup_se_existir()
            self._preencher_parametros()
            self._selecionar_filiais()
            self._gerar_planilha()
            logger.info("✅ Modelo 1 executado com sucesso")
            return {
                'status': 'success',
                'message': 'Modelo 1 completo'
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}