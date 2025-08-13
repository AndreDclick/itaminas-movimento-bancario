from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.logger import configure_logger
from .utils import UtilsScraper
from datetime import datetime, timedelta

import calendar
import time

logger = configure_logger()

class ExtracaoFinanceiro(UtilsScraper):
    def __init__(self, page):
        self.page = page
        self._definir_locators()
        
        logger.info("Financeiro inicializada")

    def _definir_locators(self):
        """Centraliza os locators específicos da extração financeira"""
        self.locators = {
            # Navegação
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),
            'menu_financeiro': self.page.get_by_text("Financeiro (2)"),
            'menu_titulos_a_pagar': self.page.get_by_text("Títulos a Pagar", exact=True),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),

            # Janela "Posição dos Títulos a Pagar"
            'planilha': self.page.get_by_role("button", name="Planilha"),
            'tipo_de_planilha': self.page.locator('#COMP4547').get_by_role('combobox'),
            'outras_acoes': self.page.get_by_role('button', name='Outras Ações'),
            'parametros_menu': self.page.get_by_text('Parâmetros'),
            'imprimir_btn': self.page.get_by_role('button', name='Imprimir'),

            # Janela de Parâmetros
            'do_vencimento': self.page.locator('#COMP6024').get_by_role('textbox'),
            'ate_o_vencimento': self.page.locator('#COMP6026').get_by_role('textbox'),
            'da_emissao': self.page.locator('#COMP6036').get_by_role('textbox'),
            'ate_a_emissao': self.page.locator('#COMP6038').get_by_role('textbox'),
            'da_data_contabil': self.page.locator('#COMP6046').get_by_role('textbox'),
            'ate_a_data_contabil': self.page.locator('#COMP6048').get_by_role('textbox'),
            'data_base': self.page.locator('#COMP6076').get_by_role('textbox'),
            'ok_btn': self.page.get_by_role('button', name='OK'),
            
            # Janela de Seleção de Filiais
            'selecao_filiais_janela': self.page.get_by_text('Seleção de filiais'),
            'matriz_filial_checkbox': self.page.get_by_text('Matriz e Filial'), # Se houver checkbox para isso
            'marcar_todos_btn': self.page.get_by_role('button', name='Marca Todos - <F4>'),
            'confirmar': self.page.get_by_role('button', name='Confirmar'),

            #Janela confirmar filiais
            'nao': self.page.get_by_role('button', name='Não'),
        }
        logger.info("Seletores definidos")

    def _navegar_e_configurar_planilha(self):
        """Navega para a tela de Títulos a Pagar e configura a extração para planilha."""
        try:
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=10000)
            self.locators['menu_relatorios'].click()
            logger.info("Iniciando navegação no menu...")
            
            time.sleep(1)  
            if not self.locators['menu_financeiro'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            self.locators['menu_financeiro'].click()
            time.sleep(1)
            self.locators['menu_titulos_a_pagar'].wait_for(state="visible")
            self.locators['menu_titulos_a_pagar'].click()    
            self._confirmar_operacao()
            self._fechar_popup_se_existir()
        except PlaywrightTimeoutError:
            logger.error("Falha na navegação ou configuração da planilha")
            raise

    def _criar_planilha (self):
        try: 
            self.locators['planilha'].wait_for(state="visible")
            time.sleep(1)
            self.locators['planilha'].click()
            time.sleep(1)
            self.locators['tipo_de_planilha'].select_option("3")
            time.sleep(1)    
        except Exception as e:
            logger.error(f"Falha na escolha impressão de planilha {e}")
            raise

    def _outras_acoes(self):
        """Método para lidar com outras ações."""
        try:
            logger.info("Acessando outras ações")
            # Na opção "Outras Ações", selecionar "Parâmetros" 
            self.locators['outras_acoes'].click()
            self.locators['parametros_menu'].click()
            self.locators['imprimir_btn'].click()
            time.sleep(5)
        except Exception as e:
            logger.error(f"Falha ao acessar outras ações: {e}")
            raise

    def fechamento_mes(self):
        hoje = datetime.today()
        mes_passado = hoje.month - 1 if hoje.month > 1 else 12
        ano_mes_passado = hoje.year if hoje.month > 1 else hoje.year - 1
        ultimo_dia = calendar.monthrange(ano_mes_passado, mes_passado)[1]
        data_formatada = datetime(ano_mes_passado, mes_passado, ultimo_dia).strftime("%d/%m/%Y")
        return data_formatada

    def _preencher_parametros(self):
        input_do_vencimento = '01/01/2000'
        input_ate_o_vencimento = '31/12/2050'
        input_da_emissao = '01/01/2000'
        input_ate_a_emissao = '31/12/2050'
        input_da_data_contabil = '01/01/2020'
        # input_ate_a_data_contabil = datetime.now().strftime("%d/%m/%Y")
        input_ate_a_data_contabil = self.fechamento_mes()
        # input_ate_a_data_contabil = '31/07/2025'
        input_data_base = datetime.now().strftime("%d/%m/%Y")
        # input_data_base = '30/04/2025'
        
        try:
            # parâmetros
            self.locators['do_vencimento'].wait_for(state="visible")
            self.locators['do_vencimento'].click()
            self.locators['do_vencimento'].fill(input_do_vencimento)
            time.sleep(0.5)
            self.locators['ate_o_vencimento'].click()
            self.locators['ate_o_vencimento'].fill(input_ate_o_vencimento)
            time.sleep(0.5)
            self.locators['da_emissao'].click()
            self.locators['da_emissao'].fill(input_da_emissao)
            time.sleep(0.5)
            self.locators['ate_a_emissao'].click()
            self.locators['ate_a_emissao'].fill(input_ate_a_emissao)
            time.sleep(0.5)
            self.locators['da_data_contabil'].click()
            self.locators['da_data_contabil'].fill(input_da_data_contabil)
            time.sleep(0.5)
            self.locators['ate_a_data_contabil'].click()
            self.locators['ate_a_data_contabil'].fill(input_ate_a_data_contabil)
            time.sleep(0.5)
            self.locators['data_base'].click()
            self.locators['data_base'].fill(input_data_base)
            time.sleep(0.5)
            self.locators['ok_btn'].click()
            logger.info("Parâmetros preenchidos com sucesso")
        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")
            raise

    def _imprimir_e_baixar(self):
        """Clica no botão de imprimir, fecha o popup e gerencia o download."""
        try:
            logger.info("Aguardando botão de impressão.")
            self.locators['imprimir_btn'].wait_for(state='visible')
            self.locators['imprimir_btn'].click()
            self._fechar_popup_se_existir()

        except PlaywrightTimeoutError:
            logger.error("Falha na impressão ou download: Tempo esgotado")
            raise
        except Exception as e:
            logger.error(f"Falha na impressão ou download: {e}")
            raise

    def _confirmar_filiais(self):
        try:
            if self.locators['nao'].is_visible():            
                self.locators['nao'].click()
                logger.info("Botão 'Não' clicado")
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=100000)
        except Exception as e:
            logger.error(f"Falha ao clicar no botão 'Não': {e}")

    def execucao(self):
        """Fluxo principal de extração de planilha financeira."""
        try:
            logger.info('Iniciando extração da planilha financeira - Títulos a Pagar')
            self._navegar_e_configurar_planilha()
            self._criar_planilha()
            self._outras_acoes()
            self._preencher_parametros()
            self._imprimir_e_baixar()
            self._selecionar_filiais()
            self._confirmar_filiais()
            logger.info("Extração da planilha financeira executada com sucesso")
            return {
                'status': 'success',
                'message': 'Financeiro completo'
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}