from playwright.sync_api import sync_playwright, TimeoutError 
from config.logger import configure_logger
from .utils import UtilsScraper
from datetime import date

import calendar
import time

logger = configure_logger()

class Contas_x_itens(UtilsScraper):
    def __init__(self, page):  
        """Inicializa o Contas X Itens com a página do navegador"""
        self.page = page
        self._definir_locators()
        logger.info("Contas_x_itens inicializado")

    def _definir_locators(self):
        """Centraliza apenas os locators específicos do Contas X Itens"""
        self.locators = {
            # submenu
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),
            'submenu_balancetes': self.page.get_by_text("Balancetes (34)"),
            'opcao_contas_x_itens': self.page.get_by_text("Contas X Itens", exact=True),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),

            # parametros
            'data_inicial': self.page.locator("#COMP4512").get_by_role("textbox"),
            'data_final': self.page.locator("#COMP4514").get_by_role("textbox"),
            'conta_inicial': self.page.locator("#COMP4516").get_by_role("textbox"),
            'conta_final': self.page.locator("#COMP4518").get_by_role("textbox"),
            'imprime_item': self.page.locator("#COMP4524").get_by_role("combobox"),
            'saldos_zerados': self.page.locator("#COMP4528").get_by_role("combobox"),
            'moeda': self.page.locator("#COMP4530").get_by_role("textbox"),
            'folha_inicial': self.page.locator("#COMP4532").get_by_role("textbox"),
            'imprime_saldos': self.page.locator("#COMP4534").get_by_role("textbox"),
            'imprime_coluna': self.page.locator("#COMP4546").get_by_role("combobox"),
            'imp_tot_cta': self.page.locator("#COMP4548").get_by_role("combobox"),
            'pula_pagina': self.page.locator("#COMP4550").get_by_role("combobox"),
            'salta_linha': self.page.locator("#COMP4552").get_by_role("combobox"),
            'imprime_valor':self.page.locator("#COMP4554").get_by_role("combobox"),
            'impri_cod_item': self.page.locator("#COMP4556").get_by_role("combobox"),
            'divide_por': self.page.locator("#COMP4558").get_by_role("combobox"),
            'impri_cod_conta': self.page.locator("#COMP4560").get_by_role("combobox"),
            'posicao_ant_lp': self.page.locator("#COMP4562").get_by_role("combobox"),
            'data_lucros': self.page.locator("#COMP4564").get_by_role("textbox"),
            'selec_filiais': self.page.locator("#COMP4566").get_by_role("combobox"),
            'botao_ok': self.page.get_by_role("button", name="OK"),
            
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
            
            time.sleep(2)  
            if not self.locators['submenu_balancetes'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            
            self.locators['submenu_balancetes'].click()
            logger.info("Submenu Balancetes clicado")
            time.sleep(1)
            
            self.locators['opcao_contas_x_itens'].wait_for(state="visible")
            self.locators['opcao_contas_x_itens'].click()
            logger.info("Contas x Itens selecionada")
            
        except Exception as e:
            logger.error(f"Falha na navegação do menu: {e}")
            
            raise

    def primeiro_e_ultimo_dia(self):
        hoje = date.today()
        mes_passado = hoje.month - 1 if hoje.month > 1 else 12
        ano_mes_passado = hoje.year if hoje.month > 1 else hoje.year - 1
        
        primeiro_dia = date(ano_mes_passado, mes_passado, 1).strftime("%d/%m/%Y")
        ultimo_dia_num = calendar.monthrange(ano_mes_passado, mes_passado)[1]
        ultimo_dia = date(ano_mes_passado, mes_passado, ultimo_dia_num).strftime("%d/%m/%Y")
        
        return primeiro_dia, ultimo_dia

    def _preencher_parametros(self):
        # primeiro, ultimo = self.primeiro_e_ultimo_dia()
        # input_data_inicial = primeiro
        # input_data_final = ultimo
        input_data_inicial = '01/04/2025'
        input_data_final = '30/04/2025'
        input_conta_inicial = '20102010001'
        input_conta_final = '20102010001'
        input_folha_inicial = '2'
        input_desc_moeda = '01'
        input_imprime_saldo = '1'
        input_data_lucros = '30/06/2024'
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
            self.locators['imprime_item'].click()     
            self.locators['imprime_item'].select_option("1")
            time.sleep(0.5) 
            self.locators['saldos_zerados'].click()  
            self.locators['saldos_zerados'].select_option("1")
            time.sleep(0.5)
            self.locators['moeda'].click()
            self.locators['moeda'].fill(input_desc_moeda)
            time.sleep(0.5) 
            self.locators['folha_inicial'].click()
            self.locators['folha_inicial'].fill(input_folha_inicial)
            time.sleep(0.5) 
            self.locators['imprime_saldos'].click()
            self.locators['imprime_saldos'].fill(input_imprime_saldo)
            time.sleep(0.5) 
            self.locators['imprime_coluna'].click()
            self.locators['imprime_coluna'].select_option("0")
            time.sleep(0.5)
            self.locators['imp_tot_cta'].click()
            self.locators['imp_tot_cta'].select_option("0")
            time.sleep(0.5)
            self.locators['salta_linha'].click()
            self.locators['salta_linha'].select_option("1")
            time.sleep(0.5)
            self.locators['imprime_valor'].click()
            self.locators['imprime_valor'].select_option("1")
            time.sleep(0.5)
            self.locators['impri_cod_item'].click()
            self.locators['impri_cod_item'].select_option("0")
            time.sleep(0.5)
            self.locators['divide_por'].click()
            self.locators['divide_por'].select_option("0")
            time.sleep(0.5)
            self.locators['impri_cod_conta'].click()
            self.locators['impri_cod_conta'].select_option("0")
            time.sleep(0.5)
            self.locators['posicao_ant_lp'].click()
            self.locators['posicao_ant_lp'].select_option("1")
            time.sleep(0.5)
            self.locators['data_lucros'].click()
            self.locators['data_lucros'].fill(input_data_lucros)
            time.sleep(0.5)
            self.locators['selec_filiais'].click()
            self.locators['selec_filiais'].select_option("0")
            time.sleep(0.5)
            self.locators['botao_ok'].click()
        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")
            raise

    def _gerar_planilha (self):
        try: 
            self.locators['aba_planilha'].wait_for(state="visible")
            time.sleep(1) 
            self.locators['aba_planilha'].click()
            time.sleep(1) 
            
            if not self.locators['formato'].is_visible():
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
            logger.info('Iniciando execução do Contas X Itens')
            self._navegar_menu()
            time.sleep(1)
            self._confirmar_operacao()
            time.sleep(1)
            self._fechar_popup_se_existir()
            self._preencher_parametros()
            self._selecionar_filiais()
            self._confirmar_operacao
            self._gerar_planilha()
            logger.info("✅ Contas X Itens executado com sucesso")
            return {
                'status': 'success',
                'message': 'Contas X Itens completo'
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}