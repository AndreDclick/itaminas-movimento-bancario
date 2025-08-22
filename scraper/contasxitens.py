from playwright.sync_api import sync_playwright, TimeoutError 
from config.logger import configure_logger
from config.settings import Settings
from .utils import UtilsScraper
from datetime import date
from pathlib import Path

import calendar
import time

logger = configure_logger()

class Contas_x_itens(UtilsScraper):
    def __init__(self, page):  
        """Inicializa o Contas X Itens com a página do navegador"""
        self.page = page
        self._definir_locators()
        self.settings = Settings() 
        self.parametros_json = 'contasxitens'
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

    def _preencher_parametros(self, conta):
        logger.info(f"Usando chave JSON: {self.parametros_json}")
        # Resolver valores dinâmicos
        # input_data_inicial = self._resolver_valor(self.parametros.get('data_inicial'))
        # input_data_final = self._resolver_valor(self.parametros.get('data_final'))
        input_data_inicial =self.parametros.get('data_inicial')
        input_data_final = self.parametros.get('data_final')
        input_folha_inicial = self.parametros.get('folha_inicial')
        input_desc_moeda = self.parametros.get('desc_moeda')
        input_imprime_saldo = self.parametros.get('imprime_saldo')
        input_data_lucros = self.parametros.get('data_lucros')
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
            self.locators['conta_inicial'].fill(conta)
            time.sleep(0.5) 
            self.locators['conta_final'].click()
            self.locators['conta_final'].fill(conta)  
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
            self.locators['pula_pagina'].click()
            self.locators['pula_pagina'].select_option("0")
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

    def _gerar_planilha(self):
        """Gera e baixa a planilha """
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
            logger.info(f"Botão download clicado")
            time.sleep(2)
            if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
                self.locators['botao_sim'].click()
            time.sleep(2)
            self._fechar_popup_se_existir()

            # Esperar pelo download com timeout aumentado
            # with self.page.expect_download(timeout=120000) as download_info:
            #     self.locators['botao_imprimir'].click()
            #     logger.info(f"Botão download clicado")
            #     time.sleep(2)
            #     if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
            #         self.locators['botao_sim'].click()
            #     time.sleep(2)
            #     self._fechar_popup_se_existir()
                
            
            # download = download_info.value
            # logger.info(f"Download iniciado: {download.suggested_filename}") 
            
            # Aguardar conclusão do download
            # download_path = download.path()
            # if download_path:
            #     settings = Settings()
            #     destino = Path(settings.CAMINHO_PLS) / settings.PLS_CONTAS_X_ITENS
            #     destino.parent.mkdir(parents=True, exist_ok=True)
                
                
            #     download.save_as(destino)
            #     logger.info(f"Arquivo Contas x itens salvo em: {destino}")
            # else:
            #     logger.error("Download falhou - caminho não disponível")
            
            # Verificar se há botão de confirmação (se necessário)
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=50000)
        except Exception as e:
            logger.error(f"Falha na geração da planilha: {e}")
            raise

    def _processar_conta(self, conta):
        """Processa uma conta individual"""
        try:
            logger.info(f'Processando conta: {conta}')
            self._carregar_parametros('parameters.json', self.parametros_json)
            self._navegar_menu()
            time.sleep(1) 
            self._confirmar_operacao()  
            time.sleep(1) 
            self._fechar_popup_se_existir()  
            self._preencher_parametros(conta)  
            self._selecionar_filiais()  
            self._gerar_planilha()
            logger.info(f"✅ Conta {conta} processada com sucesso")
            
        except Exception as e:
            logger.error(f"❌ Falha no processamento da conta {conta}: {str(e)}")
            raise

    def execucao(self):
        """Fluxo principal de execução para todas as contas"""
        try:
            contas = ["10106020001", "20102010001"]
            
            for conta in contas:
                self._processar_conta(conta)
                
            return {
                'status': 'success',
                'message': f'Todas as {len(contas)} contas processadas com sucesso'
            }
                
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}