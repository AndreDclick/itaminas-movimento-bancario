"""
Arquivo movbancario.py
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.logger import configure_logger
from config.settings import Settings
from .utils import Utils
from .exceptions import DownloadFailed, TimeoutOperacional
from datetime import datetime, timedelta
from pathlib import Path

import calendar
import time
import os
from pathlib import Path

logger = configure_logger()

# Automação e extração dos dados financeiros no sistema protheus (navegação e download).
class MovBancaria(Utils):
    # Inicialização e seleção dos seletores da interface, para carregas as configurações.
    def __init__(self, page):
        super().__init__(page)  
        self.settings = Settings()
        self.parametros_json = 'MovBancaria' 
        logger.info("Movimentação bancária inicializada")

    # armazenamento dos seletores utilizados na automação para facilitar caso haja mudanças.
    def _definir_locators(self):
        """Centraliza os locators específicos da extração financeira"""
        self.locators = {
            # Navegação
            'menu_relatorios': self.page.get_by_text("Relatorios (13)"),
            # 'menu_financeiro': self.page.get_by_text("Financeiro (2)"),
            'menu_movbancario': self.page.get_by_text("Movimento Bancario (25)"),
            'menu_extrato': self.page.get_by_text("Extrato Bancário"),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),
            'confirmar_moeda': self.page.get_by_text("Moedas"),

            'menu_pdf': self.page.get_by_role("button", name="PDF"),
            'text_paisagem': self.page.get_by_text("Paisagem"),
            'opcao_novo': self.page.locator("#COMP4560").get_by_role("combobox"),
            'outras_acoes': self.page.get_by_role("button", name="Outras Ações"),
            'parametros': self.page.get_by_text("Parâmetros"),
            
            # Locators do extrato bancário (não comentados)
            'do_bancos': self.page.locator("#COMP6012").get_by_role("textbox"),
            'agencia_banco': self.page.locator("#COMP6014").get_by_role("textbox"),
            'c_corrente_banco': self.page.locator("#COMP6016").get_by_role("textbox"),
            'da_data': self.page.locator("#COMP6018").get_by_role("textbox"),
            'ate_a_data': self.page.locator("#COMP6020").get_by_role("textbox"),
            'botao_ok': self.page.get_by_role("button", name="OK"),
            'botao_imprimir': self.page.get_by_role("button", name="Imprimir"),

            # Janela "Posição dos Títulos a Pagar"
            'planilha': self.page.get_by_role("button", name="Planilha"),
            'tipo_de_planilha': self.page.locator('#COMP4547').get_by_role('combobox'),
            'outras_acoes': self.page.get_by_role('button', name='Outras Ações'),
            'parametros_menu': self.page.get_by_text('Parâmetros'),
            'imprimir_btn': self.page.get_by_role('button', name='Imprimir'),
            'botao_sim': self.page.get_by_role("button", name="Sim"),

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

    
    # navegação pela página e tratamento de pop ups e confirmações.
    def _navegar_e_configurar_planilha(self):
        """Navega para a tela de Títulos a Pagar e configura a extração para planilha."""
        try:
            try:
                self.locators['menu_relatorios'].wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError:
                logger.error("Timeout ao aguardar menu_relatorios")
                raise TimeoutOperacional("Timeout na operação", operacao="aguardar menu_relatorios", tempo_limite=10)
            self.locators['menu_relatorios'].click()
            logger.info("Iniciando navegação no menu...")
            time.sleep(2)  
            if not self.locators['menu_movbancario'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            self.locators['menu_movbancario'].click()
            time.sleep(1)
            try:
                self.locators['menu_extrato'].wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError:
                logger.error("Timeout ao aguardar menu_extrato")
                # Exception TimeoutOperacional
                raise TimeoutOperacional("Timeout na operação", operacao="aguardar menu_extrato", tempo_limite=10)
            self.locators['menu_extrato'].click()    
            self._confirmar_operacao()
            time.sleep(2)
            self._fechar_popup_se_existir()
            time.sleep(1)
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
        # Exception TimeoutOperacional
        except TimeoutOperacional as e:
            logger.error(f"Timeout operacional: {e}")
            raise
        except Exception as e:
            logger.error("Falha na navegação ou configuração da planilha")
            raise

    
    def _confirmar_moeda(self):
        time.sleep(3)
        if self.locators['confirmar_moeda'].is_visible():
                self.locators['botao_confirmar'].click()

    # navegação para escolha do tipo de planilha que deve ser criada.
    def _criar_planilha(self):
        try:
            # Aguarda o botão PDF estar visível
            self.locators['menu_pdf'].wait_for(state="visible", timeout=120000)
            logger.info("Botão 'PDF' visível")

            # Clica no botão PDF para abrir as opções de impressão
            self.locators['menu_pdf'].click()
            logger.info("Botão 'PDF' clicado")
            time.sleep(0.5)
            self.locators['menu_pdf'].click()

            time.sleep(0.5)
            # Usa o método correto para marcar a checkbox
            self.locators['text_paisagem'].click()
            time.sleep(0.5)
            # Usa o método correto para marcar a checkbox
            self.locators['text_paisagem'].click()
            logger.info("Botão paisagem clicado")
            time.sleep(0.5)
            self.locators['opcao_novo'].click()
            self.locators['opcao_novo'].select_option("1")
            time.sleep(0.5)
            self.locators['outras_acoes'].click()
            time.sleep(0.5)
            self.locators['parametros_menu'].click()

        except PlaywrightTimeoutError:
            logger.error("Timeout ao aguardar elemento de relatório")
            raise TimeoutOperacional("Timeout na operação", operacao="aguardar botão/checkbox", tempo_limite=10)
        except Exception as e:
            logger.error(f"Falha na escolha de impressão: {e}")
            raise

    def _preencher_parametros(self, banco):
        try:
            logger.info(f"Usando chave JSON: {self.parametros_json}")
            #   Extrair dados do banco passado como parâmetro
            do_banco = banco.get('do_banco')
            da_agencia = banco.get('da_agencia')
            da_conta = banco.get('da_conta')

            # Valores carregados do JSON (novo schema)
            input_da_data = self.parametros.get('da_data')
            input_ate_a_data = self.parametros.get('ate_a_data')

            logger.info(f"Preenchendo parâmetros - Banco: {do_banco}, Agência: {da_agencia}, Conta: {da_conta}")

            self.locators['do_bancos'].wait_for(state="visible")
            self.locators['do_bancos'].click()
            self.locators['do_bancos'].fill(do_banco)
            time.sleep(0.5) 
            
            self.locators['agencia_banco'].click()
            self.locators['agencia_banco'].fill(da_agencia)
            time.sleep(0.5) 
            
            self.locators['c_corrente_banco'].click()
            self.locators['c_corrente_banco'].fill(da_conta)
            time.sleep(0.5)

            self.locators['da_data'].click()
            self.locators['da_data'].fill(input_da_data)
            time.sleep(0.5)

            self.locators['ate_a_data'].click()
            self.locators['ate_a_data'].fill(input_ate_a_data)
            time.sleep(0.5)
            try:
                self.locators['ok_btn'].click()
            except Exception:
                try:
                    self.page.get_by_role('button', name='OK').click()
                except Exception:
                    self.page.get_by_text('OK').click()

            logger.info("Parâmetros preenchidos com sucesso (novo schema)")

        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")

    # processo de impressão e download da planilha, salvando-a no local determinado. Tratando possíveis falhas no download.
    def _imprimir_e_baixar(self):
        """Clica no botão de imprimir e baixa o arquivo"""
        try:
            logger.info("Aguardando botão de impressão.")
            self.locators['imprimir_btn'].wait_for(state='visible', timeout=30000)
            time.sleep(2)
            
            # Esperar pelo download
            with self.page.expect_download(timeout=300000) as download_info:
                self.locators['imprimir_btn'].click()
                logger.info(f"botão download clicado")
                time.sleep(2)
                if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
                    self.locators['botao_sim'].click()
                    time.sleep(2)
                self._fechar_popup_se_existir()
                self._selecionar_filiais()
            self._confirmar_filiais()
            
            download = download_info.value
            logger.info(f"Download iniciado: {download.suggested_filename}")
            
        except Exception as e:
            logger.error(f"Falha na impressão/baixar da planilha: {e}")
            raise
    
    # confirmação das filiais a serem incluídas na planilha, tratando pop-ups e confirmações.
    def _confirmar_filiais(self):
        try:
            time.sleep(2) 
            if self.locators['nao'].is_visible():
                time.sleep(1)             
                self.locators['nao'].click()
                logger.info("Botão 'Não' clicado")
        except Exception as e:
            logger.error(f"Falha ao clicar no botão 'Não': {e}")

    def _processar_conta(self, banco):
        """
        Processa uma conta bancária individual completa.
        Args:
            banco (dict): Dicionário com os dados do banco (conta, banco, agência)
        """
        try:
            logger.info(f'Processando banco: {banco["do_banco"]}, agência: {banco["da_agencia"]}, conta: {banco["da_conta"]}')
            
            # Atribuir os valores às variáveis que serão usadas no preenchimento
            self.do_banco = banco["do_banco"]
            self.da_agencia = banco["da_agencia"] 
            self.da_conta = banco["da_conta"]
            
            self._navegar_e_configurar_planilha()            
            self._confirmar_moeda()
            self._criar_planilha()
            self._preencher_parametros(banco)
            self._imprimir_e_baixar()
            logger.info("Extração da planilha financeira executada com sucesso")
            self._gerar_planilha(banco)
            logger.info(f"✅ Banco {banco['do_banco']} processado com sucesso")
            
        except Exception as e:
            error_msg = f"Falha no processamento do banco {banco.get('do_banco', 'N/A')}"
            logger.error(f"{error_msg}: {str(e)}")

    # fluxo principal de execução da extração financeira, iniciando da navegação até o download da planilha.
    def execucao(self):
        """Fluxo principal de extração de planilha financeira."""
        try:
            # Definir os bancos como dicionário com os dados reais
            bancos = {
                'banco_bloqueio_judicial': {
                    'da_conta': '000000',
                    'do_banco': '000',
                    'da_agencia': '0000'
                },
                'bradesco_sa': {
                    'da_conta': '105169',
                    'do_banco': '237',
                    'da_agencia': '0895'
                },
                'banco_bmg': {
                    'da_conta': '17003248',
                    'do_banco': '318',
                    'da_agencia': '0005'
                }
            }
            
            # Carregar os parâmetros do JSON
            parameters_path = self.settings.PARAMETERS_DIR
            self._carregar_parametros(parameters_path, self.parametros_json)
            
            # Iterar sobre os bancos
            for nome_banco, dados_banco in bancos.items():
                logger.info(f"Processando banco: {nome_banco}")
                self._processar_conta(dados_banco)
            
            return {
                'status': 'success',
                'message': f'Todas as {len(bancos)} contas processadas com sucesso'
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}