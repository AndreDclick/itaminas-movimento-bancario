"""
Arquivo movbancario.py
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.logger import configure_logger
from config.settings import Settings
from .utils import Utils
from scraper.conciliacao import Conciliacao
from .exceptions import DownloadFailed, TimeoutOperacional
from datetime import datetime, timedelta
from pathlib import Path
import time
import os

logger = configure_logger()

class MovBancaria(Utils):
    def __init__(self, page):
        super().__init__(page)  
        self.settings = Settings()
        self.conciliacao = Conciliacao()
        self.parametros_json = 'MovBancaria' 
        logger.info("Movimentação bancária inicializada")

    def _definir_locators(self):
        """Centraliza os locators específicos da extração financeira"""
        self.locators = {
            # Navegação
            'menu_relatorios': self.page.get_by_text("Relatorios (13)"),
            'menu_movbancario': self.page.get_by_text("Movimento Bancario (25)"),
            'menu_extrato': self.page.get_by_text("Extrato Bancário", exact=True),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),
            'confirmar_moeda': self.page.get_by_text("Moedas"),

            'menu_pdf': self.page.get_by_role("button", name="PDF"),
            'text_paisagem': self.page.get_by_text("Paisagem"),
            'opcao_novo': self.page.locator("#COMP4560").get_by_role("combobox"),
            'outras_acoes': self.page.get_by_role("button", name="Outras Ações"),
            'parametros': self.page.get_by_text("Parâmetros"),
            
            # Locators do extrato bancário
            'do_bancos': self.page.locator("#COMP6012").get_by_role("textbox"),
            'agencia_banco': self.page.locator("#COMP6014").get_by_role("textbox"),
            'c_corrente_banco': self.page.locator("#COMP6016").get_by_role("textbox"),
            'da_data': self.page.locator("#COMP6018").get_by_role("textbox"),
            'ate_a_data': self.page.locator("#COMP6020").get_by_role("textbox"),
            'botao_ok': self.page.get_by_role("button", name="OK"),
            'botao_imprimir': self.page.get_by_role("button", name="Imprimir"),

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
            
            'banco_inv': self.page.get_by_text("Help: BCONOEXISTProblema:"),
            'nao': self.page.get_by_role('button', name='Não'),
        }
        logger.info("Seletores definidos")

    def _navegar_menu(self):
        
        try:
            try:
                self.locators['menu_relatorios'].wait_for(state="visible", timeout=18000)
            except PlaywrightTimeoutError:
                logger.error("Timeout ao aguardar menu_relatorios")
                raise TimeoutOperacional("Timeout na operação", operacao="aguardar menu_relatorios", tempo_limite=5)
            
            if not self.locators['menu_movbancario'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            
            if not self.locators['menu_movbancario'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            
            self.locators['menu_movbancario'].click()
            time.sleep(1)            
            if not self.locators['menu_extrato'].is_visible():
                self.locators['menu_movbancario'].click()
                time.sleep(1)
            
            self.locators['menu_extrato'].click()    
            time.sleep(1)
            
            self._confirmar_operacao()
            time.sleep(2)
            self._fechar_popup_se_existir()
            time.sleep(1)
            
            if not self.locators['popup_fechar'].is_visible():
                self._fechar_popup_se_existir()
                
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
                
        except TimeoutOperacional as e:
            logger.error(f"Timeout operacional: {e}")
            raise
        except Exception as e:
            logger.error("Falha na navegação")
            raise

    def _confirmar_moeda(self):
        time.sleep(3)
        if self.locators['confirmar_moeda'].is_visible():
            self.locators['botao_confirmar'].click()

    def _gerar_arquivo(self):
        try:
            self.locators['menu_pdf'].wait_for(state="visible", timeout=120000)
            logger.info("Botão 'PDF' visível")

            self.locators['menu_pdf'].click()
            logger.info("Botão 'PDF' clicado")
            time.sleep(0.5)
            self.locators['menu_pdf'].click()

            time.sleep(0.5)
            self.locators['text_paisagem'].click()
            time.sleep(0.5)
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
            
            do_banco = banco.get('do_banco')
            da_agencia = banco.get('da_agencia')
            da_conta = banco.get('da_conta')

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

            logger.info("Parâmetros preenchidos com sucesso")

        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")

    def _imprimir_e_baixar(self, banco, nome_banco):
        """Clica no botão de imprimir e baixa o arquivo com nome personalizado"""
        try:
            logger.info("Aguardando botão de impressão.")
            self.locators['imprimir_btn'].wait_for(state='visible', timeout=30000)
            time.sleep(2)
            
            # Criar diretório de downloads se não existir
            download_dir = Path(self.settings.DOWNLOADS_DIR)
            download_dir.mkdir(exist_ok=True)
            
            # Gerar nome do arquivo baseado no nome do banco
            data_atual = datetime.now().strftime("%Y%m%d")
            nome_arquivo = f"EXTRATO_{nome_banco.upper()}_{data_atual}.pdf"
            caminho_arquivo = download_dir / nome_arquivo
            
            # Verificar se o banco é inválido ANTES de tentar o download
            time.sleep(2)  # Aguardar um pouco para popup aparecer
            if self._verificar_banco_invalido():
                logger.info(f'Banco: {banco["do_banco"]}, agência: {banco["da_agencia"]}, conta: {banco["da_conta"]} é inválido')
                self.conciliacao.registrar_banco_invalido(
                    nome_banco,  
                    banco["do_banco"], 
                    banco["da_agencia"], 
                    banco["da_conta"]
                )
                return None
            
            # Esperar pelo download e salvar com nome personalizado
            with self.page.expect_download(timeout=300000) as download_info:
                self.locators['imprimir_btn'].click()
                logger.info("Botão download clicado")
                time.sleep(1)
                self._fechar_popup_se_existir()
                time.sleep(2)
                
                if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
                    self.locators['botao_sim'].click()
                
                time.sleep(3)
                
                # Verificar novamente se o banco é inválido APÓS clicar
                if self._verificar_banco_invalido():
                    logger.info(f'Banco: {banco["do_banco"]}, agência: {banco["da_agencia"]}, conta: {banco["da_conta"]} é inválido')
                    self.conciliacao.registrar_banco_invalido(
                        nome_banco,  # CORRIGIDO: passar nome_banco primeiro
                        banco["do_banco"], 
                        banco["da_agencia"], 
                        banco["da_conta"]
                    )
                    return None  # Retorna None para indicar banco inválido
                
                time.sleep(1)
                self._fechar_popup_se_existir()
                time.sleep(5)
                if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
                    self.locators['botao_sim'].click()
            
            download = download_info.value
            logger.info(f"Download iniciado: {download.suggested_filename}")
            
            # Salvar o arquivo com o nome personalizado
            download.save_as(caminho_arquivo)
            logger.info(f"Arquivo salvo como: {caminho_arquivo}")
            
            return caminho_arquivo
            
        except Exception as e:
            logger.error(f"Falha na impressão/download do pdf: {e}")
            # Em caso de erro, registrar como inválido
            self.conciliacao.registrar_banco_invalido(
                nome_banco,
                banco["do_banco"],
                banco["da_agencia"],
                banco["da_conta"]
            )
            return None

    def _verificar_banco_invalido(self):
        """Verifica se aparece mensagem de banco inválido"""
        try:
            # Tentar diferentes seletores para mensagem de banco inválido
            selectors = [
                'Help: BCONOEXISTProblema:'
            ]
            
            for selector in selectors:
                if self.page.is_visible(selector):
                    return True
                    
            return False
        except Exception as e:
            logger.warning(f"Banco inválido: {e}")
            return False

    def _processar_conta(self, banco, nome_banco):
        """Processa uma conta bancária individual completa."""
        try:
            logger.info(f'Processando banco: {nome_banco} - {banco["do_banco"]}, agência: {banco["da_agencia"]}, conta: {banco["da_conta"]}')
            
            self._navegar_menu()            
            self._confirmar_moeda()
            self._gerar_arquivo()
            self._preencher_parametros(banco)
            caminho_arquivo = self._imprimir_e_baixar(banco, nome_banco)
            
            if caminho_arquivo:  # Só processa se não for banco inválido (caminho_arquivo não é None)
                logger.info(f"✅ Banco {nome_banco} processado com sucesso - Arquivo: {caminho_arquivo.name}")
                status, si, sa, dif = self.conciliacao._processar_pdf(
                    Path(caminho_arquivo),
                    nome_banco, 
                    banco["do_banco"],
                    banco["da_agencia"],
                    banco["da_conta"]
                )
                logger.info(f"Conciliação [{nome_banco}] - Status: {status}, Inicial: {si}, Atual: {sa}, Diferença: {dif}")
                return caminho_arquivo
            else:
                logger.warning(f"⚠️ Banco {nome_banco} identificado como inválido")
                return None  # Banco inválido, não retorna arquivo
                    
        except Exception as e:
            error_msg = f"Falha no processamento do banco {nome_banco}"
            logger.error(f"{error_msg}: {str(e)}")
            # Registrar como inválido em caso de erro
            self.conciliacao.registrar_banco_invalido(
                nome_banco,
                banco["do_banco"],
                banco["da_agencia"],
                banco["da_conta"]
            )
            return None
        
    def execucao(self):
        """Fluxo principal de extração dos pdf."""
        try:
            # Definir os bancos
            bancos = {
                'banco_do_brasil': {
                    'da_conta': '118091',
                    'do_banco': '000',
                    'da_agencia': '0001'
                },
                'bradesco_sa': {
                    'da_conta': '105169',
                    'do_banco': '237',
                    'da_agencia': '0895'
                },
                'banco_bs2_sa': {
                    'da_conta': '10164',
                    'do_banco': '218',
                    'da_agencia': '0001'
                }
            }
            
            # Carregar os parâmetros do JSON
            parameters_path = self.settings.PARAMETERS_DIR
            self._carregar_parametros(parameters_path, self.parametros_json)
            
            # Processar cada banco
            arquivos_gerados = []
            for nome_banco, dados_banco in bancos.items():
                logger.info(f"Processando banco: {nome_banco}")
                caminho_arquivo = self._processar_conta(dados_banco, nome_banco)
                if caminho_arquivo:
                    arquivos_gerados.append(str(caminho_arquivo))
            
            return {
                'status': 'success',
                'message': f'{len(arquivos_gerados)}/{len(bancos)} contas processadas com sucesso',
                'arquivos_gerados': arquivos_gerados
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {
                'status': 'error', 
                'message': error_msg,
                'arquivos_gerados': []
            }