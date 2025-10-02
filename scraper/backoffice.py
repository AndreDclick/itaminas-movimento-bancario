
"""
Módulo para conciliação manual no sistema Protheus
"""
import pymupdf    
import re
from pathlib import Path
from datetime import datetime,time
from typing import Optional, List, Tuple
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import (
    LoginProtheusError,
    BrowserClosedError,
    DownloadFailed,
    TimeoutOperacional,
    ExcecaoNaoMapeadaError
)
from .utils import Utils
# Configurar logger
logger = configure_logger()


class BackOffice(Utils):
    def __init__(self, page):
        super().__init__(page)  
        self.page = page
        self._definir_locators()
    
    def _definir_locators(self):
        
        logger.info("Definindo seletores Backoffice...")
        self.locators = {
            
            # NAVEGAÇÃO MENU LATERAL
            'atualizacoes': self.page.get_by_text('Atualizações (17)'),
            'mov_bancario':self.page.get_by_text('Movimento Bancario (6)'),
            'backoffice':self.page.get_by_text('Conciliador Backoffice'),
            'btn_confirmar':self.page.get_by_role('button', name='Confirmar'),
            'btn_fechar_menu':self.page.get_by_role('button', name='Fechar'),

            # NAVEGAÇÃO CONCILIADOR
            'frame_conciliador': self.page.get_by_role("menuitem", name="Conciliador"),
            'selecione': self.page.get_by_placeholder("Selecione uma configuração de"),
            'label_conciliacao': self.page.get_by_label("-Conciliação Bancária Manual"),
            'filtros': self.page.get_by_role("button", name="Ver Filtros"),
            'data_de': self.page.get_by_label("Data Dispon. de"),
            'data_ate': self.page.get_by_label("Data Dispon. até"),
            'label_banco': self.page.get_by_label("Banco igual a", exact=True),
            'label_agencia': self.page.get_by_label("Agencia igual a"),
            'label_conta': self.page.get_by_label("Conta Banco igual a"),
            'btn_aplicar': self.page.get_by_role("button", name="Aplicar"),
            
            # aba “Dados não Encontrados”
            'btn_n_encontrados': self.page.get_by_text("Dados não Encontrados"),

            # JANELA CONFIRMAÇÃO
            'confirmar_inclusao_linha': self.page.get_by_role('button', name='Confirmar'),
            'cancelar_inclusao_linha':self.page.get_by_role('button', name='Cancelar'),

            # BOTÕES PRINCIPAIS
            'btn_gravar': self.page.get_by_role('button', name='Gravar'),
            'btn_incluir': self.page.get_by_role('button', name='incluir'),

            # NAVEGAÇÃO ENTRE PÁGINAS
            'proxima_pagina': self.page.locator('[title=\"Próxima página\"]'),
            'primeira_pagina': self.page.locator('[title=\"Primeira página\"]'),
            'pagina_anterior': self.page.locator('[title=\"Página anterior\"]'),
            'ultima_pagina': self.page.locator('[title=\"Última página\"]'),

            # FILTROS
            'filtro_banco': self.page.locator('[id=\"CA8_BANCO\"]'),
            'filtro_agencia': self.page.locator('[id=\"CA8_AGENC\"]'),
            'filtro_conta': self.page.locator('[id=\"CA8_NUMCON\"]'),
            'filtro_data_inicio': self.page.locator('[id=\"CA8_DATCNC\"]'),
            'filtro_data_fim': self.page.locator('[id=\"CA8_DATCN2\"]'),
            'aplicar_filtro': self.page.get_by_role('button', name='Filtrar'),
            'limpar_filtro': self.page.get_by_role('button', name='Limpar'),

            # OUTROS CONTROLES
            'fechar_janela': self.page.get_by_role('button', name='Fechar'),
            'maximizar_janela': self.page.get_by_role('button', name='Maximizar'),
        }
        
        logger.info("Seletores para conciliação manual definidos")
    def _navegar_menu(self):
        
        try:
            
            self.locators['atualizacoes'].wait_for(state="visible", timeout=18000)
            
            if not self.locators['mov_bancario'].is_visible():
                self.locators['atualizacoes'].click()
                time.sleep(1)
            
            self.locators['mov_bancario'].click()
            time.sleep(1)            
            if not self.locators['backoffice'].is_visible():
                self.locators['mov_bancario'].click()
                time.sleep(1)
            
            self.locators['backoffice'].click()    
            time.sleep(1)
            
            self._confirmar_operacao()
            time.sleep(3)
            self._fechar_popup_se_existir()
            
        except TimeoutOperacional as e:
            logger.error(f"Timeout operacional: {e}")
            raise
        except Exception as e:
            logger.error("Falha na navegação")
            raise

    def _extraxao(self, banco, nome_banco):
        self.locators['frame_conciliador'].wait_for(state="visible", timeout=1000)
        time.sleep(0.5)
        self.locators['frame_conciliador'].click()
        time.sleep(1)
        self.locators['selecione'].click()
        time.sleep(1)
        self.locators['selecione'].fill("0024-Conciliação Bancária Manual")
        time.sleep(0.5)
        self.locators['label_conciliacao'].click()
        time.sleep(1)
        self.locators['filtros'].click()
        time.sleep(3)
        self.locators['data_de'].click()
        time.sleep(0.5)
        self.locators['data_de'].fill("15/09/2025")
        time.sleep(0.5)
        self.locators['data_ate'].click()
        time.sleep(0.5)
        self.locators['data_ate'].fill("15/09/2025")
        time.sleep(0.5)
        self.locators['label_banco'].click()
        time.sleep(0.5)
        self.locators['label_banco'].fill("15/09/2025")
        time.sleep(0.5)
        self.locators['label_agencia'].click()
        time.sleep(0.5)
        self.locators['label_agencia'].fill("15/09/2025")
        time.sleep(0.5)
        self.locators['label_conta'].click()
        time.sleep(0.5)
        self.locators['label_conta'].fill("15/09/2025")

        # aba “Dados não Encontrados”

        self.locators['btn_n_encontrados'].wait_for(state="visible", timeout=1000)
        time.sleep(0.5)
        self.locators['btn_n_encontrados'].click()
        
    def _processar_conta(self, banco, nome_banco):
        """
        Processa uma conta bancária individual completa.
        Args:
            banco (dict): Dicionário com os dados do banco
            nome_banco (str): Nome identificador do banco
        """
        try:
            logger.info(f'Processando banco: {nome_banco} - {banco["do_banco"]}, agência: {banco["da_agencia"]}, conta: {banco["da_conta"]}')
            
            self._navegar_menu()            
            self._confirmar_moeda()
            self._gerar_arquivo()
            self._preencher_parametros(banco)
            caminho_arquivo = self._imprimir_e_baixar(banco, nome_banco)
            
            if caminho_arquivo:
                logger.info(f"✅ Banco {nome_banco} processado com sucesso - Arquivo: {caminho_arquivo.name}")
                status, si, sa, dif = self.conciliacao._processar_pdf(
                    Path(caminho_arquivo),
                    banco["do_banco"],
                    banco["da_agencia"],
                    banco["da_conta"]
                )
                logger.info(f"Conciliação [{nome_banco}] - Status: {status}, Inicial: {si}, Atual: {sa}, Diferença: {dif}")
                return caminho_arquivo
            else:
                logger.warning(f"⚠️ Banco {nome_banco} não gerou arquivo (possivelmente inválido)")
                return None
                
        except Exception as e:
            error_msg = f"Falha no processamento do banco {nome_banco}"
            logger.error(f"{error_msg}: {str(e)}")
            return None

    def execucao(self):
        """Fluxo principal de extração dos pdf."""
        try:
            # Definir os bancos
            bancos = {
                'sicoob_itaminas': {
                    'da_conta': '31413',
                    'do_banco': '756',
                    'da_agencia': '4101'
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

