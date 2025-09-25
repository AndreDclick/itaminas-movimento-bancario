
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
        
        logger.info("Definindo seletores...")
        self.locators = {
            
            # NAVEGAÇÃO MENU LATERAL
            'atualizacoes': self.page.get_by_text('Atualizações (17)'),
            'mov_bancario':self.page.get_by_text('Movimento Bancario (6)'),
            'backoffice':self.page.get_by_text('Conciliador Backoffice'),
            'btn_confirmar':self.page.get_by_role('button', name='Confirmar'),
            'btn_fechar_menu':self.page.get_by_role('button', name='Fechar'),

            # NAVEGAÇÃO CONCILIADOR
            'frame_conciliador': self.page.get_by_role("menuitem", name="Conciliador"),
            'btn_include': self.page.get_by_role('button', name='incluir'),
            'frame_conciliacao': self.page.locator('[role="row"]').nth(1),
            'editar_conciliacao': self.page.get_by_role('button', name='Alterar [F3]'),
            'alterar_item': self.page.get_by_role('cell', name='000000000109972', exact=True),
            'btn_ver_mais': self.page.get_by_role('cell', name='*').locator('span'),
            'linha_conciliacao': self.page.get_by_role('cell', name='-R$ 4.356,85', exact=True),
            'btn_alterar': self.page.locator('[role="cell"]:has-text("Alterar"), [role="cell"]:has-text("Editar")'),

            # FORM DE PREENCHIMENTO
            'texto_descricao': self.page.get_by_text('CONCILIACAO OUTUBRO'),
            'check_item': self.page.locator('[id="\01TRB_C09_0010\"]'),
            'campo_descricao': self.page.locator('[id=\aCols\\[colsCA7\\]\\[nAt\\]\\[1\\]\]'),
            'btn_salvar': self.page.get_by_role('button', name='Gravar'),
            'btn_cancelar': self.page.get_by_role('button', name='Cancelar'),

            # DADOS DA TELA
            'data_conciliacao': "page.get_by_role('cell', name='20/10/2024', exact=True)",
            'historico_conciliacao': "page.get_by_role('cell', name='CONC. BANC. OUT/24')",
            'acao_linha': "page.get_by_role('row', name='09/10/2024 CONC. BANC. OUT/24 010010020001000007').get_by_role('cell', name='').locator('span')",
            'celula_valor': "page.get_by_role('cell', name='-R$ 4.356,85', exact=True).click()",

            # SELETORES FORM
            'campo_conciliacao': "page.locator('[name=\"aCols[colsCA7][nAt][8]\"]')",
            'campo_observacao': "page.locator('[name=\"aCols[colsCA7][nAt][11]\"]')",
            'campo_natureza': "page.locator('[name=\"aCols[colsCA7][nAt][1]\"]')",
            'menu_conta': "page.locator('[name=\"aCols[colsCA7][nAt][2]\"]')",
            'menu_item_conta': "page.locator('[name=\"aCols[colsCA7][nAt][3]\"]')",
            'menu_classe_valor': "page.locator('[name=\"aCols[colsCA7][nAt][4]\"]')",
            'campo_historico': "page.locator('[name=\"aCols[colsCA7][nAt][5]\"]')",
            'campo_data_conciliacao': "page.locator('[name=\"aCols[colsCA7][nAt][6]\"]')",
            'campo_valor': "page.locator('[name=\"aCols[colsCA7][nAt][7]\"]')",
            'campo_data_base': "page.locator('[name=\"aCols[colsCA7][nAt][9]\"]')",
            'campo_numeracao': "page.locator('[name=\"aCols[colsCA7][nAt][10]\"]')",
            
            # JANELA CONFIRMAÇÃO
            'confirmar_inclusao_linha': "page.get_by_role('button', name='Confirmar')",
            'cancelar_inclusao_linha': "page.get_by_role('button', name='Cancelar')",

            # BOTÕES PRINCIPAIS
            'btn_gravar': "page.get_by_role('button', name='Gravar')",
            'btn_incluir': "page.get_by_role('button', name='incluir')",

            # NAVEGAÇÃO ENTRE PÁGINAS
            'proxima_pagina': "page.locator('[title=\"Próxima página\"]')",
            'primeira_pagina': "page.locator('[title=\"Primeira página\"]')",
            'pagina_anterior': "page.locator('[title=\"Página anterior\"]')",
            'ultima_pagina': "page.locator('[title=\"Última página\"]')",

            # FILTROS
            'filtro_banco': "page.locator('[id=\"CA8_BANCO\"]')",
            'filtro_agencia': "page.locator('[id=\"CA8_AGENC\"]')",
            'filtro_conta': "page.locator('[id=\"CA8_NUMCON\"]')",
            'filtro_data_inicio': "page.locator('[id=\"CA8_DATCNC\"]')",
            'filtro_data_fim': "page.locator('[id=\"CA8_DATCN2\"]')",
            'aplicar_filtro': "page.get_by_role('button', name='Filtrar')",
            'limpar_filtro': "page.get_by_role('button', name='Limpar')",

            # OUTROS CONTROLES
            'fechar_janela': "page.get_by_role('button', name='Fechar')",
            'maximizar_janela': "page.get_by_role('button', name='Maximizar')",
        }
        
        logger.info("Seletores para conciliação manual definidos")
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
def extrair_dados_extrato(caminho_pdf):
    """
    Extrai dados do extrato PDF usando PyMuPDF
    Args:
        caminho_pdf (str): Caminho para o arquivo PDF
    
    Returns:
        dict: Dicionário com os dados extraídos ou None em caso de erro
    """
    
    
    try:                    
        doc = pymupdf.open(caminho_pdf)                                       
        pagina = doc.load_page(0)
        texto_pagina = pagina.get_text("text")                    
        
        padroes = {
            'saldo_inicial': r"SALDO INICIAL:\s*([\d\.,]+)",
            'saldo_inicial_tabela': r"SALDO INICIAL\.\.\.\.\.\.\.\.\.\s*([\d\.,]+)",
            'saldo_atual': r"SALDO ATUAL\s*([\d\.,]+)",
            'saldo_atual_tabela': r"SALDO ATUAL\s*([\d\.,]+)"
        }

        dados_extraidos = {}
        for chave, padrao in padroes.items():
            match = re.search(padrao, texto_pagina)
            if match:
                valor_str = match.group(1).replace('.', '').replace(',', '.')
                dados_extraidos[chave] = float(valor_str)
            else:
                dados_extraidos[chave] = None
        
        doc.close()
        return dados_extraidos

    except FileNotFoundError:
        logger.error(f"Arquivo não encontrado: {caminho_pdf}")
        return None
    except Exception as e:
        logger.error(f"Erro ao processar PDF: {e}")
        return None

