from playwright.sync_api import Page
from config.logger import configure_logger
from datetime import date
from pathlib import Path

import time
import os               
import calendar
import json

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

    def _resolver_valor(self, valor):
        """Resolve valores com placeholders {{}} chamando as funções correspondentes"""
        if isinstance(valor, str) and valor.startswith('{{') and valor.endswith('}}'):
            nome_metodo = valor[2:-2].strip()
            
            # Mapeamento de métodos disponíveis para todos os scrapers
            metodos_disponiveis = {
                'primeiro_e_ultimo_dia': self.primeiro_e_ultimo_dia,
                'obter_ultimo_dia_ano_passado': self.obter_ultimo_dia_ano_passado,
                'data_atual': self._get_data_atual
            }
            
            if nome_metodo in metodos_disponiveis:
                resultado = metodos_disponiveis[nome_metodo]()
                
                # Se for tupla (primeiro_e_ultimo_dia retorna (primeiro, ultimo))
                if isinstance(resultado, tuple):
                    # Lógica para decidir qual valor usar baseado no contexto
                    if 'inicial' in valor.lower() or 'primeiro' in valor.lower():
                        return resultado[0]  # Primeiro dia
                    elif 'final' in valor.lower() or 'ultimo' in valor.lower():
                        return resultado[1]  # Último dia
                    
                        # "data_inicial": "{{primeiro_e_ultimo_dia}}",
                        # "data_final": "{{primeiro_e_ultimo_dia}}",
                        # "data_sid_art": "{{obter_ultimo_dia_ano_passado}}",

                return resultado
            
            logger.warning(f"❌ Método '{nome_metodo}' não encontrado para resolver: {valor}")
        
        return valor

    def _get_data_atual(self):
        """Retorna a data atual formatada"""
        return date.today().strftime("%d/%m/%Y")
    
    
    def primeiro_e_ultimo_dia(self):
        hoje = date.today()
        mes_passado = hoje.month - 1 if hoje.month > 1 else 12
        ano_mes_passado = hoje.year if hoje.month > 1 else hoje.year - 1
        
        primeiro_dia = date(ano_mes_passado, mes_passado, 1).strftime("%d/%m/%Y")
        ultimo_dia_num = calendar.monthrange(ano_mes_passado, mes_passado)[1]
        ultimo_dia = date(ano_mes_passado, mes_passado, ultimo_dia_num).strftime("%d/%m/%Y")
        
        return primeiro_dia, ultimo_dia


    def obter_ultimo_dia_ano_passado(self):
        ano_passado = date.today().year - 1
        ultimo_dia = date(ano_passado, 12, 31).strftime("%d/%m/%Y")
        return ultimo_dia
    
    def _carregar_parametros(self, nome_arquivo: str, parametros_json: str = None) -> dict:
        """Carrega parâmetros de um arquivo JSON"""

        parametros_path = Path("parameters.json")
        try:
            if parametros_path.exists():
                with open(parametros_path, 'r', encoding='utf-8') as f:
                    parametros = json.load(f)
                    # Usa o parâmetro passado ou tenta usar o padrão
                    if parametros_json:
                        self.parametros = parametros.get(parametros_json, {})
                    else:
                        # Fallback: tenta usar o nome do arquivo sem extensão
                        chave_padrao = nome_arquivo.replace('.json', '')
                        self.parametros = parametros.get(chave_padrao, {})
                    
                    logger.info(f"Parâmetros carregados: {nome_arquivo} -> {parametros_json if parametros_json else chave_padrao}")
                    return self.parametros
            else:
                logger.error(f"❌ Arquivo de parâmetros não encontrado: {parametros_path}")
                return {}
        except json.JSONDecodeError as e:
            logger.error(f"❌ Erro ao decodificar JSON {nome_arquivo}: {e}")
            return {}
        except Exception as e:
            logger.error(f"❌ Erro ao carregar parâmetros {nome_arquivo}: {e}")
            return {}