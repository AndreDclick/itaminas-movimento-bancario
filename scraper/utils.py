"""
Módulo de utilitários para automação com Playwright.
Contém funções auxiliares para interação com páginas web, manipulação de dados
e carregamento de parâmetros de configuração.
"""

from playwright.sync_api import Page
from config.logger import configure_logger
from .exceptions import (
    ExcecaoNaoMapeadaError,
    FormSubmitFailed
)
from config.settings import Settings
from datetime import datetime, date, timedelta
from pathlib import Path
import time
import os
import calendar
import json

# Configuração do logger para registro de atividades
logger = configure_logger()

class Utils:
    """Classe utilitária com métodos para auxiliar na automação de tarefas web."""
    
    def __init__(self, page: Page):
        """
        Inicializa a classe Utils com uma instância de página do Playwright.
        
        Args:
            page (Page): Instância da página do Playwright
        """
        self.page = page
        self._definir_locators()
    
    def _definir_locators(self):
        """
        Centraliza a definição de todos os locators usados na automação.
        Os locators são armazenados como variáveis de instância para reutilização.
        """
        self.locators = {
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"), 
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>")
        }
    
    def _fechar_popup_se_existir(self):
        """
        Tenta fechar popups que possam aparecer durante a execução.
        
        Este método verifica se um popup com botão "Fechar" está visível
        e tenta fechá-lo. Se não encontrar o popup, apenas registra um aviso.
        """
        try:
            time.sleep(5)  # Aguarda possível aparecimento do popup
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
                logger.info("Popup fechado")
        except Exception as e:
            logger.warning(f"Erro ao verificar popup: {e}")
    
    def _confirmar_operacao(self):
        """
        Confirma uma operação clicando no botão "Confirmar".
        
        Após a confirmação, verifica se há popups para fechar.
        
        Raises:
            FormSubmitFailed: Se não conseguir confirmar a operação
        """
        try:
            time.sleep(5) 
            self.locators['botao_confirmar'].click()
            logger.info("Operação confirmada")
        except Exception as e:
            error_msg = "Falha na confirmação da operação"
            logger.error(f"{error_msg}: {e}")
            raise FormSubmitFailed(error_msg) from e
    
    def obter_data_dia_anterior(self) -> str:
        """
        Calcula e retorna a data de ontem no formato DD/MM/YYYY.
        
        Returns:
            str: Data do dia anterior.
        """
        data_ontem = date.today() - timedelta(days=1)
        return data_ontem.strftime('%d/%m/%Y')

    def _resolver_valor(self, valor):
        """
        Resolve valores que contenham placeholders {{}} chamando funções correspondentes.
        
        Args:
            valor: Valor a ser resolvido (pode ser string com placeholder ou valor estático)
            
        Returns:
            Valor resolvido (pode ser string, tupla ou qualquer tipo retornado pela função)
        """
        # Verifica se o valor é uma string com placeholder
        if isinstance(valor, str) and valor.startswith('{{') and valor.endswith('}}'):
            placeholder = valor[2:-2].strip()  # Remove os {{ }}
            
            # Mapeamento de métodos disponíveis para resolução
            metodos_disponiveis = {
                "obter_data_dia_anterior": self.obter_data_dia_anterior
            }
            
            # Verifica se o método solicitado está disponível
            if placeholder in metodos_disponiveis:
                resultado = metodos_disponiveis[placeholder]()
                return resultado
            else:
                logger.warning(f"Método '{placeholder}' não encontrado para resolução")
                return valor  # Retorna o valor original se não encontrar o método
        else:
            return valor  # Retorna o valor original se não for um placeholder

    def _carregar_parametros(self, arquivo_json: str, chave: str):
        """
        Carrega parâmetros de configuração de um arquivo JSON.
        
        Args:
            arquivo_json (str): Nome do arquivo JSON com os parâmetros
            chave (str): Chave específica dentro do JSON a ser carregada
        """
        try:
            settings = Settings()
            caminho_arquivo = settings.PARAMETERS_DIR / arquivo_json
            
            with open(caminho_arquivo, 'r', encoding='utf-8') as file:
                dados = json.load(file)
            
            # Verifica se a chave existe no JSON
            if chave not in dados:
                raise KeyError(f"Chave '{chave}' não encontrada no arquivo {arquivo_json}")
            
            # Carrega os parâmetros e resolve placeholders
            self.parametros = {}
            for param, valor in dados[chave].items():
                self.parametros[param] = self._resolver_valor(valor)
            
            logger.info(f"Parâmetros carregados para chave '{chave}'")
            
        except FileNotFoundError as e:
            logger.error(f"Arquivo {arquivo_json} não encontrado: {e}")
            raise
        except KeyError as e:
            logger.error(f"Erro ao acessar chave no JSON: {e}")
            raise
        except json.JSONDecodeError as e:
            logger.error(f"Erro ao decodificar JSON: {e}")
            raise
        except Exception as e:
            error_msg = f"Erro inesperado ao carregar parâmetros: {e}"
            logger.error(error_msg)
            raise ExcecaoNaoMapeadaError(error_msg) from e

    def _validar_parametros(self, parametros_obrigatorios: list):
        """
        Valida se todos os parâmetros obrigatórios foram carregados corretamente.
        
        Args:
            parametros_obrigatorios (list): Lista de nomes de parâmetros obrigatórios
        """
        for param in parametros_obrigatorios:
            if param not in self.parametros:
                raise ValueError(f"Parâmetro obrigatório '{param}' não encontrado")
        
        logger.info("Todos os parâmetros obrigatórios validados com sucesso")