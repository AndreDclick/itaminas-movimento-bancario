"""
Módulo de utilitários para automação com Playwright.
Contém funções auxiliares para interação com páginas web, manipulação de dados
e carregamento de parâmetros de configuração.
"""

from playwright.sync_api import Page
from config.logger import configure_logger
from .exceptions import (
    ExcecaoNaoMapeadaError,
    TimeoutOperacional,
    FormSubmitFailed
)
from datetime import date
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
            time.sleep(3)  # Aguarda possível aparecimento do popup
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
            time.sleep(3)  # Aguarda carregamento do botão
            self.locators['botao_confirmar'].click()
            logger.info("Operação confirmada")
            self._fechar_popup_se_existir()  # Fecha possíveis popups pós-confirmação
        except Exception as e:
            error_msg = "Falha na confirmação da operação"
            logger.error(f"{error_msg}: {e}")
            raise FormSubmitFailed(error_msg) from e
    
    def _selecionar_filiais(self):
        """
        Seleciona todas las filiais disponíveis usando o botão "Marca Todos".
        
        Este método é útil para processos que requerem seleção de múltiplas filiais.
        
        Raises:
            FormSubmitFailed: Se não conseguir selecionar as filiais
        """
        try: 
            time.sleep(3)  # Aguarda carregamento do botão
            if self.locators['botao_marcar_filiais'].is_visible():
                self.locators['botao_marcar_filiais'].click()
                time.sleep(1)  # Pequena pausa após seleção
                self.locators['botao_confirmar'].click()  # Confirma a seleção
                logger.info("Filial selecionada")
        except Exception as e:
            error_msg = "Falha na seleção de filiais"
            logger.error(f"{error_msg}: {e}")
            raise FormSubmitFailed(error_msg) from e
    
    def _resolver_valor(self, valor):
        """
        Resolve valores que contenham placeholders {{}} chamando funções correspondentes.
        
        Este método permite usar placeholders em configurações que serão substituídos
        por valores dinâmicos durante a execução.
        
        Args:
            valor: Valor a ser resolvido (pode ser string com placeholder ou valor estático)
            
        Returns:
            Valor resolvido (pode ser string, tupla ou qualquer tipo retornado pela função)
        """
        # Verifica se o valor é uma string com placeholder
        if isinstance(valor, str) and valor.startswith('{{') and valor.endswith('}}'):
            nome_metodo = valor[2:-2].strip()  # Extrai o nome do método do placeholder
            
            # Mapeamento de métodos disponíveis para resolução
            metodos_disponiveis = {
                'primeiro_e_ultimo_dia': self.primeiro_e_ultimo_dia,
                'obter_ultimo_dia_ano_passado': self.obter_ultimo_dia_ano_passado,
                'data_atual': self._get_data_atual
            }
            
            # Verifica se o método solicitado está disponível
            if nome_metodo in metodos_disponiveis:
                resultado = metodos_disponiveis[nome_metodo]()
                
                # Trata retornos em tupla (caso do primeiro_e_ultimo_dia)
                if isinstance(resultado, tuple):
                    # Lógica para decidir qual valor usar baseado no contexto do placeholder
                    if 'inicial' in valor.lower() or 'primeiro' in valor.lower():
                        return resultado[0]  # Retorna o primeiro dia
                    elif 'final' in valor.lower() or 'ultimo' in valor.lower():
                        return resultado[1]  # Retorna o último dia
                    else:
                        return resultado  # Retorna a tupla completa
                else:
                    return resultado  # Retorna o valor simples
            else:
                logger.warning(f"Método '{nome_metodo}' não encontrado para resolução")
                return valor  # Retorna o valor original se não encontrar o método
        else:
            return valor  # Retorna o valor original se não for um placeholder
    
    def _carregar_parametros(self, arquivo_json: str, chave: str):
        """
        Carrega parâmetros de configuração de um arquivo JSON.
        
        Este método lê um arquivo JSON e extrai os parâmetros para uma chave específica,
        resolvendo quaisquer placeholders encontrados nos valores.
        
        Args:
            arquivo_json (str): Nome do arquivo JSON com os parâmetros
            chave (str): Chave específica dentro do JSON a ser carregada
            
        Raises:
            FileNotFoundError: Se o arquivo JSON não for encontrado
            KeyError: Se a chave especificada não existir no JSON
            JSONDecodeError: Se o arquivo JSON estiver mal formatado
        """
        try:
            caminho_arquivo = Path(__file__).parent.parent / 'config' / arquivo_json
            
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
    
    def _get_data_atual(self):
        """
        Retorna a data atual no formato DD/MM/YYYY.
        
        Returns:
            str: Data atual formatada
        """
        return date.today().strftime('%d/%m/%Y')
    
    def primeiro_e_ultimo_dia(self):
        """
        Retorna uma tupla com o primeiro e último dia do mês atual.
        
        Returns:
            tuple: (primeiro_dia, ultimo_dia) no formato DD/MM/YYYY
        """
        hoje = date.today()
        primeiro_dia = date(hoje.year, hoje.month, 1)
        ultimo_dia = date(hoje.year, hoje.month, calendar.monthrange(hoje.year, hoje.month)[1])
        
        return (
            primeiro_dia.strftime('%d/%m/%Y'),
            ultimo_dia.strftime('%d/%m/%Y')
        )
    
    def obter_ultimo_dia_ano_passado(self):
        """
        Retorna o último dia do ano anterior.
        
        Returns:
            str: Último dia do ano anterior no formato DD/MM/YYYY
        """
        ano_passado = date.today().year - 1
        ultimo_dia = date(ano_passado, 12, 31)
        return ultimo_dia.strftime('%d/%m/%Y')
    
    def _validar_parametros(self, parametros_obrigatorios: list):
        """
        Valida se todos os parâmetros obrigatórios foram carregados corretamente.
        
        Args:
            parametros_obrigatorios (list): Lista de nomes de parâmetros obrigatórios
            
        Raises:
            ValueError: Se algum parâmetro obrigatório estiver faltando
        """
        for param in parametros_obrigatorios:
            if param not in self.parametros:
                raise ValueError(f"Parâmetro obrigatório '{param}' não encontrado")
        
        logger.info("Todos os parâmetros obrigatórios validados com sucesso")