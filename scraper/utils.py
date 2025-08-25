"""
Módulo de utilitários para automação com Playwright.
Contém funções auxiliares para interação com páginas web, manipulação de dados
e carregamento de parâmetros de configuração.
"""

from playwright.sync_api import Page
from config.logger import configure_logger
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
        except Exception as e:
            logger.warning(f" Erro ao verificar popup: {e}")
    
    def _confirmar_operacao(self):
        """
        Confirma uma operação clicando no botão "Confirmar".
        
        Após a confirmação, verifica se há popups para fechar.
        
        Raises:
            Exception: Se não conseguir confirmar a operação
        """
        try:
            time.sleep(3)  # Aguarda carregamento do botão
            self.locators['botao_confirmar'].click()
            logger.info("operação confirmada")
            self._fechar_popup_se_existir()  # Fecha possíveis popups pós-confirmação
        except Exception as e:
            logger.error(f"Falha na confirmação: {e}")
            raise  # Propaga a exceção para tratamento superior
    
    def _selecionar_filiais(self):
        """
        Seleciona todas as filiais disponíveis usando o botão "Marca Todos".
        
        Este método é útil para processos que requerem seleção de múltiplas filiais.
        
        Raises:
            Exception: Se não conseguir selecionar as filiais
        """
        try: 
            time.sleep(3)  # Aguarda carregamento do botão
            if self.locators['botao_marcar_filiais'].is_visible():
                self.locators['botao_marcar_filiais'].click()
                time.sleep(1)  # Pequena pausa após seleção
                self.locators['botao_confirmar'].click()  # Confirma a seleção
        except Exception as e:
            logger.error(f"Falha na escolha de filiais {e}")
            raise  # Propaga a exceção para tratamento superior
    
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
                
                return resultado  # Retorna o resultado para métodos que não retornam tupla
            
            logger.warning(f"❌ Método '{nome_metodo}' não encontrado para resolver: {valor}")
        
        # "data_inicial": "{{primeiro_e_ultimo_dia}}",
                        # "data_final": "{{primeiro_e_ultimo_dia}}",
                        # "data_sid_art": "{{obter_ultimo_dia_ano_passado}}",
        
        return valor  # Retorna o valor original se não for um placeholder
    
    def _get_data_atual(self):
        """
        Retorna a data atual formatada no padrão brasileiro.
        
        Returns:
            str: Data atual no formato DD/MM/AAAA
        """
        return date.today().strftime("%d/%m/%Y")
    
    def primeiro_e_ultimo_dia(self):
        """
        Calcula o primeiro e último dia do mês anterior ao atual.
        
        Returns:
            tuple: (primeiro_dia, ultimo_dia) ambos no formato DD/MM/AAAA
        """
        hoje = date.today()
        
        # Calcula mês e ano do mês anterior
        mes_passado = hoje.month - 1 if hoje.month > 1 else 12
        ano_mes_passado = hoje.year if hoje.month > 1 else hoje.year - 1
        
        # Calcula primeiro e último dia do mês anterior
        primeiro_dia = date(ano_mes_passado, mes_passado, 1).strftime("%d/%m/%Y")
        ultimo_dia_num = calendar.monthrange(ano_mes_passado, mes_passado)[1]
        ultimo_dia = date(ano_mes_passado, mes_passado, ultimo_dia_num).strftime("%d/%m/%Y")
        
        return primeiro_dia, ultimo_dia
    
    def obter_ultimo_dia_ano_passado(self):
        """
        Retorna o último dia do ano anterior ao atual.
        
        Returns:
            str: Último dia do ano anterior no formato DD/MM/AAAA
        """
        ano_passado = date.today().year - 1
        ultimo_dia = date(ano_passado, 12, 31).strftime("%d/%m/%Y")
        return ultimo_dia
    
    def _carregar_parametros(self, nome_arquivo: str, parametros_json: str = None) -> dict:
        """
        Carrega parâmetros de um arquivo JSON de configuração.
        
        Args:
            nome_arquivo (str): Nome do arquivo (apenas para log)
            parametros_json (str, optional): Chave específica no JSON. Se None, 
                                            tenta usar o nome do arquivo sem extensão.
        
        Returns:
            dict: Dicionário com parâmetros carregados ou dicionário vazio em caso de erro
        """
        parametros_path = Path("parameters.json")
        
        try:
            if parametros_path.exists():
                with open(parametros_path, 'r', encoding='utf-8') as f:
                    parametros = json.load(f)
                    
                    # Tenta encontrar a chave específica ou usa fallback
                    if parametros_json:
                        self.parametros = parametros.get(parametros_json, {})
                    else:
                        # Fallback: usa o nome do arquivo sem extensão como chave
                        chave_padrao = nome_arquivo.replace('.json', '')
                        self.parametros = parametros.get(chave_padrao, {})
                    
                    logger.info(f"Parâmetros carregados: {nome_arquivo} -> {parametros_json if parametros_json else chave_padrao}")
                    return self.parametros
            else:
                logger.error(f"❌ Arquivo de parâmetros não encontrado: {parametros_path}")
                return {}  # Retorna dicionário vazio se arquivo não existir
                
        except json.JSONDecodeError as e:
            logger.error(f"❌ Erro ao decodificar JSON {nome_arquivo}: {e}")
            return {}  # Retorna dicionário vazio em caso de erro de decodificação
        except Exception as e:
            logger.error(f"❌ Erro ao carregar parâmetros {nome_arquivo}: {e}")
            return {}  # Retorna dicionário vazio para qualquer outro erro