"""
Configurações do Sistema de Movimentação Bancária
Arquivo: settings.py
Descrição: Configurações globais, constantes e parâmetros do sistema
Desenvolvido por: DCLICK
"""

import os
import json
import sys
from pathlib import Path
from datetime import datetime, timedelta  
from dotenv import load_dotenv


def setup_environment():
    """
    Configura o ambiente para carregar o .env corretamente tanto no 
    desenvolvimento quanto no executável PyInstaller
    """
    if getattr(sys, 'frozen', False):
        # Se está rodando como executável PyInstaller
        # O .env está no mesmo diretório do executável, não no _MEIPASS
        base_path = Path(sys.executable).parent
        env_path = base_path / '.env'
        print(f" Modo: Executável PyInstaller")
        print(f" Diretório do executável: {base_path}")
    else:
        # Se está rodando como script
        base_path = Path(__file__).resolve().parent.parent
        env_path = base_path / '.env'
        print(f" Modo: Desenvolvimento")
        print(f" Diretório do projeto: {base_path}")
    
    # Carregar .env
    if env_path.exists():
        load_dotenv(env_path)
        print(f"✅ .env carregado de: {env_path}")
        
        # Verificar se as variáveis foram carregadas
        test_vars = ['USUARIO', 'BASE_URL']
        for var in test_vars:
            value = os.getenv(var)
            print(f"   {var}: {'✅' if value else '❌'} {'***' if var == 'USUARIO' and value else value}")
        
        return True
    else:
        print(f"❌ .env NÃO encontrado em: {env_path}")
        print(f" Conteúdo do diretório:")
        try:
            for item in base_path.iterdir():
                print(f"   - {item.name}")
        except Exception as e:
            print(f"   Erro ao listar diretório: {e}")
        return False

# Executar a configuração do ambiente
env_loaded = setup_environment()

class Settings:
    """
    Classe principal de configurações do sistema.
    Centraliza todas as constantes, paths e parâmetros de configuração.
    """
    
    # =========================================================================
    # CONFIGURAÇÕES DE DIRETÓRIOS E PATHS BASE
    # =========================================================================
    
    if getattr(sys, 'frozen', False):
        BASE_DIR = Path(sys.executable).parent  # Diretório do executável
    else:
        BASE_DIR = Path(__file__).resolve().parent.parent  # Diretório do projeto

    # Carrega variáveis de ambiente do arquivo .env
    load_dotenv(BASE_DIR / ".env")

    # =========================================================================
    # DADOS SENSÍVEIS (carregados de variáveis de ambiente)
    # =========================================================================
    
    USUARIO = os.getenv("USUARIO")              # Usuário do sistema Protheus
    SENHA = os.getenv("SENHA")                  # Senha do sistema Protheus
    BASE_URL = os.getenv("BASE_URL")            # URL base do sistema Protheus
    WEB_AGENT_PATH = (r"C:\Users\rpa.dclick\Desktop\PROTHEUS DEV.lnk")
    
    # =========================================================================
    # DIRETÓRIOS DO SISTEMA
    # =========================================================================
    
    DATA_DIR = BASE_DIR / "data"          # Diretório para armazenamento de dados
    LOGS_DIR = BASE_DIR / "logs"          # Diretório para arquivos de log
    RESULTS_DIR = BASE_DIR / "results"    # Diretório para resultados e relatórios
    PARAMETERS_DIR = BASE_DIR / "parameters.json"         # Diretório para parâmetros do sistema

    # Paths para download e resultados
    DOWNLOADS_DIR = BASE_DIR / "downloads"
    DOWNLOADS_DIR.mkdir(exist_ok=True)
    RESULTS_PATH = RESULTS_DIR 
    
    # Data base para processamento (formato: DD/MM/AAAA)
    DATA_BASE = datetime.now().strftime("%d/%m/%Y")

    # Criar diretórios se não existirem
    DATA_DIR.mkdir(exist_ok=True)
    RESULTS_DIR.mkdir(exist_ok=True)

    
    TIMEOUT = 60000      # Timeout para operações (30 segundos)
    DELAY = 0.5          # Delay entre operações (0.5 segundos)
    SHUTDOWN_DELAY = 3   # Delay para desligamento (3 segundos)
    
    # =========================================================================
    # CONFIGURAÇÕES DO NAVEGADOR (BROWSER)
    # =========================================================================
    
    HEADLESS = True  # Executar navegador em modo headless (sem interface)
    
    # =========================================================================
    # CONFIGURAÇÕES DE EMAIL
    # =========================================================================
    
    # Lista de destinatários por tipo de email
    EMAILS = {
        "success": ["andre.rodrigues@dclick.com.br", "talles.salmon@itaminas.com.br", "rodrigo.couto@itaminas.com.br", "ellen.leao@itaminas.com.br"],  # Destinatários para emails de sucesso
        "error": ["andre.rodrigues@dclick.com.br", "talles.salmon@itaminas.com.br", "rodrigo.couto@itaminas.com.br", "ellen.leao@itaminas.com.br"]     # Destinatários para emails de erro
    }

    PASSWORD = os.getenv("PASSWORD") 
    # Configurações SMTP para envio de emails
    SMTP = {
        "enabled": True,                       # Habilitar/desabilitar envio de emails
        "host": "smtp.gmail.com",           # Servidor SMTP
        "port": 587,                            # Porta do servidor SMTP
        "from": "suporte@dclick.com.br",                           # Remetente dos emails
        "password": PASSWORD,                    # Senha do email remetente
        "template": "templates/email_conciliação.html",  # Template HTML para emails
        "logo": "https://www.dclick.com.br/themes/views/web/assets/logo.svg"            # Logo para incorporar nos emails
    }

    # =========================================================================
    # CONFIGURAÇÕES DE PLANILHAS E PROCESSAMENTO
    # =========================================================================
    
    # Fornecedores a serem excluídos do processamento
    # FORNECEDORES_EXCLUIR = ['NDF', 'PA']  
    
    # Data de referência para processamento (último dia do mês anterior)
    DATA_REFERENCIA = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%d/%m/%Y") 

    def __init__(self):
        """
        Inicializador da classe Settings.
        Garante que todos os diretórios necessários existam.
        """
        # Criar diretórios se não existirem
        # os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)
        
        # Validar variáveis críticas (mas não falhar imediatamente)
        self._validate_required_vars()




    def _validate_required_vars(self):
        """Valida se as variáveis obrigatórias estão presentes e corretas."""
        required_vars = {
            'USUARIO': self.USUARIO,
            'SENHA': self.SENHA, 
            'BASE_URL': self.BASE_URL,
            
        }
        
        missing_vars = []
        for var_name, var_value in required_vars.items():
            if not var_value:
                missing_vars.append(var_name)
        
        if missing_vars:
            error_msg = f"Variáveis de ambiente obrigatórias não carregadas: {', '.join(missing_vars)}"
            print(f"❌ {error_msg}")
            # Não levanta exceção imediatamente, apenas registra o erro
            # raise ValueError(error_msg)

# Instância global para importação
try:
    settings = Settings()
except Exception as e:
    print(f"❌ Erro crítico ao inicializar Settings: {e}")
    # Cria uma instância básica para evitar falha completa
    settings = None