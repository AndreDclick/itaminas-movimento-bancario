import os
from pathlib import Path
from datetime import datetime, timedelta  
from dotenv import load_dotenv

class Settings:
    BASE_DIR = Path(__file__).resolve().parent.parent

    # Carrega o .env
    load_dotenv(BASE_DIR / ".env")

    # Dados sensíveis
    USUARIO = os.getenv("USUARIO")
    SENHA = os.getenv("SENHA")
    BASE_URL = os.getenv("BASE_URL")
    CAMINHO_PLS = os.getenv("CAMINHO_PLANILHAS")
    PLS_FINANCEIRO = os.getenv("PLANILHA_FINANCEIRO")
    PLS_MODELO = os.getenv("PLANILHA_MODELO_1")
    PLS_CONTAS = os.getenv("PLANILHA_CONTAS")
    
    # Paths
    DATA_DIR = BASE_DIR / "data"
    LOGS_DIR = BASE_DIR / "logs"
    RESULTS_DIR = BASE_DIR / "results"
    
    # Files
    DOWNLOAD_PATH = DATA_DIR 
    RESULTS_PATH = RESULTS_DIR 
    
    DATA_BASE = datetime.now().strftime("%d/%m/%Y")

    # Timeouts
    TIMEOUT = 15000  
    DELAY = 0.5  
    SHUTDOWN_DELAY = 3 
    
    # Browser
    HEADLESS = False
    
    # Planilhas
    FORNECEDORES_EXCLUIR = ['NDF', 'PA']  # Siglas de fornecedores a excluir
    DATA_REFERENCIA = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%d/%m/%Y")  # Último dia do mês anterior
    COLUNAS_FINANCEIRO = {
            'fornecedor': 'B',
            'titulo': 'C',
            'parcela': 'D',
            'tipo_titulo': 'E',
            'data_emissao': 'F',
            'data_vencimento': 'G',
            'valor_original': 'K',
            'saldo_devedor': 'L',
            'situacao': 'M',
            'conta_contabil': 'N',
            'centro_custo': 'O'
        }
    COLUNAS_MODELO1 = {
            'conta_contabil': 'B',
            'descricao_conta': 'C',
            'saldo_anterior': 'D',
            'debito': 'E',
            'credito': 'F',
            'saldo_atual': 'G'
        }
    COLUNAS_CONTAS_ITENS = {
            'conta_contabil': 'B',
            'item': 'C',
            'descricao_item': 'D',
            'quantidade': 'E',
            'valor_unitario': 'F',
            'valor_total': 'G',
            'saldo': 'I'
        }

    def __init__(self):
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)