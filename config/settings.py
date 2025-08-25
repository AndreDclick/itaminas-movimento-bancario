import os
import json
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
    
    #Planilhas
    CAMINHO_PLS = os.getenv("CAMINHO_PLANILHAS")
    PLS_FINANCEIRO = os.getenv("PLANILHA_FINANCEIRO")
    PLS_MODELO_1 = os.getenv("PLANILHA_MODELO_1")
    COLUNAS_CONTAS_ITENS = os.getenv("FORNECEDOR_NACIONAL")
    COLUNAS_ADIANTAMENTO = os.getenv("ADIANTAMENTO_NACIONAL")

    # Paths
    DATA_DIR = BASE_DIR / "data"
    LOGS_DIR = BASE_DIR / "logs"
    RESULTS_DIR = BASE_DIR / "results"
    DB_PATH = DATA_DIR / "database.db"
    UPLOAD_DIR = Path("./uploads/")
    PARAMETERS_DIR = "parameters"

    # Files
    DOWNLOAD_PATH = DATA_DIR 
    RESULTS_PATH = RESULTS_DIR 
    
    DATA_BASE = datetime.now().strftime("%d/%m/%Y")

    DATA_DIR.mkdir(exist_ok=True)
    RESULTS_DIR.mkdir(exist_ok=True)

    # Nomes das tabelas 
    TABLE_FINANCEIRO = "financeiro"
    TABLE_MODELO1 = "modelo1"
    TABLE_CONTAS_ITENS = "contas_itens"
    TABLE_ADIANTAMENTO = "adiantamento"
    TABLE_RESULTADO = "resultado"
    
    # Timeouts
    TIMEOUT = 30000  
    DELAY = 0.5  
    SHUTDOWN_DELAY = 3 
    
    # Browser
    HEADLESS = True
    
    # Planilhas
    FORNECEDORES_EXCLUIR = ['NDF', 'PA']  
    DATA_REFERENCIA = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%d/%m/%Y") 

    # Planilha Financeira (finr150.xlsx)
    COLUNAS_FINANCEIRO = {
        'fornecedor': 'Codigo-Nome do Fornecedor',
        'titulo': 'Prf-Numero Parcela',  
        'tipo_titulo': 'Tp',
        'data_emissao': 'Data de Emissao',
        'data_vencimento': 'Data de Vencto',
        "vencto_real": "VenctoReal",
        'valor_original': 'Valor Original',
        'saldo_devedor': 'Tit Vencidos Valor nominal',
        'situacao': 'Natureza',
        'conta_contabil': 'Natureza',
        'centro_custo': 'Porta- dor'
    }

    # Planilha Modelo 1 (ctbr040.xlsx)
    COLUNAS_MODELO1 = {
        'conta_contabil': 'Conta',
        'descricao_conta': 'Descricao',
        'saldo_anterior': 'Saldo anterior',
        'debito': 'Debito',
        'credito': 'Credito',
        'movimento_periodo': 'Mov  periodo',
        'saldo_atual': 'Saldo atual'
    }

    # Planilha Fornecedor Nacional (ctbr140.txt)
    COLUNAS_CONTAS_ITENS = {
        'conta_contabil': 'Codigo',
        'descricao_item': 'Descricao',
        'codigo_fornecedor': 'Codigo.1',
        'descricao_fornecedor': 'Descricao.1',
        'saldo_anterior': 'Saldo anterior',
        'debito': 'Debito',
        'credito': 'Credito',
        'movimento_periodo': 'Movimento do periodo',
        'saldo_atual': 'Saldo atual'
    }

    # Planilha Adiantamento Nacional (ctbr100.txt)
    COLUNAS_ADIANTAMENTO = {
        'conta_contabil': 'Codigo',
        'descricao_item': 'Descricao',
        'codigo_fornecedor': 'Codigo.1',
        'descricao_fornecedor': 'Descricao.1',
        'saldo_anterior': 'Saldo anterior',
        'debito': 'Debito',
        'credito': 'Credito',
        'movimento_periodo': 'Movimento do periodo',
        'saldo_atual': 'Saldo atual'
    }

    def __init__(self):
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)