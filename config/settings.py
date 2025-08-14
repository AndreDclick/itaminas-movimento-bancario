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
    PLS_MODELO_1 = os.getenv("PLANILHA_MODELO_1")
    PLS_CONTAS_X_ITENS = os.getenv("PLANILHA_CONTAS_X_ITENS")
    
    # Paths
    DATA_DIR = BASE_DIR / "data"
    LOGS_DIR = BASE_DIR / "logs"
    RESULTS_DIR = BASE_DIR / "results"
    DB_PATH = BASE_DIR / "database.db"
    
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
    TABLE_RESULTADO = "resultado"
    
    # Mapeamento de colunas (exemplo)
    COLUNAS_FINANCEIRO = {
        "fornecedor": "Fornecedor",
        "titulo": "Título",
        "saldo_devedor": "Saldo Devedor"
    }
    
    DATA_REFERENCIA = "01/01/2023"
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
        'fornecedor': 'Codigo-Nome do Fornecedor',
        'titulo': 'Prf-Numero_x000D_<br>Parcela',
        'tipo_titulo': 'Tp',
        'data_emissao': 'Data de_x000D_<br>Emissao',
        'data_vencimento': 'Data de_x000D_<br>Vencto',
        'valor_original': 'Valor Original',
        'saldo_devedor': 'Tit Vencidos_x000D_<br>Valor corrigido',
        'situacao': 'Porta-_x000D_<br>dor',
        'conta_contabil': 'Historico(Vencidos+Vencer)',
        'centro_custo': 'Natureza'
    }
    COLUNAS_MODELO1 = {
        'descricao_conta': 'Descricao',           # Coluna B (descrição do item)
        'codigo_fornecedor': 'Codigo',            # Coluna C (código do fornecedor)
        'descricao_fornecedor': 'Descricao',      # Coluna D (descrição do fornecedor)
        'saldo_anterior': 'Saldo anterior',       # Coluna E
        'debito': 'Debito',                       # Coluna F
        'credito': 'Credito',                     # Coluna G
        'movimento_periodo': 'Movimento do periodo', # Coluna H
        'saldo_atual': 'Saldo atual'              # Coluna I
    }
    COLUNAS_CONTAS_ITENS = {
        'conta_contabil': 'Conta',               # Coluna A
        'descricao_item': 'Descricao',           # Coluna B
        'saldo_anterior': 'Saldo anterior',      # Coluna C
        'debito': 'Debito',                      # Coluna D
        'credito': 'Credito',                    # Coluna E
        'movimento_periodo': 'Mov  periodo',     # Coluna F
        'saldo_atual': 'Saldo atual'             # Coluna G
    }

    def __init__(self):
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)