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
    PLS_CONTAS_X_ITENS = os.getenv("PLANILHA_CONTAS_X_ITENS ")
    
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
        'fornecedor': 'Prf-NumeroParcela',  # Contém código do fornecedor + número da parcela
        'titulo': 'Tp',  # Tipo do título (de acordo com o fluxo do processo)
        'parcela': 'Parcela',  # Número da parcela (extraído do Prf-NumeroParcela)
        'tipo_titulo': 'Natureza',  # Natureza do título
        'data_emissao': 'Data deEmissao',
        'data_vencimento': 'Data deVencto',
        'valor_original': 'Valor Original',
        'saldo_devedor': 'Titulos a vencerValor nominal',  # Valor em aberto
        'situacao': 'Porta-do',  # Situação do título
        'conta_contabil': 'Historico(Vencidos+Vencer)',
        'centro_custo': 'Natureza'  # Ajustar conforme necessidade
    }
    COLUNAS_MODELO1 = {
        'conta_contabil': 'Codigo',         # Coluna B (segundo seu relatório)
        'descricao_conta': 'Descricao',     # Coluna C
        'saldo_anterior': 'Saldo anterior', # Coluna D
        'debito': 'Debito',                 # Coluna E
        'credito': 'Credito',               # Coluna F
        'saldo_atual': 'Saldo Atual'        # Coluna G (ajustado para o nome real)
    }
    COLUNAS_CONTAS_ITENS = {
        'conta_contabil': 'Descricao',       # Coluna B
        'item': 'Saldo anterior',            # Coluna C (ajustar conforme necessidade)
        'descricao_item': 'Debito',          # Coluna D (ajustar conforme necessidade)
        'quantidade': 'Credito',             # Coluna E (ajustar conforme necessidade)
        'valor_unitario': 'Mov periodo',     # Coluna F
        'valor_total': 'Saldo atual',        # Coluna G
        'saldo': 'Saldo atual'               # Coluna I (confirmar se é a mesma que G)
    }

    def __init__(self):
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)