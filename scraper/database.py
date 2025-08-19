import sqlite3
from pathlib import Path
from config.settings import Settings
from config.logger import configure_logger
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from difflib import get_close_matches
from workalendar.america import Brazil
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
import openpyxl
import re

# Configura o logger para registrar eventos
logger = configure_logger()

class DatabaseManager:
    # Implementação do padrão Singleton para garantir apenas uma instância
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DatabaseManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        # Verifica se já foi inicializado (para evitar múltiplas inicializações no Singleton)
        if self._initialized:
            return
        self.settings = Settings()  # Carrega configurações
        self.conn = None  # Conexão com o banco
        self.logger = configure_logger()  # Logger específico da classe
        self._initialize_database()  # Inicializa o banco de dados
        self._initialized = True  # Marca como inicializado

    def _initialize_database(self):
        """Inicializa o banco de dados SQLite e cria as tabelas necessárias"""
        try:
            # Conecta ao banco SQLite
            self.conn = sqlite3.connect(self.settings.DB_PATH, timeout=10)
            cursor = self.conn.cursor()
            
            # Cria tabela financeiro se não existir
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {self.settings.TABLE_FINANCEIRO} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fornecedor TEXT,
                    titulo TEXT,
                    parcela TEXT,
                    tipo_titulo TEXT,
                    data_emissao TEXT,
                    data_vencimento TEXT,
                    valor_original REAL DEFAULT 0,
                    saldo_devedor REAL DEFAULT 0,
                    situacao TEXT,
                    conta_contabil TEXT,
                    centro_custo TEXT,
                    excluido BOOLEAN DEFAULT 0,
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Cria tabela modelo1 (contábil) se não existir
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {self.settings.TABLE_MODELO1} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conta_contabil TEXT,
                    descricao_conta TEXT,
                    codigo_fornecedor TEXT,
                    descricao_fornecedor TEXT,
                    saldo_anterior REAL DEFAULT 0,
                    debito REAL DEFAULT 0,
                    credito REAL DEFAULT 0,
                    saldo_atual REAL DEFAULT 0,
                    tipo_fornecedor TEXT,
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Cria tabela contas_itens se não existir
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS contas_itens (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conta_contabil TEXT,
                    descricao_item TEXT,
                    saldo_anterior REAL DEFAULT 0,
                    debito REAL DEFAULT 0,
                    credito REAL DEFAULT 0,
                    saldo_atual REAL DEFAULT 0,
                    item TEXT DEFAULT '',
                    quantidade REAL DEFAULT 1,
                    valor_unitario REAL DEFAULT 0,
                    valor_total REAL DEFAULT 0,
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Cria tabela resultado (concatenação) se não existir
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {self.settings.TABLE_RESULTADO} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo_fornecedor TEXT,
                    descricao_fornecedor TEXT,
                    saldo_contabil REAL DEFAULT 0,
                    saldo_financeiro REAL DEFAULT 0,
                    diferenca REAL DEFAULT 0,
                    status TEXT CHECK(status IN ('OK', 'DIVERGENTE', 'PENDENTE')),
                    detalhes TEXT,
                    ordem_importancia INTEGER,
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Função auxiliar para garantir que colunas existam nas tabelas
            def ensure_column(table, column, type_):
                cursor.execute(f"PRAGMA table_info({table})")
                cols = [c[1] for c in cursor.fetchall()]
                if column not in cols:
                    cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {type_}")
            
            # Garante que colunas importantes existam
            ensure_column(self.settings.TABLE_MODELO1, 'codigo_fornecedor', 'TEXT')
            ensure_column(self.settings.TABLE_MODELO1, 'descricao_fornecedor', 'TEXT')
            ensure_column(self.settings.TABLE_RESULTADO, 'ordem_importancia', 'INTEGER')
            
            self.conn.commit()  # Confirma as alterações
            logger.info("Banco de dados inicializado com sucesso")
            
        except Exception as e:
            logger.error(f"Erro ao inicializar banco de dados: {e}")
            raise

    def aplicar_sugestoes_colunas(self, df, missing_mappings):
        """Aplica sugestões automáticas para mapeamento de colunas faltantes"""
        candidates = df.columns.tolist()
        lower_map = {c.lower(): c for c in candidates}  # Mapa case-insensitive

        # Mapeamento manual pré-definido para colunas comuns
        manual_mapping = {
            'Codigo-Nome do Fornecedor': 'fornecedor',
            'Prf-Numero Parcela': 'titulo',
            'Tp': 'tipo_titulo',
            'Data de Emissao': 'data_emissao',
            'Data de Vencto': 'data_vencimento',
            'Valor Original': 'valor_original',
            'Tit Vencidos Valor nominal': 'saldo_devedor',
            'Natureza': 'situacao',
            'Porta- dor': 'centro_custo'
        }

        # Aplica mapeamento manual primeiro
        for src, dest in manual_mapping.items():
            if src in df.columns and dest in missing_mappings:
                df.rename(columns={src: dest}, inplace=True)
                logger.warning(f"Mapeamento manual aplicado: '{src}' → '{dest}'")
                missing_mappings.remove(dest)

        # Tenta encontrar correspondências automáticas para colunas restantes
        for db_col in list(missing_mappings):
            # Busca por correspondência case-insensitive
            if db_col.lower() in lower_map:
                match = lower_map[db_col.lower()]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}' (case-insensitive match)")
                df.rename(columns={match: db_col}, inplace=True)
                missing_mappings.remove(db_col)
                continue

            # Busca por correspondência fuzzy (similaridade)
            similar = get_close_matches(db_col, candidates, n=1, cutoff=0.6)
            if similar:
                match = similar[0]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}'")
                df.rename(columns={match: db_col}, inplace=True)
                missing_mappings.remove(db_col)
                continue

            # Busca por correspondência fuzzy case-insensitive
            similar_lower = get_close_matches(db_col.lower(), list(lower_map.keys()), n=1, cutoff=0.6)
            if similar_lower:
                match = lower_map[similar_lower[0]]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}' (fuzzy case-insensitive)")
                df.rename(columns={match: db_col}, inplace=True)
                missing_mappings.remove(db_col)

        return df

    def import_from_excel(self, file_path, table_name):
        """Importa dados de arquivo Excel para a tabela especificada"""
        try:
            # Lê o arquivo Excel (pula a primeira linha como header)
            df = pd.read_excel(file_path, header=1)  
            logger.info(f"Colunas originais em {file_path}: {df.columns.tolist()}")
            # Limpa caracteres especiais dos nomes das colunas
            df.columns = df.columns.str.replace(r'_x000D_\n', ' ', regex=True).str.strip()
            logger.info(f"Colunas após limpeza: {df.columns.tolist()}")
            
            # Obtém mapeamento de colunas baseado no nome do arquivo
            column_mapping = self._get_column_mapping(Path(file_path))
            df.rename(columns=column_mapping, inplace=True)
            logger.info(f"Colunas após mapeamento: {df.columns.tolist()}")
            
            # Verifica colunas obrigatórias
            expected_columns = self.get_expected_columns(table_name)
            missing_mappings = [col for col in expected_columns if col not in df.columns]

            if missing_mappings:
                logger.warning(f"Colunas mapeadas não encontradas: {missing_mappings}")
                # Tenta aplicar sugestões automáticas
                df = self.aplicar_sugestoes_colunas(df, missing_mappings)
                remaining_missing = [col for col in expected_columns if col not in df.columns]
                
                # Trata colunas específicas que podem ser derivadas
                if remaining_missing:
                    if 'parcela' in remaining_missing and 'titulo' in df.columns:
                        # Extrai número da parcela do título
                        df['parcela'] = df['titulo'].astype(str).str.extract(r'(\d+)$').fillna('1')
                        logger.warning("Coluna 'parcela' criada a partir do título")
                        remaining_missing.remove('parcela')
                    
                    if 'conta_contabil' in remaining_missing:
                        # Define valor padrão para conta contábil
                        df['conta_contabil'] = 'CONTA_PADRAO' 
                        logger.warning("Coluna 'conta_contabil' preenchida com valor padrão")
                        remaining_missing.remove('conta_contabil')
                    
                    if remaining_missing:
                        logger.error(f"Colunas obrigatórias ausentes após tratamento: {remaining_missing}")
                        raise ValueError(f"Colunas obrigatórias ausentes: {remaining_missing}")

            # Limpa e prepara os dados
            df = self._clean_dataframe(df, table_name.lower())
            # Obtém colunas da tabela destino
            table_columns = [col[1] for col in self.conn.execute(f"PRAGMA table_info({table_name})").fetchall()]
            # Mantém apenas colunas que existem na tabela
            keep = [col for col in df.columns if col in table_columns]
            df = df[keep]
            # Insere dados no banco
            df.to_sql(table_name, self.conn, if_exists='append', index=False)
            logger.info(f"Dados importados para '{table_name}' com sucesso.")
            return True  
        except Exception as e:
            logger.error(f"Falha ao importar {file_path}: {e}", exc_info=True)
            return False
    
    def get_expected_columns(self, table_name):
        """Retorna lista de colunas esperadas para cada tipo de tabela"""
        if table_name == self.settings.TABLE_FINANCEIRO:
            return [
                'fornecedor', 'titulo', 'parcela', 'tipo_titulo',
                'data_emissao', 'data_vencimento', 'valor_original',
                'saldo_devedor', 'situacao', 'conta_contabil', 'centro_custo'
            ]
        elif table_name == self.settings.TABLE_MODELO1:
            return [
                'conta_contabil', 'descricao_conta', 'codigo_fornecedor', 'descricao_fornecedor',
                'saldo_anterior', 'debito', 'credito', 'saldo_atual'
            ]
        elif table_name == self.settings.TABLE_CONTAS_ITENS:
            return [
                'conta_contabil',
                'descricao_item',
                'saldo_anterior',
                'debito',
                'credito',
                'saldo_atual'
            ]
        else:
            raise ValueError(f"Tabela desconhecida: {table_name}")
        
    def _clean_dataframe(self, df, sheet_type):
        """Executa limpeza geral do DataFrame baseado no tipo de planilha"""
        try:
            # Limpa strings e remove valores vazios
            df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
            df = df.replace(['nan', 'None', ''], np.nan)
            df = df.dropna(how='all')  # Remove linhas completamente vazias
            
            # Aplica limpeza específica por tipo de planilha
            if sheet_type == 'financeiro':
                df = self._clean_financeiro_data(df)
            elif sheet_type == 'modelo1':
                df = self._clean_modelo1_data(df)
            elif sheet_type == 'contas_itens':
                df = self._clean_contas_itens_data(df)
            
            df = df.drop_duplicates()  # Remove duplicatas
            logger.info(f"DataFrame limpo - shape final: {df.shape}")
            return df
        except Exception as e:
            logger.error(f"Erro na limpeza dos dados ({sheet_type}): {str(e)}", exc_info=True)
            raise

    def _clean_financeiro_data(self, df):
        """Limpeza específica para dados financeiros"""
        # Remove registros de fornecedores NDF/PA
        if 'fornecedor' in df.columns:
            df = df[~df['fornecedor'].str.contains(r'\bNDF\b|\bPA\b', case=False, na=False)]
        
        # Garante que todas as colunas obrigatórias existam
        required_cols = ['fornecedor', 'titulo', 'parcela', 'tipo_titulo', 
                        'data_emissao', 'data_vencimento', 'valor_original',
                        'saldo_devedor', 'situacao', 'conta_contabil', 'centro_custo']
        
        for col in required_cols:
            if col not in df.columns:
                df[col] = np.nan
        
        # Limpa e converte colunas numéricas
        num_cols = ['valor_original', 'saldo_devedor']
        for col in num_cols:
            if col in df.columns:
                df[col] = (df[col].astype(str)
                            .str.replace(r'[^\d,-]', '', regex=True)  # Remove caracteres não numéricos
                            .str.replace(',', '.')  # Padroniza decimal
                            .replace('', '0')
                            .astype(float))
        
        # Cria coluna de comparação para validação
        if 'valor_original' in df.columns and 'saldo_devedor' in df.columns:
            df['comparacao'] = df['valor_original'] - df['saldo_devedor']
        
        return df

    def _clean_modelo1_data(self, df):
        """Limpeza específica para dados do modelo1 (contábil)"""
        # Classifica tipo de fornecedor baseado na descrição
        if 'descricao_conta' in df.columns:
            df['tipo_fornecedor'] = df['descricao_conta'].apply(
                lambda x: 'FORNECEDOR NACIONAL' if 'FORNEC' in str(x).upper() and 'NAC' in str(x).upper()
                else 'FORNECEDOR' if 'FORNEC' in str(x).upper()
                else 'OUTROS'
            )
        
        # Garante colunas de fornecedor
        if 'codigo_fornecedor' not in df.columns:
            df['codigo_fornecedor'] = None
        if 'descricao_fornecedor' not in df.columns:
            df['descricao_fornecedor'] = None
        
        # Tenta extrair código do fornecedor da descrição
        if df['codigo_fornecedor'].isna().all() and 'descricao_conta' in df.columns:
            df['codigo_fornecedor'] = df['descricao_conta'].str.extract(r'(\d{3,})', expand=False)
        
        # Limpa e converte colunas numéricas
        num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str)
                    .str.replace(r'[^\d,-]', '', regex=True)
                    .str.replace(',', '.'),
                    errors='coerce'
                ).fillna(0)
        
        return df

    def _clean_contas_itens_data(self, df):
        """Limpeza específica para dados de contas x itens"""
        # Limpa e converte colunas numéricas
        num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str)
                    .str.replace(r'[^\d,-]', '', regex=True)
                    .str.replace(',', '.'),
                    errors='coerce'
                ).fillna(0)
        
        # Preenche colunas derivadas se não existirem
        if 'item' not in df.columns:
            df['item'] = df.get('descricao_item', '').str[:50]  # Limita a 50 caracteres
            
        if 'quantidade' not in df.columns:
            df['quantidade'] = 1  # Valor padrão
            
        if 'valor_unitario' not in df.columns and 'saldo_atual' in df.columns:
            df['valor_unitario'] = df['saldo_atual']  # Usa saldo como valor unitário
            
        if 'valor_total' not in df.columns and 'saldo_atual' in df.columns:
            df['valor_total'] = df['saldo_atual']  # Usa saldo como valor total

        return df
    
    def _get_column_mapping(self, file_path: Path):
        """Retorna mapeamento de colunas baseado no nome do arquivo"""
        filename = file_path.stem.lower()
        
        # Mapeamento para arquivos financeiros
        if 'finr' in filename:
            return {
                'Codigo-Nome do Fornecedor': 'fornecedor',
                'Prf-Numero Parcela': 'titulo',
                'Tp': 'tipo_titulo',
                'Data de Emissao': 'data_emissao',
                'Data de Vencto': 'data_vencimento',
                'Valor Original': 'valor_original',
                'Tit Vencidos Valor nominal': 'saldo_devedor',
                'Natureza': 'situacao',
                'Porta- dor': 'centro_custo'
            }
        # Mapeamento para arquivos contábeis CTBR140
        elif 'ctbr140' in filename:
            return {
                'Codigo': 'conta_contabil',
                'Descricao': 'descricao_conta',
                'Codigo.1': 'codigo_fornecedor',
                'Descricao.1': 'descricao_fornecedor',
                'Saldo anterior': 'saldo_anterior',
                'Debito': 'debito',
                'Credito': 'credito',
                'Movimento do periodo': 'movimento_periodo',
                'Saldo atual': 'saldo_atual'
            }

        # Mapeamento para arquivos contábeis CTBR040
        elif 'ctbr040' in filename:
            return {
                'Conta': 'conta_contabil',
                'Descricao': 'descricao_item',   
                'Saldo anterior': 'saldo_anterior',
                'Debito': 'debito',
                'Credito': 'credito',
                'Mov  periodo': 'movimento_periodo',
                'Saldo atual': 'saldo_atual'
            }
        else:
            raise ValueError(f"Tipo de planilha não reconhecido: {file_path.name}")
    
    def process_data(self):
        """Processa os dados importados e gera resultados da conciliação"""
        try:
            self.conn.execute("BEGIN TRANSACTION")  # Inicia transação
            cursor = self.conn.cursor()
            
            # Obtém período de referência
            data_inicial, data_final = self._get_datas_referencia()
            # Limpa tabela de resultados anterior
            cursor.execute(f"DELETE FROM {self.settings.TABLE_RESULTADO}")
            
            # Insere dados financeiros na tabela de resultados
            query_financeiro = f"""
                INSERT INTO {self.settings.TABLE_RESULTADO}
                (codigo_fornecedor, descricao_fornecedor, saldo_financeiro, status)
                SELECT 
                    fornecedor as codigo_fornecedor,
                    MAX(fornecedor) as descricao_fornecedor,
                    SUM(saldo_devedor) as saldo_financeiro,
                    'PENDENTE' as status
                FROM 
                    {self.settings.TABLE_FINANCEIRO}
                WHERE 
                    excluido = 0
                    AND data_vencimento BETWEEN '{data_inicial}' AND '{data_final}'
                GROUP BY 
                    fornecedor
            """
            cursor.execute(query_financeiro)
            
            # Atualiza com dados contábeis correspondentes
            query_contabil_update = f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET 
                    saldo_contabil = (
                        SELECT COALESCE(SUM(saldo_atual),0)
                        FROM {self.settings.TABLE_MODELO1} m
                        WHERE 
                            (m.codigo_fornecedor IS NOT NULL AND m.codigo_fornecedor <> '' AND m.codigo_fornecedor = {self.settings.TABLE_RESULTADO}.codigo_fornecedor)
                            OR m.descricao_conta LIKE '%' || {self.settings.TABLE_RESULTADO}.codigo_fornecedor || '%'
                    ),
                    detalhes = (
                        SELECT GROUP_CONCAT(COALESCE(m.tipo_fornecedor,'') || ': ' || m.saldo_atual, ' | ')
                        FROM {self.settings.TABLE_MODELO1} m
                        WHERE 
                            (m.codigo_fornecedor IS NOT NULL AND m.codigo_fornecedor <> '' AND m.codigo_fornecedor = {self.settings.TABLE_RESULTADO}.codigo_fornecedor)
                            OR m.descricao_conta LIKE '%' || {self.settings.TABLE_RESULTADO}.codigo_fornecedor || '%'
                    )
                WHERE EXISTS (
                    SELECT 1
                    FROM {self.settings.TABLE_MODELO1} m2
                    WHERE 
                        (m2.codigo_fornecedor IS NOT NULL AND m2.codigo_fornecedor <> '' AND m2.codigo_fornecedor = {self.settings.TABLE_RESULTADO}.codigo_fornecedor)
                        OR m2.descricao_conta LIKE '%' || {self.settings.TABLE_RESULTADO}.codigo_fornecedor || '%'
                )
            """
            cursor.execute(query_contabil_update)
            
            # Insere dados contábeis que não tiveram match financeiro
            query_contabeis_sem_match = f"""
                INSERT INTO {self.settings.TABLE_RESULTADO}
                (codigo_fornecedor, descricao_fornecedor, saldo_contabil, status, detalhes)
                SELECT 
                    COALESCE(NULLIF(TRIM(codigo_fornecedor), ''),
                            CASE 
                                WHEN INSTR(descricao_conta, ' ') > 0 THEN SUBSTR(descricao_conta, 1, INSTR(descricao_conta, ' ') - 1)
                                ELSE descricao_conta
                            END) as codigo_fornecedor,
                    COALESCE(NULLIF(TRIM(descricao_fornecedor), ''), descricao_conta) as descricao_fornecedor,
                    saldo_atual as saldo_contabil,
                    'PENDENTE' as status,
                    tipo_fornecedor as detalhes
                FROM 
                    {self.settings.TABLE_MODELO1} m
                WHERE 
                    (descricao_conta LIKE 'FORNEC%')
                    AND NOT EXISTS (
                        SELECT 1
                        FROM {self.settings.TABLE_RESULTADO} r
                        WHERE r.codigo_fornecedor = COALESCE(NULLIF(TRIM(m.codigo_fornecedor), ''),
                                                            CASE 
                                                                WHEN INSTR(m.descricao_conta, ' ') > 0 THEN SUBSTR(m.descricao_conta, 1, INSTR(m.descricao_conta, ' ') - 1)
                                                                ELSE m.descricao_conta
                                                            END)
                    )
            """
            cursor.execute(query_contabeis_sem_match)

            # Calcula diferenças e define status
            query_diferenca = f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET 
                    diferenca = ROUND(COALESCE(saldo_contabil,0) - COALESCE(saldo_financeiro,0), 2),
                    status = CASE 
                        WHEN ABS(COALESCE(saldo_contabil,0) - COALESCE(saldo_financeiro,0)) < 0.01 THEN 'OK'
                        WHEN COALESCE(saldo_contabil,0) = 0 AND COALESCE(saldo_financeiro,0) > 0 THEN 'PENDENTE'
                        WHEN COALESCE(saldo_financeiro,0) = 0 AND COALESCE(saldo_contabil,0) > 0 THEN 'PENDENTE'
                        ELSE 'DIVERGENTE'
                    END
            """
            cursor.execute(query_diferenca)
            
            
            # Para fornecedores com status DIVERGENTE, busca detalhes na tabela contas_itens
            query_investigacao = f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET detalhes = (
                    SELECT 'Divergência: R$ ' || ABS(diferenca) || 
                        '. Itens Contábeis encontrados: ' || 
                        COALESCE(
                            (SELECT GROUP_CONCAT(
                                    'Item: ' || ci.descricao_item || 
                                    ' (Valor: R$ ' || ci.saldo_atual || ')', 
                                    '; '
                                )
                                FROM contas_itens ci
                                WHERE ci.conta_contabil LIKE '%' || {self.settings.TABLE_RESULTADO}.codigo_fornecedor || '%'
                                LIMIT 5),  -- Limita a 5 itens para não ficar muito longo
                            'Nenhum item específico encontrado'
                        )
                )
                WHERE status = 'DIVERGENTE'
                AND EXISTS (
                    SELECT 1 FROM contas_itens ci2 
                    WHERE ci2.conta_contabil LIKE '%' || {self.settings.TABLE_RESULTADO}.codigo_fornecedor || '%'
                )
            """
            cursor.execute(query_investigacao)
            
            # Para fornecedores divergentes sem itens específicos na contas_itens
            cursor.execute(f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET detalhes = 'Divergência: R$ ' || ABS(diferenca) || 
                            '. Investigar manualmente no sistema. Nenhum item contábil específico encontrado para análise automática.'
                WHERE status = 'DIVERGENTE' 
                AND (detalhes IS NULL OR detalhes = '' OR detalhes LIKE '%Conciliação%')
            """)
            
            # Classifica por ordem de importância (magnitude da diferença)
            try:
                cursor.execute(f"""
                    UPDATE {self.settings.TABLE_RESULTADO}
                    SET ordem_importancia = (
                        SELECT COUNT(*) 
                        FROM {self.settings.TABLE_RESULTADO} r2 
                        WHERE ABS(COALESCE(r2.diferenca,0)) >= ABS(COALESCE({self.settings.TABLE_RESULTADO}.diferenca,0))
                    )
                """)
            except Exception as rank_error:
                logger.error(f"Erro ao classificar por importância: {rank_error}")

            # Atualiza detalhes para registros não divergentes (mantém informações básicas)
            cursor.execute(f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET detalhes = 
                    CASE 
                        WHEN status = 'OK' THEN 'Conciliação perfeita'
                        WHEN status = 'PENDENTE' THEN 'Financeiro: ' || COALESCE(saldo_financeiro,0) || 
                                                    ' | Contábil: ' || COALESCE(saldo_contabil,0) || 
                                                    ' | Diferença: ' || COALESCE(diferenca,0)
                        ELSE detalhes  -- Mantém os detalhes da investigação para divergências
                    END
            """)
            
            self.conn.commit()  # Confirma transação
            logger.info("Processamento de dados concluído com sucesso")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao processar dados: {e}", exc_info=True)
            self.conn.rollback()  # Reverte em caso de erro
            return False
    
    def _apply_styles(self, worksheet):
        """Aplica estilos visuais à planilha Excel"""
        # Define estilos para cabeçalho
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        # Aplica estilos ao cabeçalho
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border
        
        # Identifica colunas numéricas para formatação
        header = [c.value for c in worksheet[1]]
        numeric_headers = set([
            'Saldo Contábil','Saldo Financeiro','Diferença',
            'Saldo Anterior','Débito','Crédito','Saldo Atual',
            'Valor Original','Saldo Devedor','Quantidade','Valor Unitário','Valor Total'
        ])
        numeric_cols_idx = [i+1 for i, h in enumerate(header) if h in numeric_headers]

        # Aplica bordas e formatação numérica
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                if cell.row > 1 and cell.column in numeric_cols_idx:
                    cell.number_format = '#,##0.00'  # Formato monetário
        
        # Ajusta largura das colunas automaticamente
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value is not None and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def _apply_enhanced_styles(self, worksheet, stats):
        """Aplica estilos visuais melhorados com formatação condicional avançada"""
        # Define estilos para cabeçalho
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        # Aplica estilos ao cabeçalho
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border
        
        # Identifica índices de colunas para formatação
        header = [c.value for c in worksheet[1]]
        saldo_contabil_idx = header.index("Saldo Contábil") + 1 if "Saldo Contábil" in header else None
        saldo_financeiro_idx = header.index("Saldo Financeiro") + 1 if "Saldo Financeiro" in header else None
        diferenca_idx = header.index("Diferença (Contábil - Financeiro)") + 1 if "Diferença (Contábil - Financeiro)" in header else None
        status_idx = header.index("Status") + 1 if "Status" in header else None
        
        # Cores para formatação condicional
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        
        red_font = Font(color="9C0006", bold=True)
        green_font = Font(color="006100", bold=True)
        
        # Aplica formatação condicional para cada linha
        for row in range(2, worksheet.max_row + 1):
            # Formata valores negativos em vermelho
            if diferenca_idx:
                diff_cell = worksheet.cell(row=row, column=diferenca_idx)
                if diff_cell.value is not None and diff_cell.value < 0:
                    diff_cell.font = red_font
            
            # Formatação baseada no status
            if status_idx:
                status_cell = worksheet.cell(row=row, column=status_idx)
                status_value = status_cell.value if status_cell.value else ""
                
                if status_value == 'OK':
                    for col in range(1, worksheet.max_column + 1):
                        worksheet.cell(row=row, column=col).fill = green_fill
                elif status_value == 'DIVERGENTE':
                    for col in range(1, worksheet.max_column + 1):
                        worksheet.cell(row=row, column=col).fill = red_fill
                elif status_value == 'PENDENTE':
                    for col in range(1, worksheet.max_column + 1):
                        worksheet.cell(row=row, column=col).fill = yellow_fill
            
            # Aplica bordas a todas as células
            for col in range(1, worksheet.max_column + 1):
                worksheet.cell(row=row, column=col).border = thin_border
        
        # Ajusta largura das colunas automaticamente
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value is not None and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def _apply_metadata_styles(self, worksheet):
        """Aplica estilos à aba de metadados"""
        # Define estilos
        title_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        title_font = Font(color="FFFFFF", bold=True, size=14)
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        # Formata título
        for cell in worksheet[1]:
            cell.fill = title_fill
            cell.font = title_font
            cell.border = thin_border
        
        # Formata cabeçalho
        for cell in worksheet[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
        
        # Aplica bordas a todas as células
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        
        # Ajusta largura das colunas
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value is not None and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def _get_datas_referencia(self):
        """Retorna datas de referência para o processamento"""
        cal = Brazil()
        hoje = datetime.now().date()
        data_inicial = hoje.replace(day=1)
        data_inicial = cal.add_working_days(data_inicial - timedelta(days=1), 1)
        data_final = hoje
        logger.warning("USANDO VERSÃO TEMPORÁRIA DE _get_datas_referencia() - IGNORANDO VERIFICAÇÃO DE DIA 20/ÚLTIMO DIA")
        return data_inicial.strftime("%d/%m/%Y"), data_final.strftime("%d/%m/%Y")
        """Retorna datas de referência ajustando para feriados"""
        # cal = Brazil()
        # hoje = datetime.now().date()

        # # Ajusta a data se não for dia útil
        # if not cal.is_working_day(hoje):
        #     hoje = cal.add_working_days(hoje, 1)
        #     logger.warning(f"Data ajustada para o próximo dia útil: {hoje.strftime('%d/%m/%Y')}")

        # # Verifica se é dia 20 ou último dia do mês
        # if hoje.day == 20 or hoje.day == self._ultimo_dia_mes(hoje).day:
        #     if hoje.day == 20:
        #         data_inicial = hoje.replace(day=1)
        #     else:
        #         data_inicial = hoje.replace(day=1)
            
        #     # Garante que as datas são dias úteis
        #     data_inicial = cal.add_working_days(data_inicial - timedelta(days=1), 1)
        #     data_final = hoje
            
        #     return data_inicial.strftime("%d/%m/%Y"), data_final.strftime("%d/%m/%Y")
        # else:
        #     raise ValueError("Processamento só deve ocorrer no dia 20 ou último dia do mês")

    def _ultimo_dia_mes(self, date):
        """Retorna o último dia do mês da data fornecida"""
        next_month = date.replace(day=28) + timedelta(days=4)
        return next_month - timedelta(days=next_month.day)

    def validate_output(self, output_path):
        """Valida a estrutura do arquivo Excel gerado, incluindo a nova aba Metadados"""
        try:
            wb = openpyxl.load_workbook(output_path)
            # Verifica abas obrigatórias
            required_sheets = ['Resumo', 'Títulos a Pagar', 'Balancete', 'Metadados']
            for sheet in required_sheets:
                if sheet not in wb.sheetnames:
                    raise ValueError(f"Aba '{sheet}' não encontrada no arquivo gerado")
            
            # Verifica colunas obrigatórias na aba Resumo
            resumo = wb['Resumo']
            expected_columns = [
                'Código Fornecedor', 'Descrição Fornecedor', 
                'Saldo Contábil', 'Saldo Financeiro', 'Diferença (Contábil - Financeiro)'
            ]
            header = [cell.value for cell in resumo[1]]
            for col in expected_columns:
                if col not in header:
                    raise ValueError(f"Coluna '{col}' não encontrada na aba Resumo")
            
            # Verifica se a aba Metadados contém informações essenciais
            metadados = wb['Metadados']
            has_essentials = False
            for row in metadados.iter_rows(values_only=True):
                if 'Data e Hora do Processamento' in row or 'Fórmula de Cálculo' in row:
                    has_essentials = True
                    break
                    
            if not has_essentials:
                raise ValueError("Aba Metadados não contém informações essenciais")
                
            return True
        except Exception as e:
            logger.error(f"Validação falhou: {e}")
            return False

    def validate_data_consistency(self):
        """Valida consistência geral dos dados importados"""
        try:
            cursor = self.conn.cursor()
            # Verifica fornecedores financeiros sem correspondência contábil
            query = f"""
                SELECT COUNT(*) 
                FROM {self.settings.TABLE_FINANCEIRO} f
                WHERE NOT EXISTS (
                    SELECT 1 FROM {self.settings.TABLE_MODELO1} m
                    WHERE (m.codigo_fornecedor IS NOT NULL AND m.codigo_fornecedor <> '' AND m.codigo_fornecedor = f.fornecedor)
                       OR m.descricao_conta LIKE '%' || f.fornecedor || '%'
                )
                AND f.excluido = 0
            """
            cursor.execute(query)
            missing_suppliers = cursor.fetchone()[0]
            if missing_suppliers > 0:
                logger.warning(f"{missing_suppliers} fornecedores financeiros sem correspondência contábil")
            
            # Compara totais financeiros e contábeis
            query = f"""
                SELECT 
                    (SELECT SUM(saldo_devedor) FROM {self.settings.TABLE_FINANCEIRO} WHERE excluido = 0) as total_financeiro,
                    (SELECT SUM(saldo_atual) FROM {self.settings.TABLE_MODELO1} WHERE tipo_fornecedor LIKE 'FORNEC%') as total_contabil
            """
            cursor.execute(query)
            totals = cursor.fetchone()
            logger.info(f"Total Financeiro: {totals[0]} | Total Contábil: {totals[1]}")
            return True
        except Exception as e:
            logger.error(f"Erro na validação de consistência: {e}")
            return False

    def export_to_excel(self):
        """Exporta resultados para arquivo Excel formatado com metadados e formatação melhorada"""
        logger.info("Iniciando exportação para Excel...")
        data_inicial, data_final = self._get_datas_referencia()
        output_path = self.settings.RESULTS_DIR / f"CONCILIACAO_{data_inicial.replace('/', '-')}_a_{data_final.replace('/', '-')}.xlsx"
        
        try:
            if not self.conn:
                logger.error("Tentativa de exportação com conexão fechada")
                raise RuntimeError("Conexão com o banco de dados não está aberta")
            
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            
            # Query para obter estatísticas de processamento
            query_stats = f"""
                SELECT 
                    COUNT(*) as total_registros,
                    SUM(CASE WHEN status = 'OK' THEN 1 ELSE 0 END) as conciliados_ok,
                    SUM(CASE WHEN status = 'DIVERGENTE' THEN 1 ELSE 0 END) as divergentes,
                    SUM(CASE WHEN status = 'PENDENTE' THEN 1 ELSE 0 END) as pendentes,
                    SUM(ABS(diferenca)) as total_diferenca
                FROM 
                    {self.settings.TABLE_RESULTADO}
            """
            stats = pd.read_sql(query_stats, self.conn).iloc[0]
            
            # Query para aba Resumo
            query_resumo = f"""
                SELECT 
                    codigo_fornecedor as "Código Fornecedor",
                    descricao_fornecedor as "Descrição Fornecedor",
                    saldo_contabil as "Saldo Contábil",
                    saldo_financeiro as "Saldo Financeiro",
                    diferenca as "Diferença (Contábil - Financeiro)",
                    CASE 
                        WHEN COALESCE(saldo_contabil,0) - COALESCE(saldo_financeiro,0) > 0 THEN 'Contábil > Financeiro'
                        WHEN COALESCE(saldo_contabil,0) - COALESCE(saldo_financeiro,0) < 0 THEN 'Financeiro > Contábil'
                        ELSE 'Valores Iguais'
                    END as "Tipo Diferença",
                    status as "Status",
                    detalhes as "Detalhes"
                FROM 
                    {self.settings.TABLE_RESULTADO}
                ORDER BY 
                    ordem_importancia
            """
            df_resumo = pd.read_sql(query_resumo, self.conn)
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
            
            # Query para aba Títulos a Pagar
            query_financeiro = f"""
                SELECT 
                    fornecedor as "Fornecedor",
                    titulo as "Título",
                    parcela as "Parcela",
                    tipo_titulo as "Tipo Título",
                    data_emissao as "Data Emissão",
                    data_vencimento as "Data Vencimento",
                    valor_original as "Valor Original",
                    saldo_devedor as "Saldo Devedor",
                    situacao as "Situação",
                    conta_contabil as "Conta Contábil",
                    centro_custo as "Centro Custo"
                FROM 
                    {self.settings.TABLE_FINANCEIRO}
                WHERE 
                    excluido = 0
                    AND data_vencimento BETWEEN '{data_inicial}' AND '{data_final}'
                ORDER BY 
                    fornecedor, titulo
            """
            df_financeiro = pd.read_sql(query_financeiro, self.conn)
            df_financeiro.to_excel(writer, sheet_name='Títulos a Pagar', index=False)
            
            # Query para aba Balancete
            query_contabil = f"""
                SELECT 
                    conta_contabil as "Conta Contábil",
                    descricao_conta as "Descrição Conta",
                    COALESCE(codigo_fornecedor,'') as "Código Fornecedor",
                    COALESCE(descricao_fornecedor,'') as "Descrição Fornecedor",
                    saldo_anterior as "Saldo Anterior",
                    debito as "Débito",
                    credito as "Crédito",
                    saldo_atual as "Saldo Atual",
                    tipo_fornecedor as "Tipo Fornecedor"
                FROM 
                    {self.settings.TABLE_MODELO1}
                WHERE 
                    descricao_conta LIKE 'FORNEC%'
                ORDER BY 
                    tipo_fornecedor, conta_contabil
            """
            df_contabil = pd.read_sql(query_contabil, self.conn)
            df_contabil.to_excel(writer, sheet_name='Balancete', index=False)
            
            # Cria aba de Metadados
            metadata = {
                'Item': [
                    'Data e Hora do Processamento',
                    'Período de Referência',
                    'Total de Fornecedores Processados',
                    'Conciliações OK',
                    'Conciliações Divergentes',
                    'Conciliações Pendentes',
                    'Total de Diferenças (R$)',
                    'Fórmula de Cálculo',
                    'Legenda de Status'
                ],
                'Valor': [
                    datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                    f'{data_inicial} a {data_final}',
                    int(stats['total_registros']),
                    int(stats['conciliados_ok']),
                    int(stats['divergentes']),
                    int(stats['pendentes']),
                    f"R$ {stats['total_diferenca']:,.2f}",
                    'Diferença = Saldo Contábil - Saldo Financeiro',
                    'OK: Valores iguais | DIVERGENTE: Valores diferentes | PENDENTE: Sem correspondência'
                ]
            }
            df_metadata = pd.DataFrame(metadata)
            df_metadata.to_excel(writer, sheet_name='Metadados', index=False)
            
            writer.close()
            
            # Aplica estilos ao arquivo gerado
            workbook = openpyxl.load_workbook(output_path)
            
            # Aplica estilos à aba Metadados
            if 'Metadados' in workbook.sheetnames:
                meta_sheet = workbook['Metadados']
                self._apply_metadata_styles(meta_sheet)
            
            # Aplica estilos melhorados à aba Resumo
            if 'Resumo' in workbook.sheetnames:
                resumo_sheet = workbook['Resumo']
                self._apply_enhanced_styles(resumo_sheet, stats)
                
                # Adiciona informações de cabeçalho
                resumo_sheet['J1'] = "Data de Referência:"
                resumo_sheet['K1'] = f"{data_inicial} a {data_final}"
                resumo_sheet['J2'] = "Fórmula Diferença:"
                resumo_sheet['K2'] = "Saldo Contábil - Saldo Financeiro"
            
            # Aplica estilos às demais abas
            for sheetname in workbook.sheetnames:
                if sheetname != 'Metadados':  # Já estilizamos a Metadados separadamente
                    sheet = workbook[sheetname]
                    self._apply_styles(sheet)
            
            workbook.save(output_path)
            
            # Valida o arquivo gerado
            if not self.validate_output(output_path):
                raise ValueError("A validação da planilha gerada falhou")
            
            return output_path
        except Exception as e:
            logger.error(f"Erro ao exportar resultados: {e}")
            return None

    def close(self):
        """Fecha a conexão com o banco de dados"""
        if self.conn:
            try:
                self.conn.close()
                logger.info("Conexão com o banco de dados fechada")
            except Exception as e:
                logger.error(f"Erro ao fechar conexão: {e}")

    def __enter__(self):
        """Suporte para context manager (with statement)"""
        self._initialize_database()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Suporte para context manager (with statement)"""
        self.close()
