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

logger = configure_logger()

class DatabaseManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(DatabaseManager, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        self.settings = Settings()
        self.conn = None
        self.logger = configure_logger()
        self._initialize_database()
        self._initialized = True

    def _initialize_database(self):
        try:
            self.conn = sqlite3.connect(self.settings.DB_PATH, timeout=10)
            cursor = self.conn.cursor()
            
            # Tabela para Títulos a Pagar (Financeiro)
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
            
            # Tabela para Balancete Modelo 1
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS {self.settings.TABLE_MODELO1} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conta_contabil TEXT,
                    descricao_conta TEXT,
                    saldo_anterior REAL DEFAULT 0,
                    debito REAL DEFAULT 0,
                    credito REAL DEFAULT 0,
                    saldo_atual REAL DEFAULT 0,
                    tipo_fornecedor TEXT,
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Tabela para Contas x Itens
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
            
            # Tabela para Resultados da Conciliação
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
                    data_processamento TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            self.conn.commit()
            logger.info("Banco de dados inicializado com sucesso")
            
        except Exception as e:
            logger.error(f"Erro ao inicializar banco de dados: {e}")
            raise

    def aplicar_sugestoes_colunas(self, df, missing_mappings):
        candidates = df.columns.tolist()
        lower_map = {c.lower(): c for c in candidates}

        # Mapeamento manual específico para a planilha financeira
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

        # Aplicar mapeamento manual primeiro
        for src, dest in manual_mapping.items():
            if src in df.columns and dest in missing_mappings:
                df.rename(columns={src: dest}, inplace=True)
                logger.warning(f"Mapeamento manual aplicado: '{src}' → '{dest}'")
                missing_mappings.remove(dest)

        for db_col in missing_mappings:
            # correspondência exata 
            if db_col.lower() in lower_map:
                match = lower_map[db_col.lower()]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}' (case-insensitive match)")
                df.rename(columns={match: db_col}, inplace=True)
                continue

            # tentativa com difflib usando as opções originais
            similar = get_close_matches(db_col, candidates, n=1, cutoff=0.6)
            if similar:
                match = similar[0]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}'")
                df.rename(columns={match: db_col}, inplace=True)
                continue

            # fuzzy case-insensitive
            similar_lower = get_close_matches(db_col.lower(), list(lower_map.keys()), n=1, cutoff=0.6)
            if similar_lower:
                match = lower_map[similar_lower[0]]
                logger.warning(f"Sugestão aplicada: '{match}' → '{db_col}' (fuzzy case-insensitive)")
                df.rename(columns={match: db_col}, inplace=True)

        return df

    def import_from_excel(self, file_path, table_name):
        try:
            # Carrega Excel
            df = pd.read_excel(file_path, header=1)  
            
            logger.info(f"Colunas originais em {file_path}: {df.columns.tolist()}")

            # Limpeza de colunas
            df.columns = df.columns.str.replace(r'_x000D_\n', ' ', regex=True).str.strip()
            logger.info(f"Colunas após limpeza: {df.columns.tolist()}")

            # Obter mapeamento de colunas (invertido para o formato correto)
            column_mapping = self._get_column_mapping(Path(file_path))
            
            # Renomear colunas conforme mapeamento
            df.rename(columns=column_mapping, inplace=True)
            logger.info(f"Colunas após mapeamento: {df.columns.tolist()}")
            
            # Verifica colunas esperadas vs colunas presentes
            expected_columns = self.get_expected_columns(table_name)
            missing_mappings = [col for col in expected_columns if col not in df.columns]

            if missing_mappings:
                logger.warning(f"Colunas mapeadas não encontradas: {missing_mappings}")
                df = self.aplicar_sugestoes_colunas(df, missing_mappings)
                
                # Verificar novamente após aplicar sugestões
                remaining_missing = [col for col in expected_columns if col not in df.columns]
                if remaining_missing:
                    # Tratamento especial para colunas ausentes
                    if 'parcela' in remaining_missing and 'titulo' in df.columns:
                        df['parcela'] = df['titulo'].str.extract(r'(\d+)$').fillna('1')
                        logger.warning("Coluna 'parcela' criada a partir do título")
                        remaining_missing.remove('parcela')
                    
                    if 'conta_contabil' in remaining_missing:
                        df['conta_contabil'] = 'CONTA_PADRAO' 
                        logger.warning("Coluna 'conta_contabil' preenchida com valor padrão")
                        remaining_missing.remove('conta_contabil')
                    
                    if remaining_missing:
                        logger.error(f"Colunas obrigatórias ausentes após tratamento: {remaining_missing}")
                        raise ValueError(f"Colunas obrigatórias ausentes: {remaining_missing}")

            # Limpeza e tratamento específico para cada tipo de planilha
            df = self._clean_dataframe(df, table_name.lower())

            # Garantir que as colunas do DataFrame correspondam exatamente às colunas da tabela
            table_columns = [col[1] for col in self.conn.execute(f"PRAGMA table_info({table_name})").fetchall()]
            df = df[[col for col in df.columns if col in table_columns]]
            table_columns = [col[1] for col in self.conn.execute(f"PRAGMA table_info({table_name})").fetchall()]
            df = df[[col for col in df.columns if col in table_columns]]
            df.to_sql(table_name, self.conn, if_exists='append', index=False)
            logger.info(f"Dados importados para '{table_name}' com sucesso.")
            
            return True  
            
        except Exception as e:
            logger.error(f"Falha ao importar {file_path}: {e}", exc_info=True)
            return False
    
    def get_expected_columns(self, table_name):
        """Retorna as colunas esperadas para cada tabela"""
        if table_name == self.settings.TABLE_FINANCEIRO:
            return [
                'fornecedor', 'titulo', 'parcela', 'tipo_titulo',
                'data_emissao', 'data_vencimento', 'valor_original',
                'saldo_devedor', 'situacao', 'conta_contabil', 'centro_custo'
            ]
        elif table_name == self.settings.TABLE_MODELO1:
            return [
                'conta_contabil', 'descricao_conta', 'saldo_anterior',
                'debito', 'credito', 'saldo_atual'  
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
        """Realiza a limpeza dos dados conforme o tipo de planilha"""
        try:
            # Converter todas as colunas para string (evita problemas com tipos mistos)
            df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
            
            # Remover linhas totalmente vazias
            df = df.replace(['nan', 'None', ''], np.nan)
            df = df.dropna(how='all')
            
            # Processamento específico para cada tipo de planilha
            if sheet_type == 'financeiro':
                df = self._clean_financeiro_data(df)
            elif sheet_type == 'modelo1':
                df = self._clean_modelo1_data(df)
            elif sheet_type == 'contas_itens':
                df = self._clean_contas_itens_data(df)
            
            df = df.drop_duplicates()
            logger.info(f"DataFrame limpo - shape final: {df.shape}")
            return df
            
        except Exception as e:
            logger.error(f"Erro na limpeza dos dados ({sheet_type}): {str(e)}", exc_info=True)
            raise

    def _clean_financeiro_data(self, df):
        """Limpeza específica para dados financeiros"""
        # 1. Aplicar filtros iniciais
        if 'tipo_titulo' in df.columns:
            df = df[~df['tipo_titulo'].isin(['NDF', 'PA'])]
            logger.info(f"Registros após filtrar NDF/PA: {len(df)}")
        
        if 'conta_contabil' in df.columns:
            df = df[df['conta_contabil'].str.startswith('2010201', na=False)]
            logger.info(f"Registros após filtrar fornecedores nacionais: {len(df)}")

        # 2. Tratamento de datas
        date_cols = ['data_emissao', 'data_vencimento']
        for col in date_cols:
            if col in df.columns:
                try:
                    df[col] = pd.to_datetime(
                        df[col], 
                        errors='coerce',
                        format='mixed',
                        dayfirst=True
                    )
                    
                    if df[col].isna().any():
                        invalid_dates = df[df[col].isna()][col].count()
                        logger.warning(f"{invalid_dates} registros com datas inválidas na coluna {col}")
                        
                except Exception as e:
                    logger.error(f"Erro crítico ao converter datas na coluna {col}: {str(e)}")
                    raise

        # 3. Tratamento de valores numéricos
        num_cols = ['valor_original', 'saldo_devedor']
        for col in num_cols:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.replace(r'[^\d,-]', '', regex=True) 
                    .str.replace(',', '.')
                    .replace('', np.nan)
                    .astype(float)
                    .fillna(0)
                )
        
        # 4. Tratamento de texto
        text_cols = ['fornecedor', 'titulo', 'tipo_titulo']
        for col in text_cols:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.strip()
                    .str.replace(r'\s+', ' ', regex=True)
                    .replace('nan', '')
                )
        
        # 5. Criar coluna de parcela se necessário
        if 'titulo' in df.columns and 'parcela' not in df.columns:
            df['parcela'] = df['titulo'].str.extract(r'(\d+)$').fillna('1')

        return df

    def _clean_modelo1_data(self, df):
        """Limpeza específica para dados do modelo 1"""
        num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str)
                    .str.replace(r'[^\d,-]', '', regex=True)
                    .str.replace(',', '.'),
                    errors='coerce'
                ).fillna(0)
        
        if 'tipo_fornecedor' not in df.columns and 'descricao_conta' in df.columns:
            df['tipo_fornecedor'] = df['descricao_conta'].apply(
                lambda x: 'FORNECEDOR' if 'FORNEC' in str(x).upper() else 'OUTROS'
            )

        return df

    def _clean_contas_itens_data(self, df):
        """Limpeza específica para dados de contas x itens"""
        num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str)
                    .str.replace(r'[^\d,-]', '', regex=True)
                    .str.replace(',', '.'),
                    errors='coerce'
                ).fillna(0)
        
        if 'item' not in df.columns:
            df['item'] = df.get('descricao_item', '').str[:50] 
            
        if 'quantidade' not in df.columns:
            df['quantidade'] = 1
            
        if 'valor_unitario' not in df.columns and 'saldo_atual' in df.columns:
            df['valor_unitario'] = df['saldo_atual']
            
        if 'valor_total' not in df.columns and 'saldo_atual' in df.columns:
            df['valor_total'] = df['saldo_atual']

        return df
    
    def _get_column_mapping(self, file_path: Path):
        """Retorna o mapeamento de colunas apropriado com base no nome do arquivo"""
        filename = file_path.stem.lower()
        
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
                'Saldo atual': 'saldo_atual',
                'Descricao.1': 'tipo_fornecedor'
            }
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
        """Processa os dados e gera a conciliação conforme especificado"""
        try:
            self.conn.execute("BEGIN TRANSACTION")
            cursor = self.conn.cursor()
            
            data_inicial, data_final = self._get_datas_referencia()
            cursor.execute(f"DELETE FROM {self.settings.TABLE_RESULTADO}")

            # 1. Processar dados financeiros (agrupar por fornecedor)
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
            
            # 2. Processar dados contábeis (Modelo 1)
            query_contabil = f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET 
                    saldo_contabil = (
                        SELECT SUM(saldo_atual)
                        FROM {self.settings.TABLE_MODELO1}
                        WHERE descricao_conta LIKE '%' || codigo_fornecedor || '%'
                    ),
                    detalhes = (
                        SELECT GROUP_CONCAT(tipo_fornecedor || ': ' || saldo_atual, ' | ')
                        FROM {self.settings.TABLE_MODELO1}
                        WHERE descricao_conta LIKE '%' || codigo_fornecedor || '%'
                    )
                WHERE EXISTS (
                    SELECT 1
                    FROM {self.settings.TABLE_MODELO1}
                    WHERE descricao_conta LIKE '%' || codigo_fornecedor || '%'
                )
            """
            cursor.execute(query_contabil)
            
            # 3. Calcular diferenças e atualizar status
            query_diferenca = f"""
                UPDATE {self.settings.TABLE_RESULTADO}
                SET 
                    diferenca = ROUND(saldo_contabil - saldo_financeiro, 2),
                    status = CASE 
                        WHEN ABS(saldo_contabil - saldo_financeiro) < 0.01 THEN 'OK'
                        WHEN saldo_contabil = 0 AND saldo_financeiro > 0 THEN 'PENDENTE'
                        WHEN saldo_financeiro = 0 AND saldo_contabil > 0 THEN 'PENDENTE'
                        ELSE 'DIVERGENTE'
                    END,
                    detalhes = CASE
                        WHEN ABS(saldo_contabil - saldo_financeiro) < 0.01 THEN 'Conciliação OK'
                        ELSE 'Verificar lançamentos para o fornecedor'
                    END
            """
            cursor.execute(query_diferenca)
            
            # 4. Inserir fornecedores contábeis que não estão no financeiro
            query_fornecedores_contabeis = f"""
                INSERT INTO {self.settings.TABLE_RESULTADO}
                (codigo_fornecedor, descricao_fornecedor, saldo_contabil, status, detalhes)
                SELECT 
                    SUBSTR(descricao_conta, 1, INSTR(descricao_conta, ' ') - 1) as codigo_fornecedor,
                    descricao_conta as descricao_fornecedor,
                    saldo_atual as saldo_contabil,
                    'DIVERGENTE' as status,
                    tipo_fornecedor as detalhes
                FROM 
                    {self.settings.TABLE_MODELO1}
                WHERE 
                    descricao_conta LIKE 'FORNEC%'
                    AND NOT EXISTS (
                        SELECT 1
                        FROM {self.settings.TABLE_RESULTADO}
                        WHERE descricao_fornecedor LIKE '%' || SUBSTR(descricao_conta, 1, INSTR(descricao_conta, ' ') - 1) || '%'
                    )
            """
            cursor.execute(query_fornecedores_contabeis)
            
            self.conn.commit()
            logger.info("Processamento de dados concluído com sucesso")
            return True
            
        except Exception as e:
            logger.error(f"Erro ao processar dados: {e}", exc_info=True)
            self.conn.rollback()
            return False
    
    def _apply_styles(self, worksheet):
        """Aplica formatação à planilha Excel"""
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border
        
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                if cell.column_letter in ['K', 'L', 'M', 'N']:
                    cell.number_format = '#,##0.00'
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
                
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def _get_datas_referencia(self):
        """Versão temporária para testes - ignora verificação de dia 20/último dia"""
        cal = Brazil()
        hoje = datetime.now().date()
        
        # Sempre usa o primeiro dia do mês como data inicial
        data_inicial = hoje.replace(day=1)
        
        # Garante que as datas são dias úteis
        data_inicial = cal.add_working_days(data_inicial - timedelta(days=1), 1)
        data_final = hoje
        
        logger.warning("USANDO VERSÃO TEMPORÁRIA DE _get_datas_referencia() - IGNORANDO VERIFICAÇÃO DE DIA 20/ÚLTIMO DIA")
        return data_inicial.strftime("%d/%m/%Y"), data_final.strftime("%d/%m/%Y")
        # """Retorna datas de referência ajustando para feriados"""
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
        """Retorna o último dia do mês para uma data dada"""
        next_month = date.replace(day=28) + timedelta(days=4)
        return next_month - timedelta(days=next_month.day)

    def validate_output(self, output_path):
        """Valida se o arquivo gerado atende aos requisitos"""
        try:
            wb = openpyxl.load_workbook(output_path)
            
            required_sheets = ['Resumo', 'Títulos a Pagar', 'Balancete', 'Contas x Itens']
            for sheet in required_sheets:
                if sheet not in wb.sheetnames:
                    raise ValueError(f"Aba '{sheet}' não encontrada no arquivo gerado")
            
            resumo = wb['Resumo']
            expected_columns = [
                'Código Fornecedor', 'Descrição Fornecedor', 
                'Saldo Contábil', 'Saldo Financeiro', 'Diferença'
            ]
            header = [cell.value for cell in resumo[1]]
            
            for col in expected_columns:
                if col not in header:
                    raise ValueError(f"Coluna '{col}' não encontrada na aba Resumo")
            
            return True
        except Exception as e:
            logger.error(f"Validação falhou: {e}")
            return False

    def export_to_excel(self):
        """Exporta os resultados para uma planilha Excel formatada"""
        logger.info("Iniciando exportação para Excel...")
        data_inicial, data_final = self._get_datas_referencia()
        output_path = self.settings.RESULTS_DIR / f"CONCILIACAO_{data_inicial.replace('/', '-')}_a_{data_final.replace('/', '-')}.xlsx"
        
        try:
            if not self.conn:
                logger.error("Tentativa de exportação com conexão fechada")
                raise RuntimeError("Conexão com o banco de dados não está aberta")
            
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            
            # 1. Planilha de Resumo
            query_resumo = f"""
                SELECT 
                    codigo_fornecedor as "Código Fornecedor",
                    descricao_fornecedor as "Descrição Fornecedor",
                    saldo_contabil as "Saldo Contábil",
                    saldo_financeiro as "Saldo Financeiro",
                    diferenca as "Diferença",
                    CASE 
                        WHEN diferenca > 0 THEN 'Contábil > Financeiro'
                        WHEN diferenca < 0 THEN 'Financeiro > Contábil'
                        ELSE 'OK'
                    END as "Tipo Diferença",
                    status as "Status",
                    detalhes as "Detalhes"
                FROM 
                    {self.settings.TABLE_RESULTADO}
                ORDER BY 
                    ABS(diferenca) DESC
            """
            df_resumo = pd.read_sql(query_resumo, self.conn)
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
            
            # 2. Planilha de Detalhes Financeiros
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
            
            # 3. Planilha de Detalhes Contábeis
            query_contabil = f"""
                SELECT 
                    conta_contabil as "Conta Contábil",
                    descricao_conta as "Descrição Conta",
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
            
            # 4. Planilha de Contas x Itens 
            query_contas_itens = f"""
                SELECT 
                    r.codigo_fornecedor as "Código Fornecedor",
                    r.descricao_fornecedor as "Descrição Fornecedor",
                    ci.conta_contabil as "Conta Contábil",
                    ci.item as "Item",
                    ci.descricao_item as "Descrição Item",
                    ci.quantidade as "Quantidade",
                    ci.valor_unitario as "Valor Unitário",
                    ci.valor_total as "Valor Total",
                    ci.saldo_atual as "Saldo Atual"
                FROM 
                    {self.settings.TABLE_CONTAS_ITENS} ci
                JOIN 
                    {self.settings.TABLE_RESULTADO} r ON ci.conta_contabil LIKE '%' || r.codigo_fornecedor || '%'
                WHERE 
                    r.status = 'DIVERGENTE'
                ORDER BY 
                    r.codigo_fornecedor, ci.conta_contabil
            """
            df_contas_itens = pd.read_sql(query_contas_itens, self.conn)
            df_contas_itens.to_excel(writer, sheet_name='Contas x Itens', index=False)
            
            writer.close()
            
            workbook = openpyxl.load_workbook(output_path)
            for sheetname in workbook.sheetnames:
                sheet = workbook[sheetname]
                self._apply_styles(sheet)
            
            sheet = workbook['Resumo']
            sheet['I1'] = "Data de Referência:"
            sheet['J1'] = f"{data_inicial} a {data_final}"
            sheet['I2'] = "Observações:"
            sheet['J2'] = "Conciliação automática gerada pelo sistema"
            
            workbook.save(output_path)
            logger.info(f"Planilha de resultados exportada para {output_path}")
            
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
        self._initialize_database()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()