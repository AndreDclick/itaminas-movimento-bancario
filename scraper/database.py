import sqlite3
from pathlib import Path
from config.settings import Settings
from config.logger import configure_logger
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import re
from difflib import get_close_matches

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

            # Obter mapeamento de colunas
            column_mapping = self._get_column_mapping(Path(file_path))
            
            # Renomear colunas conforme mapeamento
            df.rename(columns=column_mapping, inplace=True)
            
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
                        # Extrai parcela do número do título
                        df['parcela'] = df['titulo'].str.extract(r'(\d+)$').fillna('1')
                        logger.warning("Coluna 'parcela' criada a partir do título")
                        remaining_missing.remove('parcela')
                    
                    if 'conta_contabil' in remaining_missing:
                        # Define um valor padrão para conta contábil
                        df['conta_contabil'] = 'CONTA_PADRAO' 
                        logger.warning("Coluna 'conta_contabil' preenchida com valor padrão")
                        remaining_missing.remove('conta_contabil')
                    
                    if remaining_missing:
                        logger.error(f"Colunas obrigatórias ausentes após tratamento: {remaining_missing}")
                        raise ValueError(f"Colunas obrigatórias ausentes: {remaining_missing}")

            # Limpa dados e seleciona apenas as colunas esperadas
            df = df[expected_columns]
            df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
            logger.info(f"DataFrame limpo - shape final: {df.shape}")

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
            return  [
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
                # 1. Tratamento de datas
                date_cols = ['data_emissao', 'data_vencimento']
                for col in date_cols:
                    if col in df.columns:
                        try:
                            # Tenta converter de vários formatos
                            df[col] = pd.to_datetime(
                                df[col], 
                                errors='coerce',
                                format='mixed',
                                dayfirst=True
                            )
                        except Exception as e:
                            logger.warning(f"Erro ao converter datas na coluna {col}: {str(e)}")
                            df[col] = pd.NaT

                # 2. Tratamento de valores numéricos
                num_cols = ['valor_original', 'saldo_devedor']
                for col in num_cols:
                    if col in df.columns:
                        df[col] = (
                            df[col].astype(str)
                            .str.replace(r'[^\d,-]', '', regex=True) 
                            .str.replace(',', '.')  # Padroniza decimal
                            .replace('', np.nan)    # Strings vazias para NaN
                            .astype(float)          # Converte para float
                            .fillna(0)              # Preenche NaN com 0
                        )
                
                # 3. Tratamento de texto
                text_cols = ['fornecedor', 'titulo', 'tipo_titulo']
                for col in text_cols:
                    if col in df.columns:
                        df[col] = (
                            df[col].astype(str)
                            .str.strip()
                            .str.replace(r'\s+', ' ', regex=True)  
                            .replace('nan', '')  
                        )
                
                # 4. Filtrar registros indesejados
                if 'tipo_titulo' in df.columns:
                    df = df[~df['tipo_titulo'].isin(self.settings.FORNECEDORES_EXCLUIR)]
                
                # 5. Criar coluna de parcela se necessário
                if 'titulo' in df.columns and 'parcela' not in df.columns:
                    df['parcela'] = df['titulo'].str.extract(r'(\d+)$').fillna('1')

            elif sheet_type == 'modelo1':
                # Tratamento específico para o modelo 1
                num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
                for col in num_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(
                            df[col].astype(str)
                            .str.replace(r'[^\d,-]', '', regex=True)
                            .str.replace(',', '.'),
                            errors='coerce'
                        ).fillna(0)
                
                # Criar coluna tipo_fornecedor se não existir
                if 'tipo_fornecedor' not in df.columns and 'descricao_conta' in df.columns:
                    df['tipo_fornecedor'] = df['descricao_conta'].apply(
                        lambda x: 'FORNECEDOR' if 'FORNEC' in str(x).upper() else 'OUTROS'
                    )

            elif sheet_type == 'contas_itens':
            # Tratamento específico para contas x itens
                num_cols = ['saldo_anterior', 'debito', 'credito', 'saldo_atual']
                for col in num_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(
                            df[col].astype(str)
                            .str.replace(r'[^\d,-]', '', regex=True)
                            .str.replace(',', '.'),
                            errors='coerce'
                        ).fillna(0)
                
                # Criar colunas ausentes com valores padrão
                if 'item' not in df.columns:
                    df['item'] = df.get('descricao_item', '').str[:50] 
                    
                if 'quantidade' not in df.columns:
                    df['quantidade'] = 1
                    
                if 'valor_unitario' not in df.columns and 'saldo_atual' in df.columns:
                    df['valor_unitario'] = df['saldo_atual']
                    
                if 'valor_total' not in df.columns and 'saldo_atual' in df.columns:
                    df['valor_total'] = df['saldo_atual']
            
            df = df.drop_duplicates()

            # Log final
            logger.info(f"DataFrame limpo - shape final: {df.shape}")
            
            return df
            
        except Exception as e:
            logger.error(f"Erro na limpeza dos dados ({sheet_type}): {str(e)}", exc_info=True)
            raise
    
    def _get_column_mapping(self, file_path: Path):
        """Retorna o mapeamento de colunas apropriado com base no nome do arquivo"""
        filename = file_path.stem.lower()
        
        if 'finr' in filename:
            return {
                'fornecedor': 'Codigo-Nome do Fornecedor',
                'titulo': 'Prf-Numero Parcela',
                'tipo_titulo': 'Tp',
                'data_emissao': 'Data de Emissao',
                'data_vencimento': 'Data de Vencto',
                'vencto_real': 'Vencto Real',
                'valor_original': 'Valor Original',
                'saldo_devedor': 'Tit Vencidos Valor nominal',
                'situacao': 'Natureza',
                'conta_contabil': 'Natureza',
                'centro_custo': 'Porta- dor'
            }
        elif 'ctbr140' in filename:
            return {
                'conta_contabil': 'Codigo',
                'descricao_conta': 'Descricao',
                'codigo_fornecedor': 'Codigo.1',
                'descricao_fornecedor': 'Descricao.1',
                'saldo_anterior': 'Saldo anterior',
                'debito': 'Debito',
                'credito': 'Credito',
                'movimento_periodo': 'Movimento do periodo',
                'saldo_atual': 'Saldo atual',
                'tipo_fornecedor': 'Descricao.1'  
            }
        elif 'ctbr040' in filename:
            return {
                'conta_contabil': 'Conta',
                'descricao_item': 'Descricao',
                'saldo_anterior': 'Saldo anterior',
                'debito': 'Debito',
                'credito': 'Credito',
                'saldo_atual': 'Saldo atual',                
                'item': 'Descricao',  
                'quantidade': '1',     # Valor padrão
                'valor_unitario': 'Saldo atual',  
                'valor_total': 'Saldo atual'      
        }
        else:
            raise ValueError(f"Tipo de planilha não reconhecido: {file_path.name}")
    
    def process_data(self):
        """Processa os dados e gera a conciliação conforme especificado"""
        try:
            self.conn.execute("BEGIN TRANSACTION")
            cursor = self.conn.cursor()
            
            # Limpar tabela de resultados antes de processar
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
                        ELSE 'DIVERGENTE'
                    END
            """
            cursor.execute(query_diferenca)
            
            # 4. Inserir fornecedores contábeis que não estão no financeiro (QUERY CORRIGIDA)
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
        # Definir estilos
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        align_center = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), 
                            right=Side(style='thin'), 
                            top=Side(style='thin'), 
                            bottom=Side(style='thin'))
        
        # Aplicar aos cabeçalhos
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = align_center
            cell.border = thin_border
        
        # Aplicar bordas a todas as células
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = thin_border
                if cell.column_letter in ['K', 'L', 'M', 'N']:  # Colunas numéricas
                    cell.number_format = '#,##0.00'
        
        # Ajustar largura das colunas
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

    def export_to_excel(self):
        """Exporta os resultados para uma planilha Excel formatada na pasta results"""
        logger.info("Iniciando exportação para Excel...")
        try:
            if not self.conn:
                logger.error("Tentativa de exportação com conexão fechada")
                raise RuntimeError("Conexão com o banco de dados não está aberta")
            
            output_path = self.settings.RESULTS_DIR / f"CONCILIACAO_{self.settings.DATA_REFERENCIA.replace('/', '-')}.xlsx"
            
            # Criar um novo arquivo Excel
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            
            # 1. Planilha de Resumo
            query_resumo = f"""
                SELECT 
                    codigo_fornecedor as "Código Fornecedor",
                    descricao_fornecedor as "Descrição Fornecedor",
                    saldo_contabil as "Saldo Contábil",
                    saldo_financeiro as "Saldo Financeiro",
                    diferenca as "Diferença",
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
            
            # Salvar o arquivo Excel
            writer.close()
            
            # Aplicar formatação a todas as planilhas
            workbook = openpyxl.load_workbook(output_path)
            for sheetname in workbook.sheetnames:
                sheet = workbook[sheetname]
                self._apply_styles(sheet)
            
            # Adicionar informações adicionais
            sheet = workbook['Resumo']
            sheet['I1'] = "Data de Referência:"
            sheet['J1'] = self.settings.DATA_REFERENCIA
            sheet['I2'] = "Observações:"
            sheet['J2'] = "Conciliação automática gerada pelo sistema"
            
            workbook.save(output_path)
            logger.info(f"Planilha de resultados exportada para {output_path}")
            
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
