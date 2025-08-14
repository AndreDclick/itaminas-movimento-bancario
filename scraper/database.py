import sqlite3
from pathlib import Path
from config.settings import Settings
from config.logger import configure_logger
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

logger = configure_logger()

class DatabaseManager:
    def __init__(self):
        self.settings = Settings()
        self.conn = None
        self._initialize_database()

    def _initialize_database(self):
        """Cria o banco de dados e as tabelas se não existirem"""
        try:
            self.conn = sqlite3.connect(self.settings.DB_PATH)
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
                CREATE TABLE IF NOT EXISTS {self.settings.TABLE_CONTAS_ITENS} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    conta_contabil TEXT,
                    item TEXT,
                    descricao_item TEXT,
                    quantidade REAL DEFAULT 0,
                    valor_unitario REAL DEFAULT 0,
                    valor_total REAL DEFAULT 0,
                    saldo REAL DEFAULT 0,
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

    def _clean_dataframe(self, df, sheet_type):
        """Realiza a limpeza dos dados conforme o tipo de planilha"""
        try:
            # Converter todas as colunas para string primeiro para evitar problemas
            df = df.astype(str)
            
            # Remover linhas totalmente vazias
            df = df.replace('nan', np.nan)
            df = df.dropna(how='all')
            
            # Processamento específico para cada tipo de planilha
            if sheet_type == 'financeiro':
                # Separar Prf-NumeroParcela em fornecedor e parcela
                df[['fornecedor', 'parcela']] = df['Prf-NumeroParcela'].str.split('-', n=1, expand=True)
                
                # Filtrar registros indesejados
                df = df[~df['Tp'].isin(['NDF', 'PA'])]
                
                # Converter valores vazios para 0
                num_cols = ['Valor Original', 'Titulos a vencerValor nominal']
                for col in num_cols:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            elif sheet_type == 'modelo1':
                # Filtrar apenas contas de fornecedores
                df = df[df['descricao_conta'].str.contains('FORNEC', case=False, na=False)]
                
                # Classificar tipo de fornecedor
                df['tipo_fornecedor'] = np.where(
                    df['descricao_conta'].str.contains('NACIONAL', case=False, na=False),
                    'Fornecedores Nacionais',
                    'Fornecedores - Outros'
                )
            
            elif sheet_type == 'contas_itens':
                # Filtrar contas específicas de fornecedores
                contas_fornecedores = ['10106020001', '20102010001']
                df = df[df['conta_contabil'].isin(contas_fornecedores)]
                
                # Garantir que valores numéricos estejam corretos
                num_cols = ['quantidade', 'valor_unitario', 'valor_total', 'saldo']
                for col in num_cols:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            return df
        except Exception as e:
            logger.error(f"Erro na limpeza dos dados ({sheet_type}): {e}")
            raise

    def import_from_excel(self, file_path: Path, table_name: str):
        """Importa dados de uma planilha Excel para a tabela especificada"""
        try:
            # Determinar o tipo de planilha
            sheet_type = None
            if 'finr' in file_path.stem.lower():
                sheet_type = 'financeiro'
                col_map = self.settings.COLUNAS_FINANCEIRO
            elif 'ctbr140' in file_path.stem.lower():
                sheet_type = 'modelo1'
                col_map = self.settings.COLUNAS_MODELO1
            elif 'ctbr040' in file_path.stem.lower():
                sheet_type = 'contas_itens'
                col_map = self.settings.COLUNAS_CONTAS_ITENS
            
            if not sheet_type:
                raise ValueError(f"Tipo de planilha não reconhecido: {file_path.name}")
            
            # Ler o arquivo Excel
            df = pd.read_excel(file_path, header=0)
                
            # Renomear colunas conforme mapeamento
            df = df.rename(columns={v: k for k, v in col_map.items()})
            
            # Manter apenas as colunas necessárias
            df = df[list(col_map.keys())]
            
            # Limpar os dados
            df = self._clean_dataframe(df, sheet_type)
            
            # Salvar no banco de dados
            df.to_sql(table_name, self.conn, if_exists='append', index=False)
            logger.info(f"Dados importados de {file_path.name} para tabela {table_name}")
            
            return True
        except Exception as e:
            logger.error(f"Erro ao importar dados de {file_path.name}: {e}")
            return False

    def process_data(self):
        """Processa os dados e gera a conciliação conforme especificado"""
        try:
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
                        WHERE descricao_fornecedor LIKE '%' || SUBSTR(descricao_conta, 1, INSTR(descricao_conta, ' ') - 1 || '%'
                    )
            """
            cursor.execute(query_fornecedores_contabeis)
            
            self.conn.commit()
            logger.info("Processamento de dados concluído com sucesso")
            
            return True
        except Exception as e:
            logger.error(f"Erro ao processar dados: {e}")
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
        try:
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
            
            # 4. Planilha de Contas x Itens (apenas para divergências)
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
                    ci.saldo as "Saldo"
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
            self.conn.close()
            logger.info("Conexão com o banco de dados fechada")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()