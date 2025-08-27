class Exceptions(Exception):
    """Classe base para todas as exceções do projeto RPA"""
    pass

class PlanilhaFormatacaoErradaError(Exceptions):
    def __init__(self, message="Planilha com formatação errada", caminho_arquivo=None):
        self.code = "FE1"
        self.caminho_arquivo = caminho_arquivo
        super().__init__(message)

class LoginProtheusError(Exceptions):
    def __init__(self, message="Falha de Login em sistema Protheus", usuario=None):
        self.code = "FE2"
        self.usuario = usuario
        super().__init__(message)

class ExcecaoNaoMapeadaError(Exceptions):
    def __init__(self, message="Exceção não mapeada", detalhes=None):
        self.code = "FE3"
        self.detalhes = detalhes
        super().__init__(message)

class ExtracaoRelatorioError(Exceptions):
    def __init__(self, message="Falha ao extrair relatório do Protheus", relatorio=None):
        self.code = "FE4"
        self.relatorio = relatorio
        super().__init__(message)

class BrowserClosedError(Exceptions):
    def __init__(self, message="Navegador fechado durante a operação"):
        self.code = 1001
        super().__init__(message)

class DownloadFailed(Exceptions):
    def __init__(self, message="Falha ao baixar arquivo", url=None, caminho_destino=None):
        self.code = 1002
        self.url = url
        self.caminho_destino = caminho_destino
        super().__init__(message)

class FormSubmitFailed(Exceptions):
    def __init__(self, message="Falha no envio do formulário", campo=None, valor=None):
        self.code = 1003
        self.campo = campo
        self.valor = valor
        super().__init__(message)

class InvalidDataFormat(Exceptions):
    def __init__(self, message="Formato inválido nos dados", detalhes=None, tipo_dado=None):
        self.code = 1004
        self.detalhes = detalhes
        self.tipo_dado = tipo_dado
        super().__init__(message)

class ResultsSaveError(Exceptions):
    def __init__(self, message="Falha ao salvar resultados", caminho=None, dados=None):
        self.code = 1005
        self.caminho = caminho
        self.dados = dados
        super().__init__(message)

class TimeoutOperacional(Exceptions):
    def __init__(self, message="Timeout na operação", operacao=None, tempo_limite=None):
        self.code = 1006
        self.operacao = operacao
        self.tempo_limite = tempo_limite
        super().__init__(message)

# Exceções específicas do processo de conciliação
class DiferencaValoresEncontrada(Exceptions):
    def __init__(self, message="Diferença de valores encontrada na conciliação", 
                valor_financeiro=None, valor_contabil=None, fornecedor=None):
        self.code = "CONC001"
        self.valor_financeiro = valor_financeiro
        self.valor_contabil = valor_contabil
        self.fornecedor = fornecedor
        super().__init__(message)

class DataInvalidaConciliação(Exceptions):
    def __init__(self, message="Data inválida para conciliação", data_informada=None):
        self.code = "CONC002"
        self.data_informada = data_informada
        super().__init__(message)

class FornecedorNaoEncontrado(Exceptions):
    def __init__(self, message="Fornecedor não encontrado nos relatórios", 
                codigo_fornecedor=None, nome_fornecedor=None):
        self.code = "CONC003"
        self.codigo_fornecedor = codigo_fornecedor
        self.nome_fornecedor = nome_fornecedor
        super().__init__(message)