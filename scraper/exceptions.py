class RPAChallengeException(Exception): pass

class BrowserClosedError(RPAChallengeException):
    def __init__(self, message="Navegador fechado durante a operação"):
        self.code = 1001
        super().__init__(message)

class DownloadFailed(RPAChallengeException):
    def __init__(self, message="Falha ao baixar arquivo do desafio", url=None):
        self.code = 1002
        self.url = url
        super().__init__(message)

class FormSubmitFailed(RPAChallengeException):
    def __init__(self, message="Falha no envio do formulário", field=None):
        self.code = 1003
        self.field = field
        super().__init__(message)

class InvalidDataFormat(RPAChallengeException):
    def __init__(self, message="Formato inválido nos dados", details=None):
        self.code = 1004
        self.details = details
        super().__init__(message)

class ResultsSaveError(RPAChallengeException):
    def __init__(self, message="Falha ao salvar resultados", path=None):
        self.code = 1005
        self.path = path
        super().__init__(message)

class TimeoutOperacional(RPAChallengeException):
    def __init__(self, message="Timeout na operação", operation=None):
        self.code = 1006
        self.operation = operation
        super().__init__(message)