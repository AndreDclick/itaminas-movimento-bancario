import logging
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional
from .settings import Settings

class CustomLogger:
    """Logger personalizado com configurações alinhadas à documentação do RPA"""
    
    _instance: Optional[logging.Logger] = None
    
    @classmethod
    def get_logger(cls, name: Optional[str] = None) -> logging.Logger:
        """Retorna uma instância do logger configurado (singleton)"""
        if cls._instance is None:
            cls._instance = cls._configure_logger(name or __name__)
        return cls._instance
    
    @staticmethod
    def _configure_logger(name: str) -> logging.Logger:
        """Configura o logger conforme especificado na documentação"""
        settings = Settings()
        logger = logging.getLogger(name)
        
        # Define nível base como INFO, mas permite DEBUG se necessário
        logger.setLevel(logging.INFO)
        
        # Remove handlers existentes para evitar duplicação
        if logger.handlers:
            for handler in logger.handlers[:]:
                logger.removeHandler(handler)
        
        # Formato alinhado com a documentação: Data/Hora + Passo + Status + Mensagem
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - [Paso:%(filename)s:%(lineno)d] - %(message)s'
        )
        
        # Cria diretório de logs se não existir
        settings.LOGS_DIR.mkdir(exist_ok=True)
        
        # File handler com timestamp
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_file = settings.LOGS_DIR / f"conciliação_fornecedores_{timestamp}.log"
        
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)
        
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
        
        return logger

# Função de conveniência para manter compatibilidade
def configure_logger():
    """Mantém compatibilidade com código existente"""
    return CustomLogger.get_logger()