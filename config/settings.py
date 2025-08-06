import os
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

class Settings:
    BASE_DIR = Path(__file__).resolve().parent.parent

    # Carrega o .env
    load_dotenv(BASE_DIR / ".env")

    # Dados sensíveis
    EMAIL = os.getenv("EMAIL")
    SENHA = os.getenv("SENHA")
    BASE_URL = os.getenv("BASE_URL")

    
    # Paths
    DATA_DIR = BASE_DIR / "data"
    LOGS_DIR = BASE_DIR / "logs"
    RESULTS_DIR = BASE_DIR / "results"
    
    # Files
    DOWNLOAD_PATH = DATA_DIR 
    RESULTS_PATH = RESULTS_DIR 
    
    DATA_BASE = datetime.now().strftime("%d/%m/%Y")

    # Timeouts
    TIMEOUT = 10000  
    DELAY = 0.5  
    SHUTDOWN_DELAY = 3 
    
    # Browser
    HEADLESS = False

    def __init__(self):
        os.makedirs(self.DATA_DIR, exist_ok=True)
        os.makedirs(self.LOGS_DIR, exist_ok=True)
        os.makedirs(self.RESULTS_DIR, exist_ok=True)
