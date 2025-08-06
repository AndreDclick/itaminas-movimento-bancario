import logging
import os
from datetime import datetime
from pathlib import Path
from config.settings import Settings

def configure_logger():
    settings = Settings()
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_file = settings.LOGS_DIR / f"{timestamp}.log"

        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(formatter)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger
