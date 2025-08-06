from pathlib import Path
from config.settings import Settings

def check_data_files():
    settings = Settings()
    if not settings.DATA_DIR.exists():
        return False
    if not settings.DOWNLOAD_PATH.exists():
        return False
    if settings.DOWNLOAD_PATH.stat().st_size == 0:
        return False
    return True