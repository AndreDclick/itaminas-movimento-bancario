from playwright.sync_api import sync_playwright
import openpyxl
from openpyxl import Workbook
import time
from pathlib import Path
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import (BrowserClosedError, DownloadFailed,
                        FormSubmitFailed, InvalidDataFormat, ResultsSaveError)

logger = configure_logger()

class ProtheusScraper:
    def __init__(self, settings=Settings()):
        self.settings = settings
        self.playwright = sync_playwright().start()
        self._initialize_browser()
        logger.info("Browser initialized")

    def _initialize_browser(self):
        self.browser = self.playwright.chromium.launch(
            headless=self.settings.HEADLESS,
            args=["--start-maximized"],
            channel="msedge"
        )
        self.context = self.browser.new_context(no_viewport=True)
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.settings.TIMEOUT)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        time.sleep(self.settings.SHUTDOWN_DELAY)
        self.context.close()
        self.browser.close()
        self.playwright.stop()
        logger.info("Browser closed")

    def start_scraper(self):
        try:
            self.page.goto(self.settings.BASE_URL)
            self.page.locator("button:text('Start')").click()
            logger.info("scraper started")
        except Exception as e:
            logger.error(f"Failed to start scraper: {e}")
            raise FormSubmitFailed(f"Failed to start scraper: {e}")

    def _check_completion(self):
        try:
            return self.page.locator("div.congratulations").is_visible(timeout=1000)
        except:
            return False

    

    def run(self):
        self.start_scraper()
        return self.fill_form(self._load_data())