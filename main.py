from scraper.protheus import ProtheusScraper
from config.logger import configure_logger
from config.settings import Settings

logger = configure_logger()

def main():
    custom_settings = Settings()
    custom_settings.HEADLESS = False 
    try:
        with ProtheusScraper(settings=custom_settings) as scraper:
            results = scraper.run() or []  
            success_count = len([r for r in results if r.get('status') == 'success'])
            logger.info(f"Process completed: {success_count}/{len(results)} successful submissions")
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
    return 0

if __name__ == "__main__":
    main()
