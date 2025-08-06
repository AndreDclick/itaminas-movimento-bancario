from scraper.protheus import ProtheusScraper
from config.logger import configure_logger

logger = configure_logger()

def main():
    try:
        with ProtheusScraper() as scraper:
            results = scraper.run() or []  # Garante que results nunca será None
            success_count = len([r for r in results if r.get('status') == 'success'])
            logger.info(f"Process completed: {success_count}/{len(results)} successful submissions")
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)  # Adiciona traceback completo
        return 1
    return 0

if __name__ == "__main__":
    print("Methods in ProtheusScraper:", dir(ProtheusScraper))
    main()