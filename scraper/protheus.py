from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import FormSubmitFailed
from .utils import UtilsScraper
from .modelo_1 import Modelo_1
from .financeiro import ExtracaoFinanceiro
from .contasxitens import Contas_x_itens
# from .database import DatabaseManager
import time

logger = configure_logger()

class ProtheusScraper(UtilsScraper):
    def __init__(self, settings=Settings()):
        self.settings = settings
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self._initialize_resources()
        logger.info("Navegador inicializado")

    def _initialize_resources(self):
        """Inicializa todos os recursos do Playwright"""
        self.playwright = sync_playwright().start()
        self._setup_browser()
        self._setup_page()
        self._definir_locators()

    def _setup_browser(self):
        """Configura o navegador Edge"""
        self.browser = self.playwright.chromium.launch(
            headless=self.settings.HEADLESS,
            args=["--start-maximized"],
            channel="msedge"
        )

    def _setup_page(self):
        """Configura a página e contexto"""
        self.context = self.browser.new_context(no_viewport=True)
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.settings.TIMEOUT)

    def _definir_locators(self):
        """Centraliza todos os locators como variáveis"""
        self.locators = {
            'iframe': self.page.locator("iframe"),
            'botao_ok': self.page.locator('button:has-text("Ok")'),
            'campo_usuario': self.page.frame_locator("iframe").get_by_placeholder("Ex. sp01\\nome.sobrenome"),
            'campo_senha': self.page.frame_locator("iframe").get_by_label("Insira sua senha"),
            'botao_entrar': self.page.frame_locator("iframe").get_by_role("button", name="Entrar"),            
            'campo_grupo': self.page.frame_locator("iframe").get_by_label("Grupo"),
            'campo_filial': self.page.frame_locator("iframe").get_by_label("Filial"),
            'campo_ambiente': self.page.frame_locator("iframe").get_by_label("Ambiente"),
            'popup_fechar': self.page.get_by_role("button", name="Fechar")
        }

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._fechar_recursos()
        logger.info("Navegador fechado")

    def _fechar_recursos(self):
        """Fecha todos os recursos de forma segura"""
        time.sleep(self.settings.SHUTDOWN_DELAY)
        self.context.close()
        self.browser.close()
        self.playwright.stop()

    def start_scraper(self):
        """Inicia o navegador e página inicial"""
        try:
            logger.info(f"Navegando para: {self.settings.BASE_URL}")
            self.page.goto(self.settings.BASE_URL)
            if self.locators['botao_ok'].is_visible():
                self.locators['botao_ok'].click()
                logger.info("Botão 'Ok' clicado")
            
        except Exception as e:
            logger.error(f"Falha ao iniciar scraper: {e}")
            raise FormSubmitFailed(f"Erro inicial: {e}")

    def login(self):
        """Realiza o login no sistema"""
        try:
            logger.info("Iniciando login...")
            # Interação com elementos
            self.locators['iframe'].wait_for(state="visible")
            self.locators['campo_usuario'].fill(self.settings.USUARIO)
            self.locators['campo_senha'].fill(self.settings.SENHA)
            self.locators['botao_entrar'].click()
            
            input_campo_grupo = '01'
            input_campo_filial = '0101'
            input_campo_ambiente = '34'
            self.locators['campo_grupo'].wait_for(state="visible")
            self.locators['campo_grupo'].click()
            self.locators['campo_grupo'].fill(input_campo_grupo)
            self.locators['campo_filial'].click()
            self.locators['campo_filial'].fill(input_campo_filial)
            self.locators['campo_ambiente'].click()
            self.locators['campo_ambiente'].fill(input_campo_ambiente)
            self.locators['botao_entrar'].click()
            self._fechar_popup_se_existir()
            time.sleep(3)
            logger.info("Login realizado com sucesso")
            
        except PlaywrightTimeoutError:
            logger.error("Tempo esgotado ao esperar elementos de login")
            raise
        except Exception as e:
            logger.error(f"Falha no login: {str(e)}")
            raise FormSubmitFailed(f"Erro de login: {e}")


    def run(self):
        """Fluxo principal de execução"""
        results = []
        try:
            # 0. Inicialização e login
            self.start_scraper()
            self.login()
            results.append({
                'status': 'success',
                'message': 'Login realizado com sucesso',
                'etapa': 'autenticação'
            })

            # 1. Financeiro            
            # financeiro = ExtracaoFinanceiro(self.page)
            # resultado_financeiro = financeiro.execucao()
            # results.append(resultado_financeiro)

            # # 2. Execução do Modelo 1
            # modelo_1 = Modelo_1(self.page)
            # resultado_modelo = modelo_1.execucao()
            # results.append(resultado_modelo)

            # 3. Execução do Contas x Itens
            contasxitens = Contas_x_itens(self.page)
            resultado_contas = contasxitens.execucao()
            results.append(resultado_contas)

            # 4. Processamento dos dados
            # with DatabaseManager() as db:
            #     # Importa as planilhas baixadas
            #     db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_FINANCEIRO, 
            #                     self.settings.TABLE_FINANCEIRO)
            #     db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_MODELO, 
            #                     self.settings.TABLE_MODELO1)
            #     db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_CONTAS, 
            #                     self.settings.TABLE_CONTAS_ITENS)
                
            #     # Processa os dados
            #     db.process_data()
                
            #     # Exporta os resultados
            #     output_path = db.export_to_excel()
            #     if output_path:
            #         results.append({
            #             'status': 'success',
            #             'message': f'Conciliação gerada em {output_path}',
            #             'etapa': 'processamento'
            #         })

            # 5. Verificação final
            if any(r['status'] == 'error' for r in results):
                logger.warning("Processo concluído com erros parciais")
            else:
                logger.info("Processo concluído com sucesso total")

        except Exception as e:
            error_msg = f"Falha: {str(e)}"
            logger.error(error_msg)
            results.append({
                'status': 'error',
                'message': error_msg,
                'etapa': 'execução geral'
            })
        
        return results