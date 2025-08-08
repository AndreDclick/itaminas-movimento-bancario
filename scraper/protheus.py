from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import FormSubmitFailed
from .modelo_1 import Modelo_1
import time

logger = configure_logger()

class ProtheusScraper:
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
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),

            # submenu
            'submenu_balancetes': self.page.get_by_text("Balancetes (34)"),
            'opcao_modelo1': self.page.get_by_text("Modelo 1", exact=True),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),

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

    def _fechar_popup_se_existir(self):
        """Método reutilizável para fechar popups"""
        try:
            time.sleep(3)
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
        except Exception as e:
            logger.warning(f" Erro ao verificar popup: {e}")

    def start_scraper(self):
        """Inicia o navegador e página inicial"""
        try:
            logger.info(f"Navegando para: {self.settings.BASE_URL}")
            self.page.goto(self.settings.BASE_URL)
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
            
            time.sleep(3)
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

    def _navegar_menu(self):
        """navegação no menu"""
        try:
            logger.info("Iniciando navegação no menu...")
            
            # Espera o menu principal estar disponível
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=5000)
            self.locators['menu_relatorios'].click()
            logger.info("Menu Relatórios clicado")
            
            time.sleep(2)  
            
            self.locators['submenu_balancetes'].wait_for(state="visible")
            self.locators['submenu_balancetes'].click()
            logger.info("Submenu Balancetes clicado")
            
            time.sleep(2)
            
            self.locators['opcao_modelo1'].wait_for(state="visible")
            self.locators['opcao_modelo1'].click()
            logger.info("Modelo 1 selecionada")
            
        except Exception as e:
            logger.error(f"Falha na navegação do menu: {e}")
            
            raise
    
    def _confirmar_operacao(self):
        """confirmação da operação"""
        try:
            time.sleep(2)
            self.locators['botao_confirmar'].click()
            logger.info("operação confirmada")
            time.sleep(5)
            self._fechar_popup_se_existir()
            
        except Exception as e:
            logger.error(f"Falha na confirmação: {e}")
            raise
    def run(self):
        """Fluxo principal de execução"""
        results = []
        try:
            # 1. Inicialização e login
            self.start_scraper()
            self.login()
        
            self._navegar_menu()
            self._confirmar_operacao()
            results.append({
                'status': 'success',
                'message': 'Login realizado com sucesso',
                'etapa': 'autenticação'
                })

            # 2. Execução do Modelo 1
            modelo_1 = Modelo_1(self.page)
            resultado_modelo = modelo_1.execucao()
            results.append(resultado_modelo)

            # 3. Verificação final
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