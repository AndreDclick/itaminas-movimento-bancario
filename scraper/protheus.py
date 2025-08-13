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
        results = []

        # Definir todas as etapas web que podem ser reiniciadas
        etapas_web = [
            ("financeiro", ExtracaoFinanceiro),
            ("modelo_1", Modelo_1),
            ("contas_x_itens", Contas_x_itens)
        ]

        # 0. Inicialização e login
        try:
            self.start_scraper()
            self.login()
            results.append({
                'status': 'success',
                'message': 'Login realizado com sucesso',
                'etapa': 'autenticação'
            })
        except Exception as e:
            error_msg = f"Falha no login: {str(e)}"
            logger.error(error_msg)
            results.append({
                'status': 'error',
                'message': error_msg,
                'etapa': 'autenticação'
            })
            return results  

        # Executa as etapas web com reinício em caso de falha
        for i, (nome_etapa, classe_etapa) in enumerate(etapas_web):
            tentativas = 0
            max_tentativas = 2  # Tentar no máximo 2 vezes cada etapa
            
            while tentativas < max_tentativas:
                try:
                    # Criar nova instância da classe para cada tentativa
                    etapa = classe_etapa(self.page)
                    resultado = etapa.execucao()
                    results.append(resultado)
                    break  # Sai do while se bem sucedido
                except Exception as e:
                    tentativas += 1
                    error_msg = f"Tentativa {tentativas} na etapa {nome_etapa} falhou: {str(e)}"
                    logger.error(error_msg)
                    
                    if tentativas >= max_tentativas:
                        results.append({
                            'status': 'error',
                            'message': error_msg,
                            'etapa': nome_etapa
                        })
                        
                        # Se não for a última etapa, recarrega a página e faz login novamente
                        if i < len(etapas_web) - 1:
                            logger.info("Recarregando página e fazendo login novamente para próxima etapa...")
                            try:
                                self.page.reload()
                                time.sleep(3)  # Espera a página recarregar
                                self.login()
                            except Exception as e:
                                logger.error(f"Falha ao recarregar página: {e}")
                                break  # Se não conseguir recarregar, para o loop
                    else:
                        logger.info(f"Tentando novamente a etapa {nome_etapa}...")

        # 4. Processamento dos dados (se falhar, interrompe tudo)
        # try:
        #     with DatabaseManager() as db:
        #         db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_FINANCEIRO, 
        #                             self.settings.TABLE_FINANCEIRO)
        #         db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_MODELO, 
        #                             self.settings.TABLE_MODELO1)
        #         db.import_from_excel(self.settings.DATA_DIR / self.settings.PLS_CONTAS, 
        #                             self.settings.TABLE_CONTAS_ITENS)
                
        #         db.process_data()
        #         output_path = db.export_to_excel()
        #         if output_path:
        #             results.append({
        #                 'status': 'success',
        #                 'message': f'Conciliação gerada em {output_path}',
        #                 'etapa': 'processamento'
        #             })
        # except Exception as e:
        #     error_msg = f"Falha no processamento de dados: {str(e)}"
        #     logger.error(error_msg)
        #     results.append({
        #         'status': 'error',
        #         'message': error_msg,
        #         'etapa': 'processamento'
        #     })
        #     return results  # Interrompe execução

        # 5. Verificação final
        if any(r['status'] == 'error' for r in results):
            logger.warning("Processo concluído com erros parciais")
        else:
            logger.info("Processo concluído com sucesso total")

        return results