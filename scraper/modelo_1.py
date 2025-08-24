from playwright.sync_api import sync_playwright, TimeoutError 
from config.logger import configure_logger
from config.settings import Settings
from .utils import Utils
from datetime import date
from pathlib import Path


import time

logger = configure_logger()

class Modelo_1(Utils):
    def __init__(self, page):  
        """Inicializa o Modelo 1 com a página do navegador"""
        self.page = page
        self.settings = Settings() 
        self.parametros_json = 'modelo_1'
        self._definir_locators()
        logger.info("Modelo_1 inicializado")

    def _definir_locators(self):
        """Centraliza apenas os locators específicos do Modelo 1"""
        self.locators = {
            # submenu
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),
            'submenu_balancetes': self.page.get_by_text("Balancetes (34)"),
            'opcao_modelo1': self.page.get_by_text("Modelo 1", exact=True),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'botao_marcar_filiais': self.page.get_by_role("button", name="Marca Todos - <F4>"),

            # parametros
            'data_inicial': self.page.locator("#COMP4512").get_by_role("textbox"),
            'data_final': self.page.locator("#COMP4514").get_by_role("textbox"),
            'conta_inicial':self.page.locator("#COMP4516").get_by_role("textbox"),
            'conta_final':self.page.locator("#COMP4518").get_by_role("textbox"),
            'data_lucros_perdas': self.page.locator("#COMP4556").get_by_role("textbox"),
            'grupos_receitas_despesas': self.page.locator("#COMP4562").get_by_role("textbox"),
            'data_sid_art': self.page.locator("#COMP4564").get_by_role("textbox"),
            'num_linha_balancete': self.page.locator("#COMP4566").get_by_role("textbox"),
            'desc_moeda': self.page.locator("#COMP4568").get_by_role("textbox"),
            'selec_filiais': self.page.locator("#COMP4570").get_by_role("combobox"),
            'botao_ok': self.page.locator('button:has-text("Ok")'),

            # Gerar planilha
            'aba_planilha': self.page.get_by_role("button", name="Planilha"),
            'formato': self.page.locator("#COMP4547").get_by_role("combobox"),
            'botao_imprimir': self.page.get_by_role("button", name="Imprimir"),
            'botao_sim': self.page.get_by_role("button", name="Sim")
        }
        logger.info("Seledores definidos")


    def _navegar_menu(self):
        """navegação no menu"""
        try:
            logger.info("Iniciando navegação no menu...")
            # Espera o menu principal estar disponível            
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=10000)
            self.locators['menu_relatorios'].click()
            logger.info("Menu Relatórios clicado")
            time.sleep(1)  
            if not self.locators['submenu_balancetes'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            self.locators['submenu_balancetes'].click()
            logger.info("Submenu Balancetes clicado")
            time.sleep(1)
            self.locators['opcao_modelo1'].wait_for(state="visible")
            self.locators['opcao_modelo1'].click()
            logger.info("Modelo 1 selecionada")
            
        except Exception as e:
            logger.error(f"Falha na navegação do menu: {e}")
            
            raise

    def _preencher_parametros(self):
        try:
            logger.info(f"Usando chave JSON: {self.parametros_json}")
            # Resolver valores dinâmicos
            # input_data_inicial = self._resolver_valor(self.parametros.get('data_inicial'))
            # input_data_final = self._resolver_valor(self.parametros.get('data_final'))
            # input_data_sid_art  = self._resolver_valor(self.parametros.get('data_sid_art'))

            input_data_inicial = self.parametros.get('data_inicial')
            input_data_final =self.parametros.get('data_final')
            input_data_inicial =self.parametros.get('data_sid_art')
            # Obter outros parâmetros
            input_conta_inicial = self.parametros.get('conta_inicial')
            input_conta_final = self.parametros.get('conta_final')
            input_data_lucros_perdas = self.parametros.get('data_lucros_perdas')
            input_grupos_receitas_despesas = self.parametros.get('grupos_receitas_despesas')
            input_data_sid_art = self.parametros.get('data_sid_art')
            input_num_linha_balancete = self.parametros.get('num_linha_balancete')
            input_desc_moeda = self.parametros.get('desc_moeda')

            # parâmetros
            self.locators['data_inicial'].wait_for(state="visible")
            self.locators['data_inicial'].click()
            self.locators['data_inicial'].fill(input_data_inicial)
            time.sleep(0.5) 
            self.locators['data_final'].click()
            self.locators['data_final'].fill(input_data_final)
            time.sleep(0.5) 
            self.locators['conta_inicial'].click()
            self.locators['conta_inicial'].fill(input_conta_inicial)
            time.sleep(0.5) 
            self.locators['conta_final'].click()
            self.locators['conta_final'].fill(input_conta_final)
            time.sleep(0.5) 
            self.locators['data_lucros_perdas'].click()
            self.locators['data_lucros_perdas'].fill(input_data_lucros_perdas)
            time.sleep(0.5) 
            self.locators['grupos_receitas_despesas'].click()
            self.locators['grupos_receitas_despesas'].fill(input_grupos_receitas_despesas)
            time.sleep(0.5) 
            self.locators['data_sid_art'].click()
            self.locators['data_sid_art'].fill(input_data_sid_art)
            time.sleep(0.5) 
            self.locators['num_linha_balancete'].click()
            self.locators['num_linha_balancete'].fill(input_num_linha_balancete)
            time.sleep(0.5) 
            self.locators['desc_moeda'].click()
            self.locators['desc_moeda'].fill(input_desc_moeda)
            time.sleep(0.5)
            self.locators['selec_filiais'].click()
            time.sleep(0.5)
            self.locators['selec_filiais'].select_option("0")
            time.sleep(0.5)
            self.locators['botao_ok'].click()
        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")
            raise


    def _gerar_planilha(self):
        """Gera e baixa a planilha do Modelo 1"""
        try: 
            self.locators['aba_planilha'].wait_for(state="visible")
            time.sleep(1) 
            self.locators['aba_planilha'].click()
            time.sleep(1) 
            
            if not self.locators['formato'].is_visible():
                self.locators['aba_planilha'].click()
                time.sleep(1)
            
            self.locators['formato'].select_option("3")
            time.sleep(1) 
            
            # Esperar pelo download com timeout aumentado
            with self.page.expect_download(timeout=120000) as download_info:
                self.locators['botao_imprimir'].click()
                time.sleep(2)
                self._fechar_popup_se_existir()
                
            
            download = download_info.value
            logger.info(f"Download iniciado: {download.suggested_filename}") 
            
            # Aguardar conclusão do download
            download_path = download.path()
            if download_path:
                settings = Settings()
                destino = Path(settings.CAMINHO_PLS) / settings.PLS_MODELO_1
                destino.parent.mkdir(parents=True, exist_ok=True)
                
                
                download.save_as(destino)
                logger.info(f"Arquivo Modelo 1 salvo em: {destino}")
            else:
                logger.error("Download falhou - caminho não disponível")
            
            
            # Verificar se há botão de confirmação (se necessário)
            if 'botao_sim' in self.locators and self.locators['botao_sim'].is_visible():
                self.locators['botao_sim'].click()
                
        except Exception as e:
            logger.error(f"Falha na geração da planilha: {e}")
            raise

    def execucao(self):
        """Fluxo principal de execução"""
        try:
            
            logger.info('Iniciando execução do Modelo 1')
            self._carregar_parametros('parameters.json', self.parametros_json)
            self._navegar_menu()
            time.sleep(1) 
            self._confirmar_operacao()
            time.sleep(1) 
            self._fechar_popup_se_existir()
            self._preencher_parametros()
            self._selecionar_filiais()
            self._gerar_planilha()
            logger.info("✅ Modelo 1 executado com sucesso")
            return {
                'status': 'success',
                'message': 'Modelo 1 completo'
            }
            
        except Exception as e:
            error_msg = f"❌ Falha na execução: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}