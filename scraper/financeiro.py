from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.logger import configure_logger
import time

logger = configure_logger()

class ExtracaoFinanceiro:
    def __init__(self, page):
        """Inicializa a automação de extração de planilha financeira"""
        self.page = page
        self._definir_locators()
        logger.info("ExtracaoFinanceiro inicializada")

    def _definir_locators(self):
        """Centraliza os locators específicos da extração financeira"""
        self.locators = {
            # Navegação
            'menu_relatorios': self.page.get_by_text("Relatorios (9)"),
            'menu_financeiro': self.page.get_by_text("Financeiro (2)"),
            'menu_titulos_a_pagar': self.page.get_by_text("Títulos a Pagar", exact=True),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'popup_fechar': self.page.get_by_role("button", name="Fechar"),
            
            
            # Janela "Posição dos Títulos a Pagar"
            
            'planilha': self.page.get_by_role("button", name="Planilha"),
            'tipo_de_planilha': self.page.locator('#COMP4547').get_by_role('combobox'),
            'outras_acoes': self.page.get_by_role('button', name='Outras Ações'),
            'parametros_menu': self.page.get_by_text('Parâmetros'),
            'imprimir_btn': self.page.get_by_role('button', name='Imprimir'),

            # Janela de Parâmetros
            'do_vencimento_campo': self.page.get_by_label('Do Vencimento?'),
            'ate_o_vencimento_campo': self.page.get_by_label('Até o Vencimento?'),
            'da_emissao_campo': self.page.get_by_label('Da Emissao?'),
            'ate_emissao_campo': self.page.get_by_label('Ate a Emissão?'),
            'da_data_contabil_campo': self.page.get_by_label('Da data contabil?'),
            'ate_data_contabil_campo': self.page.get_by_label('Até data contabil?'),
            'data_base_campo': self.page.get_by_label('Data Base?'),
            'ok_btn': self.page.get_by_role('button', name='OK'),
            
            # Janela de Seleção de Filiais
            'selecao_filiais_janela': self.page.get_by_text('Seleção de filiais'),
            'matriz_filial_checkbox': self.page.get_by_text('Matriz e Filial'), # Se houver checkbox para isso
            'marcar_todos_btn': self.page.get_by_role('button', name='Marca Todos- <F4>'),
            'confirmar_btn': self.page.get_by_role('button', name='Confirmar')
        }
        logger.info("Seletores definidos")

    def _navegar_e_configurar_planilha(self):
        """Navega para a tela de Títulos a Pagar e configura a extração para planilha."""
        try:
            logger.info("Iniciando navegação no menu...")
            
            # Espera o menu principal estar disponível
            time.sleep(2)
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=10000)
            self.locators['menu_relatorios'].click()
            logger.info("Menu Relatórios clicado")
            
            time.sleep(2)  
            if not self.locators['menu_financeiro'].is_visible():
                self.locators['menu_relatorios'].click()
                time.sleep(1)
            self.locators['menu_financeiro'].click()
            
            time.sleep(1)

            self.locators['menu_titulos_a_pagar'].wait_for(state="visible")
            self.locators['menu_titulos_a_pagar'].click()    
            
        except PlaywrightTimeoutError:
            logger.error("Falha na navegação ou configuração da planilha")
            raise

    def _fechar_popup_se_existir(self):
        """Método reutilizável para fechar popups"""
        try:
            time.sleep(3)
            if self.locators['popup_fechar'].is_visible():
                self.locators['popup_fechar'].click()
        except Exception as e:
            logger.warning(f" Erro ao verificar popup: {e}")
    
    def _confirmar_operacao(self):
        """confirmação da operação"""
        try:
            time.sleep(2)
            self.locators['botao_confirmar'].click()
            logger.info("operação confirmada")
            time.sleep(5)
            self. _fechar_popup_se_existir() 
        except Exception as e:
            logger.error(f"Falha na confirmação: {e}")
            raise  

    def _criar_planilha(self):
        """Cria a planilha com os dados necessários."""
        try:
            # Clicar para selecionar a opção "Planilha" 
            self.locators['planilha'].click()
            
            # Na opção "Tipo de planilha", selecionar "Formato de planilha" 
            self.locators['tipo_de_planilha'].select_option('Formato de Tabela Xlsx').click() # Usando o texto, pois é mais robusto
            
            logger.info("Configuração de planilha realizada")
        except Exception as e:
            logger.error(f"Falha ao criar planilha: {e}")
            raise

    def _outras_acoes(self):
        """Método para lidar com outras ações."""
        try:
            logger.info("Acessando outras ações")
            # Na opção "Outras Ações", selecionar "Parâmetros" 
            self.locators['outras_acoes'].click()
            self.locators['parametros_menu'].click()
        except Exception as e:
            logger.error(f"Falha ao acessar outras ações: {e}")
            raise

    def _preencher_parametros(self):
        input_do_vencimento = '01/01/2000'
        input_ate_o_vencimento = '31/12/2050'
        input_da_emissao = '01/01/2000'
        input_ate_a_emissao = '31/12/2050'
        input_da_data_contabil = '01/01/2020'
        input_ate_a_data_contabil = '31/07/2025'
        input_data_base = '30/04/2025'
        
        try:
            # parâmetros
            self.locators['do_vencimento'].wait_for(state="visible")
            self.locators['do_vencimento'].click()
            self.locators['do_vencimento'].fill(input_do_vencimento)
            time.sleep(1) 

            self.locators['ate_o_vencimento'].click()
            self.locators['ate_o_vencimento'].fill(input_ate_o_vencimento)
            time.sleep(1) 

            self.locators['da_emissao'].click()
            self.locators['da_emissao'].fill(input_da_emissao)
            time.sleep(1) 

            self.locators['ate_a_emissao'].click()
            self.locators['ate_a_emissao'].fill(input_ate_a_emissao)
            time.sleep(1) 

            self.locators['da_data_contabil'].click()
            self.locators['da_data_contabil'].fill(input_da_data_contabil)
            time.sleep(1) 

            self.locators['ate_a_data_contabil'].click()
            self.locators['ate_a_dcata_contabil'].fill(input_ate_a_data_contabil)
            time.sleep(1) 

            self.locators['data_base'].click()
            self.locators['data_base'].fill(input_data_base)
            time.sleep(1)

            self.locators['ok_btn'].click()

            
        except Exception as e:
            logger.error(f"Falha no preenchimento de parâmetros {e}")
            raise

    def _selecionar_filiais_e_imprimir(self):
        """Seleciona as filiais e inicia a impressão/download."""
        try:
            logger.info("Iniciando seleção de filiais e impressão")
            
            # A janela de seleção de filiais deve aparecer aqui, então esperamos por ela.
            self.locators['marcar_todos_btn'].wait_for(state='visible')
            
            self.locators['marcar_todos_btn'].click() 
            
            # Confirmar seleção
            self.locators['confirmar_btn'].click() 

            # Clicar em "Imprimir"
            # Clicar na opção "Imprimir". 
            self.locators['imprimir_btn'].click()
            
            # Gerenciar download da planilha
            with self.page.expect_download() as download_info:
                # Se for necessário um clique adicional de confirmação,
                pass

            download = download_info.value
            download_path = f'./downloads/{download.suggested_filename()}'
            download.save_as(download_path)
            logger.info(f"Planilha baixada com sucesso: {download_path}")

        except PlaywrightTimeoutError:
            logger.error("Falha na seleção de filiais ou na impressão")
            raise

    def execucao(self):
        """Fluxo principal de extração de planilha financeira."""
        try:
            logger.info('Iniciando extração da planilha financeira - Títulos a Pagar')
            
            self._navegar_e_configurar_planilha()
            self._confirmar_operacao()
            self._criar_planilha()
            self._outras_acoes()
            self._preencher_parametros()
            self._selecionar_filiais_e_imprimir()
            logger.info("Extração da planilha financeira executada com sucesso")
            return {
                'status': 'success',
                'message': 'Extração de Títulos a Pagar completa'
            }
        except Exception as e:
            error_msg = f" Falha na execução da extração: {str(e)}"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}