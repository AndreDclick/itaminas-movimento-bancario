from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from config.logger import configure_logger
import time
#teste
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
            'menu_titulos_a_pagar': self.page.get_by_text('Títulos a Pagar'),
            
            # Janela "Posição dos Títulos a Pagar"
            'opcao_planilha': self.page.get_by_role('option', name='Planilha'),
            'tipo_planilha_combobox': self.page.locator('#COMP4547').get_by_role('combobox'),
            'outras_acoes_btn': self.page.get_by_role('button', name='Outras Ações'),
            'parametros_menu': self.page.get_by_text('Parâmetros'),
            'imprimir_btn': self.page.get_by_role('button', name='Imprimir'),

            # Janela de Parâmetros
            'do_vencimento_campo': self.page.get_by_label('Do Vencimento ?'),
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
            self.locators['menu_relatorios'].wait_for(state="visible", timeout=5000)
            self.locators['menu_relatorios'].dblclick()
            logger.info("Menu Relatórios clicado")
            
            time.sleep(1)
            self.locators['menu_financeiro'].click()
            self.locators['menu_titulos_a_pagar'].click()
            
            # Clicar para selecionar a opção "Planilha" 
            self.locators['opcao_planilha'].click()
            
            # Na opção "Tipo de planilha", selecionar "Formato de planilha" 
            self.locators['tipo_planilha_combobox'].select_option('Formato de Tabela Xlsx') # Usando o texto, pois é mais robusto
            
            logger.info("Configuração de planilha realizada")

        except PlaywrightTimeoutError:
            logger.error("Falha na navegação ou configuração da planilha")
            raise

    def _preencher_parametros(self):
        """Preenche os parâmetros conforme o detalhamento do processo."""
        try:
            logger.info("Acessando e preenchendo parâmetros")
            # Na opção "Outras Ações", selecionar "Parâmetros" 
            self.locators['outras_acoes_btn'].click()
            self.locators['parametros_menu'].click()

            # Preencher campos de data de vencimento
            # Do Vencimento: 01/01/2000 
            # Até o Vencimento: 31/12/2050 
            self.locators['do_vencimento_campo'].fill('01/01/2000')
            self.locators['ate_o_vencimento_campo'].fill('31/12/2050')

            # Preencher campos de data de emissão
            self.locators['da_emissao_campo'].fill('01/01/2000')
            self.locators['ate_emissao_campo'].fill('31/12/2050')
            
            # Preencher campos de data contábil
            self.locators['da_data_contabil_campo'].fill('01/01/2000') 
            # A data contábil final deverá ser a do mês que está sendo conciliado.
            self.locators['ate_data_contabil_campo'].fill('20/05/2025') # Exemplo com a data do documento

            # Clicar em OK para fechar a janela de parâmetros
            self.locators['ok_btn'].click()

            logger.info("Parâmetros preenchidos com sucesso")

        except PlaywrightTimeoutError:
            logger.error("Falha ao preencher os parâmetros")
            raise

    def _selecionar_filiais_e_imprimir(self):
        """Seleciona as filiais e inicia a impressão/download."""
        try:
            logger.info("Iniciando seleção de filiais e impressão")
            
            # A janela de seleção de filiais deve aparecer aqui, então esperamos por ela.
            self.locators['marcar_todos_btn'].wait_for(state='visible')
            
            # Marcar todos
            # "Selecionar as opções de 'Matriz e Filial' " 
            # Assumindo que "Marca Todos" faz isso.
            self.locators['marcar_todos_btn'].click() 
            
            # Confirmar seleção
            self.locators['confirmar_btn'].click() 

            # Clicar em "Imprimir"
            # Clicar na opção "Imprimir". 
            self.locators['imprimir_btn'].click()
            
            # Gerenciar download da planilha
            with self.page.expect_download() as download_info:
                # Se for necessário um clique adicional de confirmação,
                # ele pode ser colocado aqui.
                # Exemplo: self.locators['confirmar_btn'].click()
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