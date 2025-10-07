"""
Módulo para conciliação manual no sistema Protheus
"""

import time
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from config.settings import Settings
from config.logger import configure_logger
from .exceptions import TimeoutOperacional
from .utils import Utils

logger = configure_logger()

class BackOffice(Utils):
    def __init__(self, page):
        super().__init__(page)
        self.page = page
        self.settings = Settings()
        self._definir_locators()

    def _definir_locators(self):
        logger.info("Definindo seletores Backoffice...")
        
        # Aguardar o iframe principal carregar
        self.page.wait_for_load_state('networkidle')
        time.sleep(3)
        
        # Tentar diferentes seletores de iframe
        self.iframe = self.page.frame_locator("iframe").first
        
        self.locators = {
            # NAVEGAÇÃO MENU LATERAL
            'atualizacoes': self.page.get_by_text('Atualizações', exact=False),
            'mov_bancario': self.page.get_by_text('Movimento Bancario', exact=False),
            'backoffice': self.page.get_by_text('Conciliador Backoffice', exact=False),
            'botao_confirmar': self.page.get_by_role("button", name="Confirmar"),
            'btn_fechar_menu': self.page.get_by_role('button', name='Fechar'),

            # 1. TELA CONCILIADOR BACKOFFICE
            'menu_conciliador': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_role("menuitem", name="Conciliador"),
            'config_conciliacao_dropdown': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_placeholder("Selecione uma configuração de"),

            # 2. CONFIGURAÇÃO DE CONCILIAÇÃO
            'select_conciliacao_manual': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").locator(".po-field-container-content"),
            'label_conciliacao': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("-Conciliação Bancária Manual"),
            'btn_ver_filtros': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_role("button", name="Ver Filtros"),
            
            # 3. FILTROS
            'data_dispon_de': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("Data Dispon. de"),
            'data_dispon_ate': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("Data Dispon. até"),
            'banco': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("Banco igual a", exact=True),
            'agencia': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("Agencia igual a"),
            'conta': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_label("Conta Banco igual a"),
            # 4. CONCILIAÇÃO
            'nao_encontrados': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_text("Dados não Encontrados"),
            'checkbox_selecionar': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_role("row", name="Filial Orig. Data Dispon.").get_by_role("checkbox"),
            'btn_acoes': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_text("Ações", exact=True),
            'btn_conciliar': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_text("Conciliar"),
            
            # 5. APLICAÇÃO
            'aba_dados_conciliacao': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_text("Dados da Conciliação"),
            'btn_aplicar': self.page.frame_locator("internal:attr=[title=\"Ctba940_env_ceos62_dev\"i] >> iframe").get_by_role("button", name="Aplicar"),
            
            # 6. POP UP DE CONFIRMAÇÃO FINAL
            'popup_btn_confirmar': self.page.get_by_role("button", name="Ok"),
            'popup_fechar': self.page.get_by_role("button", name="Fechar")
        }

    def _navegar_menu(self):
        """Navega pelo menu do Protheus até o BackOffice"""
        logger.info("Navegando no menu do Protheus...")
        
        try:
            # Clicar em Atualizações
            self.locators['atualizacoes'].wait_for(state="visible", timeout=30000)
            self.locators['atualizacoes'].click()
            logger.info("Clicou em Atualizações")
            time.sleep(1)
            
            # Clicar em Movimento Bancário
            self.locators['mov_bancario'].wait_for(state="visible", timeout=30000)
            self.locators['mov_bancario'].click()
            logger.info("Clicou em Movimento Bancário")
            time.sleep(1)
            
            # Clicar em Conciliador Backoffice
            self.locators['backoffice'].wait_for(state="visible", timeout=30000)
            self.locators['backoffice'].click()
            logger.info("Clicou em Conciliador Backoffice")
            time.sleep(1)
            # Clicar em Confirmar se existir
            self._confirmar_operacao()
            # Fechar popups novamente
            time.sleep(1)
            self._fechar_popup_se_existir()
            
        except Exception as e:
            error_msg = f"Erro na navegação do menu: {e}"
            logger.error(error_msg)
            raise

    def _navegar_para_conciliador(self):
        """Navega para a tela de Conciliador"""
        logger.info("Navegando para a tela de Conciliador...")
        try:
            # Clicar no menu Conciliador
            self.locators['menu_conciliador'].wait_for(state="visible", timeout=30000)
            self.locators['menu_conciliador'].click()
            logger.info("Clicou no menu Conciliador")
            time.sleep(1)
            
            # Clicar no dropdown de configuração
            self.locators['config_conciliacao_dropdown'].wait_for(state="visible", timeout=30000)
            self.locators['config_conciliacao_dropdown'].click()
            logger.info("Clicou no dropdown de configuração")
            time.sleep(1)
            # Selecionar conciliação manual
            self.locators['select_conciliacao_manual'].wait_for(state="visible", timeout=30000)
            
            time.sleep(2)
            self.locators['label_conciliacao'].click()
            logger.info("Selecionou Conciliação Manual")
            
        except PlaywrightTimeoutError as e:
            error_msg = "Tempo esgotado ao navegar para a tela de Conciliador."
            logger.error(f"{error_msg}: {e}")
            raise TimeoutOperacional(error_msg, "navegacao_conciliador", self.settings.TIMEOUT) from e
        except Exception as e:
            error_msg = f"Falha ao navegar para a tela de Conciliador: {e}"
            logger.error(error_msg)
            raise

    def _preencher_filtros(self, dados_banco):
        """Preenche os filtros de pesquisa"""
        logger.info("Preenchendo filtros de pesquisa...")
        do_banco = dados_banco.get('do_banco')
        da_agencia = dados_banco.get('da_agencia')
        da_conta = dados_banco.get('da_conta')
        data_ontem = self.obter_data_dia_anterior()
        
        try:
            # Expandir filtros 
            self.locators['btn_ver_filtros'].wait_for(state="visible", timeout=30000)
            self.locators['btn_ver_filtros'].click()
            logger.info("seleção de filtros")
            time.sleep(1)
            
            # Preencher datas
            self.locators['data_dispon_de'].fill('01/04/2025')
            self.locators['data_dispon_ate'].fill('01/04/2025')
            time.sleep(1)

            logger.info(f"Preenchendo parâmetros - Banco: {do_banco}, Agência: {da_agencia}, Conta: {da_conta}")

            self.locators['banco'].fill(do_banco)
            time.sleep(1)
            self.locators['agencia'].fill(da_agencia)
            time.sleep(1)
            self.locators['conta'].fill(da_conta)
            time.sleep(1)
            # Aplicar filtros
            self.locators['btn_aplicar'].click()
            logger.info("Filtros aplicados")
            
            
        except Exception as e:
            error_msg = f"Erro ao preencher filtros: {e}"
            logger.error(error_msg)
            raise

    def _selecionar_e_conciliar(self):
        """Seleciona e concilia movimentações"""
        logger.info("Selecionando e conciliando movimentações...")
        try:
            # Navegar para aba "Dados não Encontrados"
            self.locators['nao_encontrados'].wait_for(state="visible", timeout=15000)
            self.locators['nao_encontrados'].click()
            logger.info("Clicou na aba 'Dados não Encontrados'")
            time.sleep(3)
            
            # Verificar se existem registros
            tabela_resultados = self.page.locator('.po-table-container table tbody tr')
            count = tabela_resultados.count()
            
            if count > 0:
                logger.info(f"Foram encontrados {count} registros para conciliação.")
                
                # Selecionar todos os registros
                self.locators['checkbox_selecionar'].wait_for(state="visible", timeout=10000)
                self.locators['checkbox_selecionar'].click()
                logger.info("Todas as movimentações selecionadas")
                time.sleep(2)
                
                # Clicar em Ações > Conciliar
                self.locators['btn_acoes'].wait_for(state="visible", timeout=10000)
                self.locators['btn_acoes'].click()
                time.sleep(1)
                
                self.locators['btn_conciliar'].wait_for(state="visible", timeout=10000)
                self.locators['btn_conciliar'].click()
                logger.info("Botão 'Conciliar' clicado")
                
                self.page.wait_for_load_state('networkidle')
                time.sleep(5)
            else:
                logger.info("Nenhum registro encontrado para conciliação")
                
        except Exception as e:
            error_msg = f"Erro ao selecionar e conciliar: {e}"
            logger.error(error_msg)
            raise
    def _processar_conta(self, dados_do_banco, nome_banco):
        """Processa uma conta bancária individual completa."""
        try:
            logger.info(f'Processando banco: {nome_banco} - {dados_do_banco["do_banco"]}, agência: {dados_do_banco["da_agencia"]}, conta: {dados_do_banco["da_conta"]}')
            # 1. Navegar para o menu backoffice
            self._navegar_menu()
            
            # 2. Navegar para a tela de Conciliador
            self._navegar_para_conciliador()
            self._preencher_filtros(dados_do_banco)

            # 4. Selecionar e conciliar dados
            self._selecionar_e_conciliar()
            
        except Exception as e:
            error_msg = f"Falha no processamento do banco {nome_banco}"
            logger.error(f"{error_msg}: {str(e)}")
            # Registrar como inválido em caso de erro
            self.conciliacao.registrar_banco_invalido(
                nome_banco,
                dados_do_banco["do_banco"],
                dados_do_banco["da_agencia"],
                dados_do_banco["da_conta"]
            )
            return None
        
    def execucao(self):
        """Executa o processo completo do BackOffice"""
        logger.info("Iniciando a conciliação manual no Backoffice do Protheus...")
        try:
            # 3. parametros
            dados_do_banco = {
                'banco_do_brasil': {
                    'da_conta': '118091',
                    'do_banco': '001',
                    'da_agencia': '2115'
                },
                'bradesco_sa': {
                    'da_conta': '105169',
                    'do_banco': '237',
                    'da_agencia': '0895'
                },
                'banco_bs2_sa': {
                    'da_conta': '10164',
                    'do_banco': '218',
                    'da_agencia': '0001'
                }
            }
            arquivos_gerados = []
            for nome_banco, dados_banco in dados_do_banco.items():
                logger.info(f"Processando banco: {nome_banco}")
                self._processar_conta(dados_banco, nome_banco)
                
            return {
                'status': 'success',
                'message': f'Todas as contas processadas com sucesso',
                'arquivos_gerados': arquivos_gerados
            }

        except Exception as e:
            logger.error(f"Falha durante a conciliação no Backoffice: {str(e)}")
            return {'status': 'error', 'message': f'Falha no Backoffice: {str(e)}'}