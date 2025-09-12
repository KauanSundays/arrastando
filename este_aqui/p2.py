import time
import logging
import subprocess
from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import send_keys
import os
import time
import pythoncom
from win32com.client import Dispatch
from pywinauto.keyboard import send_keys



logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)



class LoginRobusto:
    def __init__(self, saj_executable_path):
        self.saj_executable_path = saj_executable_path
        self.app = None
        self.login_window = None
        self.campo_usuario = None

    def abrir_saj(self):
        """Abre o aplicativo SAJ pelo caminho fornecido."""
        try:
            logger.info(f"Iniciando o SAJ: {self.saj_executable_path}")
            subprocess.Popen([self.saj_executable_path], shell=True)
            logger.info("SAJ iniciado com sucesso. Aguarde o carregamento da tela de login...")
            time.sleep(10)  # Aguarda o carregamento inicial
            return True
        except Exception as e:
            logger.error(f"Erro ao iniciar o SAJ: {e}")
            return False

    # Funções de inspeção comentadas para futura referência
    # def inspecionar_controles_janela_login(self):
    #     ...
    # def inspecionar_controles_janela_principal(self, janela_principal):
    #     ...
    # def inspecionar_controles_janela_emissao(self, janela_emissao):
    #     ...

    def localizar_janela_ou_campo_usuario(self):
        """Tenta localizar a janela de login ou, alternativamente, o campo usuário."""
        logger.info("Tentando localizar a janela de login...")
        try:
            self.app = Application(backend="uia").connect(title_re=".*Login.*", timeout=5)
            self.login_window = self.app.window(title_re=".*Login.*")
            if self.login_window.exists():
                logger.info("Janela de login localizada pelo título.")
                self.login_window.set_focus()
                # (Inspeção removida do fluxo principal)
                # self.inspecionar_controles_janela_login()
                try:
                    campo_usuario = self.login_window.child_window(class_name="TspCampo", found_index=0)
                    if campo_usuario.exists():
                        logger.info("Campo usuário localizado dentro da janela de login.")
                        campo_usuario.set_focus()
                        self.campo_usuario = campo_usuario
                        return True
                    else:
                        logger.error("Campo usuário não encontrado dentro da janela de login.")
                except Exception as e:
                    logger.error(f"Erro ao localizar campo usuário dentro da janela de login: {e}")
                return False
        except Exception as e:
            logger.warning(f"Não foi possível localizar a janela de login pelo título: {e}")

        logger.info("Tentando localizar o campo usuário diretamente...")
        try:
            app = Application(backend="uia").connect(path=self.saj_executable_path)
            campo_usuario = app.window(class_name="TspCampo")
            if campo_usuario.exists():
                logger.info("Campo usuário localizado diretamente.")
                campo_usuario.set_focus()
                self.campo_usuario = campo_usuario
                return True
        except Exception as e:
            logger.error(f"Não foi possível localizar o campo usuário: {e}")
        return False

    def preencher_usuario(self, usuario):
        """Preenche o campo usuário e pressiona TAB para avançar para o campo senha."""
        if not self.campo_usuario:
            logger.error("Campo usuário não está disponível para preenchimento.")
            return False
        try:
            self.campo_usuario.set_focus()
            try:
                self.campo_usuario.double_click_input()
                logger.info("Clique duplo realizado no campo usuário.")
            except Exception as e:
                logger.warning(f"Não foi possível clicar duplo no campo usuário: {e}")
            time.sleep(0.7)
            logger.info(f"Preenchendo usuário: {usuario}")
            send_keys(usuario, with_spaces=True)
            time.sleep(0.5)
            logger.info("Pressionando TAB para avançar para o campo senha...")
            send_keys('{TAB}')
            time.sleep(0.5)
            return True
        except Exception as e:
            logger.error(f"Erro ao preencher usuário: {e}")
            return False

    def validar_tela_principal(self, janela_principal):
        """Valida se a tela principal do SAJ está realmente disponível após o login usando o elemento 'tbarPrincipal'."""
        if not janela_principal:
            logger.error("❌ Janela principal não está disponível para validação.")
            return False
        
        try:
            # Verificar se a janela existe e está visível
            try:
                if not janela_principal.exists():
                    logger.error("❌ Janela principal não existe.")
                    return False
                
                if not janela_principal.is_visible():
                    logger.warning("⚠️ Janela principal não está visível, mas tentando continuar...")
            except Exception as e:
                logger.warning(f"⚠️ Verificação de existência/visibilidade falhou: {e}")
            
            # Método principal: Buscar especificamente pelo elemento "tbarPrincipal"
            logger.info("🔍 Validando tela principal pelo elemento 'tbarPrincipal'...")
            
            try:
                # Buscar pelo elemento "tbarPrincipal" com características específicas
                tbar_principal = janela_principal.child_window(
                    name="tbarPrincipal",
                    class_name="TspToolBar",
                    control_type="ToolBar"
                )
                
                if tbar_principal.exists() and tbar_principal.is_visible():
                    logger.info("✅ Validação principal: elemento 'tbarPrincipal' encontrado e visível!")
                    return True
                else:
                    logger.warning("⚠️ Elemento 'tbarPrincipal' não está visível...")
                    
            except Exception as e:
                logger.info(f"⚠️ Busca direta pelo 'tbarPrincipal' falhou: {e}")
            
            # Método alternativo: Buscar por todos os elementos filhos
            logger.info("🔄 Tentando validação alternativa por busca em todos os elementos...")
            
            try:
                filhos = janela_principal.descendants()
                logger.info(f"🔍 Analisando {len(filhos)} elementos filhos da janela principal...")
                
                # Critérios de validação baseados no elemento "tbarPrincipal"
                elementos_encontrados = []
                
                for ctrl in filhos:
                    try:
                        nome = ctrl.window_text().strip()
                        classe = ctrl.element_info.class_name
                        
                        # Critério principal: elemento "tbarPrincipal"
                        if nome == "tbarPrincipal" and classe == "TspToolBar":
                            elementos_encontrados.append(f"tbarPrincipal: '{nome}' (classe: {classe})")
                            logger.info(f"✅ Elemento 'tbarPrincipal' encontrado na busca alternativa!")
                            return True
                        
                        # Critérios secundários: outros elementos da barra de ferramentas
                        if classe == "TspToolBar":
                            elementos_encontrados.append(f"TspToolBar: '{nome}'")
                            logger.info(f"✅ Barra de ferramentas encontrada: '{nome}'")
                            return True
                        
                        # Critérios de fallback: elementos específicos do SAJ
                        if nome in ["Fluxo de trabalho", "Agenda", "Sistema"] and classe in ["TspButton", "TspMenuItem"]:
                            elementos_encontrados.append(f"botão/menu: '{nome}' (classe: {classe})")
                            logger.info(f"✅ Elemento específico do SAJ encontrado: '{nome}'")
                            return True
                            
                    except Exception as e:
                        continue
                
                # Validação final: verificar se há elementos suficientes indicando carregamento
                if len(filhos) > 15:
                    logger.info(f"✅ Validação final: janela tem {len(filhos)} elementos filhos (indicando carregamento completo)")
                    return True
                    
            except Exception as e:
                logger.warning(f"⚠️ Busca por elementos filhos falhou: {e}")
            
            # Verificar o título da janela como último recurso
            try:
                titulo_janela = janela_principal.window_text().lower()
                if 'saj' in titulo_janela and 'procurador' in titulo_janela:
                    logger.info(f"✅ Validação final: título da janela confirma SAJ: {titulo_janela}")
                    return True
            except Exception as e:
                logger.warning(f"⚠️ Verificação do título falhou: {e}")
            
            # Se chegou até aqui, a validação falhou
            logger.error(f"❌ Não foi possível validar a tela principal do SAJ.")
            return False
            
        except Exception as e:
            logger.error(f"❌ Erro durante validação da tela principal: {e}")
            return False

    def preencher_senha_e_entrar(self, senha):
        """Preenche o campo senha, aciona o login e valida se a janela principal abriu."""
        try:
            if not self.login_window:
                logger.error("Janela de login não está disponível para localizar o campo senha.")
                return False
            logger.info("Localizando campo senha na janela de login...")
            campo_senha = self.login_window.child_window(class_name="TspCampo", found_index=1)
            if not campo_senha.exists():
                logger.error("Campo senha não localizado na janela de login.")
                return False
            campo_senha.set_focus()
            try:
                campo_senha.double_click_input()
                logger.info("Clique duplo realizado no campo senha.")
            except Exception as e:
                logger.warning(f"Não foi possível clicar duplo no campo senha: {e}")
            time.sleep(0.7)
            logger.info("Preenchendo senha usando type_keys.")
            try:
                campo_senha.type_keys(senha, with_spaces=True, set_foreground=True)
                logger.info("Senha enviada com type_keys.")
            except Exception as e:
                logger.error(f"Erro ao usar type_keys no campo senha: {e}")
                logger.info("Tentando fallback com send_keys global.")
                send_keys(senha, with_spaces=True)
            time.sleep(0.5)
            logger.info("Pressionando TAB para garantir navegação até o botão Entrar.")
            send_keys('{TAB}')
            time.sleep(0.5)
            logger.info("Acionando Alt+E para entrar.")
            send_keys('%E')  # Alt+E
            logger.info("Aguardando abertura da janela principal do SAJ...")
            
            # Estratégia robusta de detecção baseada no elemento específico "tbarPrincipal"
            tempo_max_espera = 120  # Aumentado para 2 minutos
            tempo_inicial = time.time()
            janela_principal_aberta = False
            janela = None
            
            logger.info(f"🔍 Iniciando detecção da janela principal pelo elemento 'tbarPrincipal' (tempo máximo: {tempo_max_espera}s)")
            
            while time.time() - tempo_inicial < tempo_max_espera:
                try:
                    from pywinauto import Desktop
                    janelas = Desktop(backend="uia").windows()
                    
                    # Buscar pela janela principal do SAJ primeiro
                    for w in janelas:
                        try:
                            classe = w.element_info.class_name
                            titulo = w.window_text().lower()
                            
                            # Método 1: Buscar pela classe específica TfspjPrincipal
                            if classe == "TfspjPrincipal":
                                logger.info(f"✅ Janela principal detectada pelo método 1 (classe): {classe}")
                                janela = w
                                break
                            
                            # Método 2: Buscar por título contendo 'saj' e 'procurador'
                            if 'saj' in titulo and 'procurador' in titulo:
                                logger.info(f"✅ Janela principal detectada pelo método 2 (título): {titulo}")
                                janela = w
                                break
                            
                            # Método 3: Buscar por classe TfspjTelaMenu (alternativa)
                            if classe == "TfspjTelaMenu":
                                logger.info(f"✅ Janela principal detectada pelo método 3 (TfspjTelaMenu): {classe}")
                                janela = w
                                break
                                
                        except Exception as e:
                            continue
                    
                    # Se encontrou uma janela candidata, verificar se contém o elemento "tbarPrincipal"
                    if janela:
                        logger.info("🔍 Verificando se a janela contém o elemento 'tbarPrincipal'...")
                        
                        # Tentar busca alternativa por todos os elementos filhos (método mais confiável)
                        try:
                            filhos = janela.descendants()
                            for filho in filhos:
                                try:
                                    nome = filho.window_text()
                                    classe_filho = filho.element_info.class_name
                                    
                                    if nome == "tbarPrincipal" and classe_filho == "TspToolBar":
                                        logger.info("✅ Elemento 'tbarPrincipal' encontrado na busca alternativa!")
                                        janela_principal_aberta = True
                                        break
                                except Exception:
                                    continue
                            
                            if janela_principal_aberta:
                                break
                                
                        except Exception as e:
                            logger.warning(f"⚠️ Busca por elementos filhos falhou: {e}")
                            
                            # Tentar busca direta pelo child_window se disponível
                            try:
                                if hasattr(janela, 'child_window'):
                                    tbar_principal = janela.child_window(
                                        name="tbarPrincipal",
                                        class_name="TspToolBar",
                                        control_type="ToolBar"
                                    )
                                    
                                    if tbar_principal.exists() and tbar_principal.is_visible():
                                        logger.info("✅ Elemento 'tbarPrincipal' encontrado! Confirmação de que estamos na janela principal do SAJ.")
                                        janela_principal_aberta = True
                                        break
                                    else:
                                        logger.info("⚠️ Janela encontrada, mas elemento 'tbarPrincipal' não está visível ainda...")
                                else:
                                    logger.info("⚠️ Janela não suporta child_window, mas elemento pode estar presente...")
                                    
                            except Exception as e2:
                                logger.info(f"⚠️ Busca direta pelo 'tbarPrincipal' falhou: {e2}")
                    
                    # Log de progresso a cada 10 segundos
                    tempo_decorrido = int(time.time() - tempo_inicial)
                    if tempo_decorrido % 10 == 0 and tempo_decorrido > 0:
                        logger.info(f"⏳ Aguardando... {tempo_decorrido}s decorridos. Janelas encontradas: {len(janelas)}")
                        
                        # Listar algumas janelas para debug
                        janelas_saj = [w for w in janelas if 'saj' in w.window_text().lower()]
                        if janelas_saj:
                            logger.info(f"🔍 Janelas SAJ detectadas: {[w.window_text()[:30] for w in janelas_saj[:3]]}")
                    
                except Exception as e:
                    logger.warning(f"⚠️ Erro na tentativa de detecção: {e}")
                
                time.sleep(1)
            
            if not janela_principal_aberta:
                logger.error(f"❌ Elemento 'tbarPrincipal' não foi encontrado após {tempo_max_espera}s. Possível erro de login.")
                
                # Tentativa final: verificar se há alguma janela do SAJ
                try:
                    janelas_finais = Desktop(backend="uia").windows()
                    janelas_saj_finais = [w for w in janelas_finais if 'saj' in w.window_text().lower()]
                    if janelas_saj_finais:
                        logger.info(f"🔍 Janelas SAJ encontradas na verificação final: {[w.window_text() for w in janelas_saj_finais]}")
                        # Tentar usar a primeira janela SAJ encontrada
                        janela = janelas_saj_finais[0]
                        logger.info(f"🔄 Tentando usar janela alternativa: {janela.window_text()}")
                        if self.validar_tela_principal(janela):
                            logger.info("✅ Validação bem-sucedida com janela alternativa!")
                            return True
                except Exception as e:
                    logger.error(f"❌ Falha na verificação final: {e}")
                    
                
                return False
            
            logger.info(f"✅ Janela principal confirmada pelo elemento 'tbarPrincipal' após {int(time.time() - tempo_inicial)}s")
            time.sleep(3)  # Aguarda estabilização
            
            # Validação da tela principal
            if self.validar_tela_principal(janela):
                logger.info("✅ Validação final: etapa 1 (login) concluída com sucesso! Tela principal disponível.")
                
                # Aguardar 1 segundo após a validação da tela principal
                logger.info("⏳ Aguardando 1 segundo após validação da tela principal...")
                time.sleep(5)
                
                # Iniciar emissão de documentos através do atalho CTRL+D
                logger.info("🚀 Iniciando emissão de documentos com atalho CTRL+D...")
                try:
                    logger.info("🚀 Enviando CTRL+D para abrir emissão de documentos...")
                    send_keys('^d')

                    inicio = time.time()
                    timeout = 15

                    while time.time() - inicio < timeout:
                        try:
                            saj = Desktop(backend="uia").window(title="SAJ Procuradorias", control_type="Window")

                            if saj.exists(timeout=1):
                                emissao = saj.child_window(title="Emissão de Documentos", control_type="Window")
                                if emissao.exists(timeout=1) and emissao.is_visible():
                                    logger.info("✅ 'Emissão de Documentos' encontrada dentro do SAJ.")
                                    emissao.set_focus()
                                    return True
                        except Exception:
                            pass
                except Exception as e:
                    logger.error(f"❌ Erro ao abrir 'Emissão de Documentos': {e}")
                    return None

            else:
                logger.error("❌ Validação final: etapa 1 (login) FALHOU. Tela principal não validada.")
                return False
            return True
        except Exception as e:
            logger.error(f"Erro ao preencher senha e acionar login: {e}")
            return False

    def abrir_emissao_documentos(self):
        """Aciona Ctrl+D para abrir a janela de Emissão de Documentos e confirma sua abertura pelo backend 'uia' como filho da janela principal. Tenta novamente caso não encontre na primeira tentativa."""
        try:
            logger.info("Acionando Ctrl+D para abrir a janela de Emissão de Documentos...")
            send_keys('^d')  # Ctrl+D
            tempo_max_espera = 20
            tempo_inicial = time.time()
            janela_emissao = None
            # Conecta à janela principal do SAJ
            app = Application(backend="uia").connect(title_re="SAJ Procuradorias")
            main_win = app.window(title_re="SAJ Procuradorias", class_name="TfspjTelaMenu")

            for w in Desktop(backend="uia").windows(): 
                print("->", w.window_text(), "| classe:", w.class_name()) # Se for a janela "Emissão de Documentos", trabalha nela 
                if w.window_text().strip() == "Emissão de Documentos": 
                    logger.info("📝 Preenchendo campos da janela 'Emissão de Documentos'...")
                    return True
            logger.error("Deu ruim")
            return False
        except Exception as e:
            logger.error(f"Erro ao abrir ou localizar a janela de Emissão de Documentos: {e}")
            return False

    def preencher_nosso_numero(self, valor_referencia):
        """Preenche o campo 'Nosso Número' na janela de Emissão de Documentos, removendo apenas o trecho '.01' intermediário."""
        import re
        logger.info("Aguardando posicionamento no campo 'Nosso Número'...")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        saj = Desktop(backend="uia").window(title="SAJ Procuradorias", control_type="Window")

        if saj.exists(timeout=1):
            logger.info("🔍 Exportando elementos dentro do SAJ...")

            # with open("saj_elementos.txt", "w", encoding="utf-8") as f:
            #     for elem in saj.descendants():
            #         try:
            #             f.write(
            #                 f"Name: {elem.window_text()} | "
            #                 f"Class: {elem.element_info.class_name} | "
            #                 f"Type: {elem.element_info.control_type}\n"
            #             )
            #         except Exception as e:
            #             f.write(f"[ERRO AO LER ELEMENTO] {e}\n")

            # logger.info("✅ Elementos exportados para 'saj_elementos.txt'")


        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")
        logger.info("--------------------------------")


        # remove apenas o '.01' que aparece entre pontos
        valor_digitado = re.sub(r'\.01\.', '.', valor_referencia)

        logger.info(f"Preenchendo campo 'Nosso Número' com o valor: {valor_digitado}")
        print("valor é valor referencia...", valor_digitado)

        time.sleep(3)
        try:
            send_keys('^a{BACKSPACE}')
            time.sleep(0.5)
            send_keys('^a{BACKSPACE}')
            time.sleep(0.5)

            for digito in valor_digitado:
                send_keys(digito)
                time.sleep(0.5)

            time.sleep(3)
            send_keys('{TAB}')
            time.sleep(4)
            send_keys('{ENTER}')
            send_keys('{ENTER}')
            logger.info("✅ Campo 'Nosso Número' preenchido e TAB pressionado.")
            return True
        except Exception as e:
            logger.error(f"❌ Erro ao preencher campo 'Nosso Número': {e}")
            return False


    def preencher_codigo_modelo(self, valor, nosso_numero, pasta_selecionada):
        """Preenche o campo 'Código do Modelo' na janela de Emissão de Documentos."""
        logger.info("🔎 Procurando janela principal do SAJ para localizar o campo Código do Modelo...")

        try:
            janela_saj = Desktop(backend="uia").window(title="SAJ Procuradorias", class_name="TfspjTelaMenu")
            logger.info("✅ Janela principal localizada.")

            # Busca todos os campos do tipo TspCampoMascara
            campos = janela_saj.descendants(control_type="Pane", class_name="TspCampoMascara")
            logger.info(f"🔍 Total de campos TspCampoMascara encontrados: {len(campos)}")

            # Usa o primeiro candidato da lista
            campo_alvo = campos[4]

            # Clica e envia o texto
            campo_alvo.click_input(double=True)
            time.sleep(0.5)
            send_keys(str(valor))
            time.sleep(0.2)
            send_keys('{TAB}')
            logger.info("✅ Valor inserido e TAB pressionado no campo Código do Modelo.")
            time.sleep(5)
            # Chama o preenchimento do campo Nosso Número
            if not self.preencher_nosso_numero(nosso_numero):
                return False
            logger.info("Campo Código do Modelo' preenchido com sucesso!")

            # logger.info("Abrindo o editor de texto do SAJ...")
            # modal_selecao_processos = self.verificar_selecao_processos()

            # NOVO: Abrir o editor do SAJ antes de inserir a minuta e o texto do Word
            logger.info("Abrindo o editor de texto do SAJ...")
            janela_principal = Desktop(backend="uia").window(
                title="SAJ Procuradorias",
                class_name="TfspjTelaMenu"
            )
            print("------------------------")
            print("------------------------")
            print("------------------------")
            print("------------------------")
            print("------------------------")
            editor_element = self.abrir_editor_texto_saj()
            if not editor_element:
                logger.error("Falha ao abrir o editor de texto do SAJ.")
                return False
            logger.info("Editor de texto do SAJ aberto com sucesso!")

            # Após abrir o editor, executa a etapa de Inserção Minuta
            logger.info("Iniciando etapa de Inserção Minuta...")
            if not self.inserir_minuta():
                logger.error("Falha na etapa de Inserção Minuta.")
                return False
            logger.info("Etapa de Inserção Minuta concluída com sucesso!")

            # Integração: inserir texto do Word no editor do SAJ
            categoria = str(valor)
            logger.info(f"Iniciando inserção do texto do Word no editor do SAJ para nosso_numero={nosso_numero}, categoria={categoria}")
            if not self.inserir_texto_word_no_editor_saj(nosso_numero, categoria, pasta_selecionada):
                logger.error("Falha ao inserir texto do Word no editor do SAJ.")
                return False
            logger.info("Texto do Word inserido com sucesso no editor do SAJ!")
            
            # NOVA FASE: Salvar Minuta
            logger.info("Iniciando fase: Salvar Minuta...")
            if not self.salvar_minuta():
                logger.error("Falha na fase de Salvar Minuta.")
                return False
            logger.info("Fase de Salvar Minuta concluída com sucesso!")
            
            return True

        except Exception as e:
            logger.error(f"❌ Erro ao preencher código do modelo: {e}")
            return False

    def verificar_selecao_processos(self):
        logger.info("✅ Verificando se tela de Seleção de processos está aberta.")

        try:
            inicio = time.time()
            timeout = 15

            while time.time() - inicio < timeout:
                try:
                    saj = Desktop(backend="uia").window(title="SAJ Procuradorias", control_type="Window")

                    if saj.exists(timeout=1):
                        logger.info("🔍 Exportando elementos dentro do SAJ...")

                        with open("saj_elementos22.txt", "w", encoding="utf-8") as f:
                            for elem in saj.descendants():
                                try:
                                    f.write(
                                        f"Name: {elem.window_text()} | "
                                        f"Class: {elem.element_info.class_name} | "
                                        f"Type: {elem.element_info.control_type}\n"
                                    )
                                except Exception as e:
                                    f.write(f"[ERRO AO LER ELEMENTO] {e}\n")

                        logger.info("✅ Elementos exportados para 'saj_elementos.txt'")


                        emissao = saj.child_window(title="Seleção de Processos Dependentes", control_type="Window")
                        if emissao.exists(timeout=1) and emissao.is_visible():
                            logger.info("✅ 'Seleção de Processos Dependentes' encontrada dentro do SAJ.")
                            emissao.set_focus()
                            return True

                except Exception:
                    pass
        except Exception as e:
            logger.error(f"❌ Erro ao abrir 'Seleção de Processos Dependentes': {e}")
            return None
        
    def abrir_editor_texto_saj(self):
        send_keys('{ENTER}')
        send_keys('{ENTER}')
        janela_principal = Desktop(backend="uia").window(
            title="SAJ Procuradorias",
            class_name="TfspjTelaMenu"
        )
        editor_element = janela_principal.child_window(
            class_name="TedtWPRichText",
            control_type="Pane"
        )
        saj = Desktop(backend="uia").window(title="SAJ Procuradorias", control_type="Window")
        emissao = saj.child_window(title="Emissão de Documentos", control_type="Window")

        toolbar = emissao.child_window(title="spToolBar1", control_type="ToolBar")
        # emissao.print_control_identifiers()
        toolbar.print_control_identifiers()
        botoes = toolbar.descendants(control_type="Button")

        time.sleep(2)
        send_keys('%t')
        time.sleep(2)

        # Clicar no segundo botão (índice 1, porque começa do 0)
        if len(botoes) > 1:
            botoes[1].click_input()
            print("✅ Segundo botão clicado!")
        else:
            print("⚠️ Não há botão suficiente na toolbar.")

        logger.info("Aguardando 3 segundos antes de abrir o editor de texto SAJ...")
        time.sleep(3)

        try:
            janela_existente = Desktop(backend="uia").window(title_re=".*documento associado.*", control_type="Window")
            if janela_existente.exists(timeout=2):
                logger.info("⚠️ Documento já associado detectado. Clicando em 'Editar'...")
                # Procura botão "Editar" (pode ser por nome ou índice)
                botao_editar = janela_existente.child_window(title_re=".*Editar.*", control_type="Button")
                if botao_editar.exists():
                    botao_editar.click_input()
                    logger.info("✅ Botão 'Editar' clicado.")
                    time.sleep(2)
        except Exception as e:
            logger.warning(f"Não foi possível clicar em 'Editar': {e}")


        logger.info("✅ Atalho Alt+t enviado para abrir o editor de texto SAJ.")

        # 2️⃣ fallback no loop
        tempo_max_espera = 20
        tempo_inicial = time.time()
        while time.time() - tempo_inicial < tempo_max_espera:
            logger.info("procurando via fallback")
            send_keys('%d')
            try:

                if editor_element.exists() and editor_element.is_visible():
                    logger.info("✅ Editor encontrado no fallback!")
                    return editor_element
            except Exception:
                pass
            time.sleep(0.5)

        logger.error("❌ Editor não encontrado dentro do tempo total.")
        return None


    def simular_setas_para_baixo_simples(self, quantidade=14):
        """Simula pressionar a tecla seta para baixo várias vezes - mais rápido."""
        try:
            logger.info(f"Simulando {quantidade} setas para baixo...")
            from pywinauto.keyboard import send_keys

            janela_principal = Desktop(backend="uia").window(
                title="SAJ Procuradorias",
                class_name="TfspjTelaMenu"
            )
            editor_element = janela_principal.child_window(
                class_name="TedtWPRichText",
                control_type="Pane"
            )

            if editor_element.exists() and editor_element.is_visible():
                editor_element.set_focus()
                time.sleep(0.1)  # foco rápido
                send_keys("{DOWN " + str(quantidade) + "}")  # envia todas as setas de uma vez
                logger.info(f"[OK] {quantidade} setas para baixo executadas")
                return True
            else:
                logger.error("❌ Editor de texto não localizado para setas para baixo")
                return False
        except Exception as e:
            logger.error(f"Erro ao simular setas para baixo: {e}")
            return False

    def inserir_minuta(self):
        """Etapa de Inserção Minuta: envia 14 toques para baixo revalidando o editor."""
        logger.info("🔄 Iniciando etapa: Inserção Minuta")
        logger.info("Aguardando 1 segundo da etapa anterior...")
        time.sleep(1.0)
        return self.simular_setas_para_baixo_simples(quantidade=14)

    def inserir_texto_word_no_editor_saj(self, nosso_numero, categoria, pasta_selecionada):
        """Abre o arquivo Word na pasta do processo, copia o texto com atalhos e cola no editor do SAJ, sem alterar o foco do editor. Sempre revalida o editor antes de colar."""
        import os
        import time
        import pythoncom
        from win32com.client import Dispatch
        from pywinauto.keyboard import send_keys
        base_dir = pasta_selecionada
        pasta = os.path.join(base_dir, nosso_numero)
        pasta = os.path.join(base_dir, nosso_numero)
        logger.info(f"Buscando arquivo Word na pasta: {pasta}")
        if not os.path.exists(pasta):
            logger.error(f"Pasta do processo não encontrada: {pasta}")
            return False
        arquivos_docx = [f for f in os.listdir(pasta) if f.lower().endswith('.docx')]
        if not arquivos_docx:
            logger.error(f"Nenhum arquivo .docx encontrado na pasta: {pasta}")
            return False
        nome_arquivo = arquivos_docx[0]
        caminho_arquivo = os.path.join(pasta, nome_arquivo)
        logger.info(f"Arquivo Word encontrado: {nome_arquivo}")
        logger.info(f"Caminho completo: {caminho_arquivo}")
        try:
            pythoncom.CoInitialize()
            word = Dispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(caminho_arquivo)
            logger.info("Word aberto em background.")
            doc.Content.Copy()
            logger.info("Texto copiado do Word usando métodos COM (sem alterar foco).")
            doc.Close(False)
            word.Quit()
            logger.info("Word fechado.")
            time.sleep(1.5)
            # Sempre revalida o editor antes de colar
            janela_principal = Desktop(backend="uia").window(
                title="SAJ Procuradorias",
                class_name="TfspjTelaMenu"
            )
            editor_element = janela_principal.child_window(
                class_name="TedtWPRichText",
                control_type="Pane"
            )
            if editor_element.exists() and editor_element.is_visible():
                editor_element.set_focus()
                send_keys('^v')
                logger.info("✅ Texto colado no editor do SAJ com Ctrl+V (foco preservado).")
                return True
            else:
                logger.error("❌ Editor de texto não localizado para colagem do Word")
                return False
        except Exception as e:
            logger.error(f"Erro ao inserir texto do Word no editor do SAJ: {e}")
            return False


    def salvar_minuta(self):
        """Fase de Salvar Minuta: salva o documento com ALT+B e depois fecha o editor usando o botão 'Sair do Editor'."""
        from pywinauto.keyboard import send_keys
        logger.info("🔄 Iniciando fase: Salvar Minuta (salvamento e fechamento editor)")
        logger.info("Aguardando 2 segundos para processamento do texto...")
        time.sleep(2.0)
        try:
            # 1. Salvar o documento usando o atalho CTRL+B
            logger.info("💾 Salvando documento com atalho CTRL+B...")
            send_keys('^b')  # CTRL+B
            logger.info("✅ Atalho CTRL+B enviado para salvar o documento")
            
            # 2. Aguardar 5 segundos conforme solicitado
            logger.info("⏳ Aguardando 5 segundos após o salvamento...")
            time.sleep(5.0)
            
            # 3. Localizar e clicar no botão "Sair do Editor" usando o elemento TdxBarSubMenuControl
            logger.info("✅ Salvo")
            return True
        except Exception as e:
            logger.error(f"❌ Erro ao salvar minuta: {e}")
            return False

    def login_robusto(self, usuario=None, senha=None):
        """Executa o login robusto simplificado: localizar campos, preencher usuário, preencher senha e validar login. Após o login, aciona a emissão de documentos e confirma a janela."""
        localizado = self.localizar_janela_ou_campo_usuario()
        if not localizado:
            logger.error("Não foi possível localizar a janela de login nem o campo usuário.")
            return False
        logger.info("Pronto para iniciar o preenchimento dos campos.")
        if usuario is None:
            usuario = "USUARIO_EXEMPLO"
        if not self.preencher_usuario(usuario):
            logger.error("Falha ao preencher usuário.")
            return False
        logger.info("Usuário preenchido. Preenchendo senha...")
        if senha is None:
            senha = "SENHA_EXEMPLO"
        if not self.preencher_senha_e_entrar(senha):
            logger.error("Falha ao preencher senha ou ao validar login.")
            return False
        logger.info("Login robusto realizado com sucesso!")
        return True

if __name__ == "__main__":
    import os
    import sys
    import time
    import pandas as pd
    import subprocess
    import tkinter as tk
    from tkinter import filedialog

    def prompt_config():
        resultado = {}

        root = tk.Tk()
        root.title("Configuração")
        root.geometry("420x360")
        root.resizable(False, False)

        usuario_var = tk.StringVar(value="Franciscoluiz")
        senha_var = tk.StringVar(value="Flvn1191")
        exe_var = tk.StringVar(value="C:/Users/luizvnogueira25/Downloads/saj/spjcliente.exe")
        planilha_var = tk.StringVar(value="C:/Users/luizvnogueira25/Documents/27.08/27.08_tratada.xlsx")
        pasta_var = tk.StringVar(value=os.getcwd())

        tk.Label(root, text="Usuário").pack(pady=(10, 0))
        tk.Entry(root, textvariable=usuario_var).pack(fill="x", padx=16)

        tk.Label(root, text="Senha").pack(pady=(10, 0))
        senha_entry = tk.Entry(root, textvariable=senha_var, show="*")
        senha_entry.pack(fill="x", padx=16)

        def toggle_senha():
            senha_entry.config(show="" if senha_entry.cget("show") == "*" else "*")

        tk.Checkbutton(root, text="Mostrar senha", command=toggle_senha).pack(pady=4)

        def escolher_exe():
            path = filedialog.askopenfilename(filetypes=[("Executável", "*.exe")])
            if path:
                exe_var.set(path)

        def escolher_planilha():
            path = filedialog.askopenfilename(filetypes=[("Planilha Excel", "*.xlsx")])
            if path:
                planilha_var.set(path)

        tk.Label(root, text="Executável do SAJ").pack(pady=(10, 0))
        exe_frame = tk.Frame(root)
        exe_frame.pack(fill="x", padx=16)
        tk.Entry(exe_frame, textvariable=exe_var).pack(side="left", fill="x", expand=True)
        tk.Button(exe_frame, text="...", command=escolher_exe).pack(side="right")

        tk.Label(root, text="Planilha base").pack(pady=(10, 0))
        planilha_frame = tk.Frame(root)
        planilha_frame.pack(fill="x", padx=16)
        tk.Entry(planilha_frame, textvariable=planilha_var).pack(side="left", fill="x", expand=True)
        tk.Button(planilha_frame, text="...", command=escolher_planilha).pack(side="right")

        tk.Label(root, text="Pasta base").pack(pady=(10, 0))
        pasta_frame = tk.Frame(root)
        pasta_frame.pack(fill="x", padx=16)
        tk.Entry(pasta_frame, textvariable=pasta_var).pack(fill="x", expand=True)

        def confirmar():
            resultado.update({
                "usuario": usuario_var.get().strip(),
                "senha": senha_var.get(),
                "executavel": exe_var.get(),
                "planilha": planilha_var.get(),
                "pasta": pasta_var.get()
            })
            root.destroy()

        tk.Button(root, text="Salvar e Iniciar", command=confirmar).pack(pady=16)
        root.mainloop()
        return resultado

    time.sleep(2)

    try:
        config = prompt_config()
        if not config.get("usuario") or not config.get("senha"):
            sys.exit(0)

        df = pd.read_excel(config["planilha"])

        for _, row in df.iterrows():
            subprocess.call("taskkill /f /im spjcliente.exe", shell=True)
            subprocess.call("taskkill /f /im spjclienteapp.exe", shell=True)
            subprocess.call("taskkill /f /im spjChromium.exe", shell=True)

            SAJ_EXECUTABLE_PATH = config["executavel"]
            login = LoginRobusto(SAJ_EXECUTABLE_PATH)

            if login.abrir_saj():
                login.login_robusto(usuario=config["usuario"], senha=config["senha"])
                nosso_numero = str(row["Nosso número"])
                codigo_modelo = int(row["Código do Modelo"])
                logger.info(f"=== Iniciando automação para processo: {nosso_numero} | Código do Modelo: {codigo_modelo} ===")

                logger.info("Iniciando preenchimento do campo Código do Modelo...")

                if login.preencher_codigo_modelo(codigo_modelo, nosso_numero, config["pasta"]):
                    logger.info("Campo Código do Modelo preenchido com sucesso!")
                else:
                    logger.error("Falha ao preencher o campo Código do Modelo.")
                    break
            else:
                logger.error("Não foi possível iniciar o SAJ.")
    except Exception as e:
        logger.exception("Erro inesperado")
