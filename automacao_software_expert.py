import sys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
from time import sleep
import tkinter as tk
from tkinter import simpledialog
import threading
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
import pandas as pd


class SoftwareExpertRNC:
    def __init__(self, dados_excel=None):

        self.modo_headless = False  #  VAR PARA RODAR EM 2° PLANO.
        self.executar_selecao_dropdown = True # VAR PARA CONTROLAR SE A "AÇÃO ISOLADA" SERÁ EXECUTADA OU NÃO.
        self.janela_principal = None  # Janela principal do sistema 
        self.form_window_handle = None  # Janela do formulário
        options = webdriver.ChromeOptions() #

        if self.modo_headless: # ADICIONA AS OPÇÕES NO WEBDRIVER
            options.add_argument('--headless') 
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
        
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        self.wait = WebDriverWait(self.driver, 30)

        # Variável que armazenará o código de autenticação quando fornecido pelo usuário
        self.codigo_obtido = None

        self.dados = {  # RESPOSTAS PADRÃO DO FORMULÁRIO
            # Login
            "usuario": "igor.silva",
            "senha": "Pira4321++",
            
            # RNC Principal
            "titulo_rnc": "RNC Automatizada",
            "reincidente": "Não", # PADRÃO
            "procedencia": "Qualquer coisa para eu poder mudar depois",
            "unidade": "Guararema", # PADRÃO
            "data_ocorrencia": datetime.now().strftime("%d/%m/%Y"), # PADRÃO
            "origem": "Interna",
            "formulario_cliente": "Não", # PADRÃO
            "descricao_detalhada": "Descrição detalhada da RNC. Fácil para mudar.",
            "porque_falha": "Explicação do porquê é uma falha. Fácil para mudar.",
            "quem_detectou": "Nome da pessoa/departamento. Fácil para mudar.",
            "como_detectada": "Descrição de como foi detectada. Fácil para mudar.",
            "item_suspeito_codigo": "COD12345",
            "item_suspeito_descricao": "Peça plástica com defeito de injeção",
            "fornecedor": "Fornecedor ABC Ltda",
            "lote": "LOTE2025-001",
            "quantas_pecas_falha": "10",
            "quantidade_total_verificada": "100",
            "quantidade_afetada": "15",
            "unidade_medida_1": "pçs", # PADRÃO
            "unidade_medida_2": "pçs", # PADRÃO
            "unidade_medida_3": "pçs", # PADRÃO
            "nome_responsavel": "IGOR FELIX DA SILVA",
            
            # Ação Isolada
            "titulo_acao_isolada": "Ação Isolada Automatizada",
            "nome_responsavel_acao_isolada": "MIRIAN"  
        }

        self.ids_campos = { # ID DOS CAMPOS DE RESPOSTA 
            # Campos de seleção (zoom)
            "reincidente": "oidzoom_8a97ecb484effd70018508398f1b16ae",
            "unidade": "oidzoom_8a97ecb484effd70018507f36f7b0820",
            "origem": "oidzoom_8a97ecb484effd70018507f462480860",
            "formulario_cliente": "oidzoom_8a97ecb484effd70018507f261ad07e9",
            "unidade_medida_1": "oidzoom_8a97ecb484effd7001850bd7ed140df5",
            "unidade_medida_2": "oidzoom_8a973a148518b31901851bae15b77736",
            "unidade_medida_3": "oidzoom_8a973a148518b31901851bad93837708",
            
            # Campos de texto
            "procedencia": "field_8a97ecb484effd70018507e8ea7004e0",
            "data_ocorrencia": "field_8a97ecb484effd70018507f3f62e0851",
            "descricao_detalhada": "field_8a97edfc84acd8730184c350d9b900cb",
            "porque_falha": "field_8a973a148518b31901851ad4e93d4626",
            "quem_detectou": "field_8a973a148518b31901851ad5b6fe4653",
            "como_detectada": "field_8a973a148518b31901851ad6019b465e",
            "item_suspeito_codigo": "field_8a97ecb484effd70018507f5bff408c8",
            "item_suspeito_descricao": "field_8a97ecb484effd70018507f8739b0966",
            "fornecedor": "field_8a97ecb484effd70018507f9c42b09a7",
            "lote": "field_8a97ecb484effd70018507fa07d809ef",
            "quantas_pecas_falha": "field_8a97ecb484effd70018507ffaa9e0b1a",
            "quantidade_total_verificada": "field_8a97ecb484effd7001850800be450b2f",
            "quantidade_afetada": "field_8a97ecb484effd70018508027f5c0b86"
        }

        if dados_excel:
            self.atualizar_dados_excel(dados_excel)

######################################## FUNÇÕES NÃO PRINCIPAIS ################################################

    def mostrar_janela_codigo(self): # FUNÇÃO QUE ABRE A JANELA PARA PEDIR O CODIGO DE LOGIN (CHATO PRA KCT SLK)
        try:
            root = tk.Tk()                          # CRIA A JANELA UTILIZANDO A BIBLIOTECA tkinter.
            root.withdraw()                         # SOME COM A JANELA EM UM PRIMEIRO MOMENTO.
            root.attributes('-topmost', True)       # FAZ A JANELA APARECER NA FRENTE DE TODAS AS OUTRAS JANELAS.
            
            codigo = simpledialog.askstring(        # ADICIONA O CONTEUDO DA JANELA E PEDE O CODIGO.
                "Código SoftExpert",
                "Digite o código de 6 dígitos:",
                parent=root
            )
            
            root.destroy()                          # FECHA A JANELA.
            
            if codigo and codigo.isdigit() and 4 <= len(codigo) <= 8: # VERIFICA SE O CODIGO É FORMADO POR SOMENTE NUMEROS E SE A QUANTIDADE DE CARACTERES CERTA.
                return codigo
            return None
                
        except:
            return None
        
    def pedir_codigo_usuario(self):                 # CHAMA A FUNÇÃO ACIMA E ARMAZENA A RESPOSTA DIGITADA. 
        self.codigo_obtido = self.mostrar_janela_codigo()

    def verificar_tela_codigo(self):                # VERIFICA SE A TELA DE CODIGO DE LOGIN AINDA ESTÁ ATIVA
        try:
            time.sleep(2)
            elementos = self.driver.find_elements(  # PROCURA OS ELEMENTOS DA TELA
                By.CSS_SELECTOR, 
                "input[placeholder*='código'], input[placeholder*='code'], input[maxlength='6']"
            )
            
            for elemento in elementos:
                try:
                    if elemento.is_displayed():     # CASO OS ELEMENTOS TENHAM SIDO ENCONTRADOS A AUTOMAÇÃO SE ENCERRA AQUI POIS O CODIGO DISPOSTO NÃO FUNCIONOU.
                        return True
                except:
                    continue
            
            return False
            
        except:
            return False
        

    def fechar_popup_alerta(self):                 # VERIFICA SE O ALERTA DE LOGIN APARECEU NA TELA E O FECHA.  
        try:
            time.sleep(1)
            
            try:
                alerta = self.driver.find_element(By.ID, 'alertConfirm')
                if alerta.is_displayed():
                    alerta.click()
                    time.sleep(1)
                    return True
            except:
                pass
            
            try:
                botoes = self.driver.find_elements(
                    By.XPATH, 
                    "//button[contains(text(), 'OK') or contains(text(), 'Ok') or contains(text(), 'Confirmar')]"
                )
                
                for botao in botoes:
                    if botao.is_displayed():
                        botao.click()
                        time.sleep(1)
                        return True
            except:
                pass
            
            return False
            
        except:
            return False
        
######################################## FUNÇÕES EXCEL ################################################
 
    @staticmethod
    def processar_arquivo_excel(caminho_excel=None, max_consecutive_errors=5):
        """Processa todas as linhas com status 'pendente' deixando o histórico.

        Em vez de remover a primeira linha (fila), marcamos a coluna status
        diretamente. Isso evita loops infinitos e preserva o arquivo completo.
        """
        if caminho_excel is None:
            caminho_excel = r"C:\Users\igor.silva\OneDrive - Steck Indústria Elétrica Ltda\Área de Trabalho\excel_enviar_se.xlsx"

        if not os.path.exists(caminho_excel):
            return 0, 0, "Arquivo não encontrado"

        total_processadas = 0
        total_erros = 0
        consecutive_errors = 0

        # única leitura; marcamos cada pendente e salvamos imediatamente
        try:
            df = pd.read_excel(caminho_excel, engine="openpyxl")
        except Exception as e:
            print(f"Erro lendo Excel: {e}")
            return 0, 1, "Falha ao abrir"

        if df.empty:
            return 0, 0, "Arquivo vazio"

        # garante coluna de status existindo caso alguém não tenha chamado antes
        if 'status' not in df.columns:
            df['status'] = 'pendente'

        pendentes = df[df['status'].astype(str).str.lower() == 'pendente']
        if pendentes.empty:
            return 0, 0, "Nenhuma linha pendente"

        for idx, linha in pendentes.iterrows():
            dados_excel = {col: linha[col] for col in df.columns}
            automacao = SoftwareExpertRNC(dados_excel=dados_excel)
            sucesso = automacao.executar()

            if sucesso:
                df.at[idx, 'status'] = 'concluido'
                total_processadas += 1
            else:
                df.at[idx, 'status'] = 'erro'
                total_erros += 1

            try:
                df.to_excel(caminho_excel, index=False, engine="openpyxl")
            except Exception as e:
                print(f"Erro ao salvar Excel: {e}")
                total_erros += 1

            sleep(1)

        return total_processadas, total_erros, "Concluído"


    def atualizar_dados_excel(self, dados_excel):   # APÓS O PROCESSAMENTO BUSCA AS INFORMAÇÕES DAS COLUNAS E AS COLOCA COMO RESPOSTAS PARA O FORMS.
        mapeamento = {                              # MAPEAMENTO DAS RESPOSTAS QUE SERÃO ALTERADAS.
            "origem": "indice_origem",
            "titulo_nc": "titulo_rnc",
            "posto": "procedencia",
            "data_ocorrencia": "data_ocorrencia",
            "desc_nc": "descricao_detalhada",
            "cod_produto": "item_suspeito_codigo",
            "cod_lote": "lote",
            "quant_pecas": "quantas_pecas_falha",
            "quant_pecas_analise": "quantidade_total_verificada",
            "quant_pecas_total": "quantidade_afetada",
            "responsavel": "nome_responsavel",
            "fornecedor": "fornecedor",
            "como_falha": "como_detectada",
            "pq_falha": "porque_falha",
            "quem_falha": "quem_detectou"
        }
        
        for coluna_excel, valor in dados_excel.items():
            
            coluna_excel_str = str(coluna_excel).lower().strip()
            valor_puro = valor
            valor_str = str(valor).strip() if not pd.isna(valor) else ""
            
            if pd.isna(valor_puro):
                continue
            
            if not valor_puro:
                continue
            
            if coluna_excel_str in mapeamento:
                chave = mapeamento[coluna_excel_str]
                
                if chave == "indice_origem":
                    try:
                        if isinstance(valor, (int, float)):
                            self.dados[chave] = int(valor)
                        else:
                            numeros = re.findall(r'\d+', valor_str)
                            if numeros:
                                self.dados[chave] = int(numeros[0])
                            else:
                                mapeamento_texto = {
                                    "auditoria": 1, "auditorias": 1,
                                    "cliente": 2,
                                    "fornecedor": 3,
                                    "fornecedor importado": 4,
                                    "internas": 5
                                }
                                texto_lower = valor_str.lower()
                                for chave_texto, num in mapeamento_texto.items():
                                    if chave_texto in texto_lower:
                                        self.dados[chave] = num
                                        break
                    except:
                        self.dados[chave] = None
                else:
                    self.dados[chave] = valor_str
        
        if not self.dados['titulo_rnc']:            # COLOCA UM TITULO PADRÃO CASO UM NÃO SEJA FORNECIDO
            self.dados['titulo_rnc'] = f"RNC - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        if not self.dados['data_ocorrencia']:       # COLOCA UMA DATA PADRÃO CASO UM NÃO SEJA FORNECIDO
            self.dados['data_ocorrencia'] = datetime.now().strftime("%d/%m/%Y")
        elif isinstance(self.dados['data_ocorrencia'], datetime) or '2025' in str(self.dados['data_ocorrencia']):
            try:
                data_obj = pd.to_datetime(self.dados['data_ocorrencia'])
                self.dados['data_ocorrencia'] = data_obj.strftime("%d/%m/%Y")
            except:
                pass
        
        if not self.dados['item_suspeito_codigo']:  # COLOCA UM CODIGO DO ITEM PADRÃO CASO UM NÃO SEJA FORNECIDO
            self.dados['item_suspeito_codigo'] = "0000"
        
        if not self.dados['nome_responsavel']:      # COLOCA UM RESPONSÁVEL PADRÃO CASO UM NÃO SEJA FORNECIDO
            self.dados['nome_responsavel'] = "Guilherme Augusto Fernandes"
        
        return True
    
    def ler_excel(self):
        # CAMINHO_EXCEL = r"C:\Users\igor.silva\Steck Indústria Elétrica Ltda\QUALIDADE GUARAREMA - Documentos\04. Backup de Códigos\4.1. Automação SoftwareExpert\excel_enviar_se.xlsx"
        CAMINHO_EXCEL = r"C:\Users\igor.silva\OneDrive - Steck Indústria Elétrica Ltda\Área de Trabalho\excel_enviar_se.xlsx"    

        if not os.path.exists(CAMINHO_EXCEL):
            print("Arquivo Excel não encontrado:", CAMINHO_EXCEL)
            return False
        try:
            df = pd.read_excel(CAMINHO_EXCEL, engine='openpyxl')
            if df.empty:
                print("Arquivo Excel está vazio:", CAMINHO_EXCEL)
                return False
            self.dados_excel = df.iloc[0].to_dict()
            return True
        except Exception as e:
            print("Erro ao ler o arquivo Excel:", e)
            return False
        
    def garantir_coluna_status(self): # GARANTE QUE A COLUNA "STATUS" EXISTA NO EXCEL PARA QUE O PROCESSAMENTO POSSA SER CONTROLADO POR ELA.
        CAMINHO_EXCEL = r"C:\Users\igor.silva\OneDrive - Steck Indústria Elétrica Ltda\Área de Trabalho\excel_enviar_se.xlsx"
        df = pd.read_excel(CAMINHO_EXCEL, engine='openpyxl')
        if 'status' not in df.columns:
            df['status'] = 'Pendente'
            df.to_excel(CAMINHO_EXCEL, index=False)
        
######################################## FUNÇÕES LOGIN ################################################

    def _encontrar_campo_codigo(self):              # ENCONTRA E RETORNA O CAMPO PARA PREENCHER O CODIGO DE LOGIN
        seletores = [
            "input[placeholder*='código']",
            "input[placeholder*='code']", 
            "input[maxlength='6']",
            "input[type='text'][maxlength='6']"
        ]
        
        for seletor in seletores:
            try:
                elemento = self.driver.find_element(By.CSS_SELECTOR, seletor)
                if elemento.is_displayed():
                    return elemento
            except:
                continue
        
        return None
    
    def _encontrar_botao_confirmar(self):           # ENCONTRA UM BOTÃO DE CONFIRMAR NA TELA E CLICA NELE.
        seletores = [
            "button[type='submit']",
            "input[type='submit']",
            "//button[contains(text(), 'Confirmar')]",
            "//button[contains(text(), 'Verificar')]"
        ]
        
        for i, seletor in enumerate(seletores):
            try:
                if i < 2:
                    elemento = self.driver.find_element(By.CSS_SELECTOR, seletor)
                else:
                    elemento = self.driver.find_element(By.XPATH, seletor)
                
                if elemento.is_displayed():
                    return elemento
            except:
                continue
        
        return None
    
    def inserir_codigo(self, codigo):       # DÁ INPUT DO CODIGO DE LOGIN NO CAMPO 
        try:
            campo = self._encontrar_campo_codigo()
            if not campo:
                return False
            
            campo.clear()
            time.sleep(0.5)
            campo.send_keys(codigo)
            time.sleep(1)
            
            botao = self._encontrar_botao_confirmar()
            if botao:
                try:
                    botao.click()
                except:
                    self.driver.execute_script("arguments[0].click();", botao)
            else:
                campo.send_keys(Keys.RETURN)
            
            time.sleep(5)
            self.fechar_popup_alerta()
            
            return True
            
        except:
            return False
    
    def login(self):                        # FAZ LOGIN UÉ
        try:
            self.driver.get('https://steck-test.softexpert.com/softexpert/login?page=home') # ACESSA A PAGINA DE LOGIN.
            time.sleep(3)
            
            self.driver.find_element(By.ID, 'user').send_keys(self.dados["usuario"])        # DA INPUT NO USERNAME.
            self.driver.find_element(By.ID, 'password').send_keys(self.dados["senha"])      # DA INPUT NA SENHA.
            self.driver.find_element(By.ID, 'loginButton').click()                          # CLICA NO BOTÃO DE LOGIN.
            
            time.sleep(4)
            
            if self.verificar_tela_codigo():                                                # CHAMA FUNÇÃO DO CODIGO DE LOGIN
                self.pedir_codigo_usuario() # PEDE CODIGO PARA O USUÁRIO DIGITAR.
                self.fechar_popup_alerta()
                

                
                timeout = 300
                inicio = time.time()
                
                while time.time() - inicio < timeout:                                       # VERIFICA SE O TEMPO LIMITE NÃO ACABOU.
                    if self.codigo_obtido is not None:                                      # VERIFICA SE O CODIGO NÃO ESTÁ VAZIO.
                        break
                    time.sleep(1)
                
                if not self.codigo_obtido:
                    return False
                
                sucesso = self.inserir_codigo(self.codigo_obtido)                           # CHAMA FUNÇÃO DE INSERIR O CODIGO.
                
                if sucesso:
                    time.sleep(3)
                    self.fechar_popup_alerta()                                              # FECHA O POP-UP.
                    return True
                else:
                    return False
            
            else:
                self.fechar_popup_alerta()
                return True
                
        except Exception as e:
            return False
        
######################################## FUNÇÕES NAVEGAÇÃO ################################################

    def navegar_para_formulario(self):  # NAVEGA ATÉ O FORMULÁRIO 
            try:
                self.driver.get('https://steck-test.softexpert.com/softexpert/workspace?page=108084,104') # ACESSA O LINK DIRETO DA PAGINA PARA ABERTURA DE RNC.
                time.sleep(5)
                
                try:
                    btn_busca = self.driver.find_element(By.XPATH, "//button[@data-test-selector='filterSearchBtn']")
                    btn_busca.click() # CLICA NO BOTÃO DE FILTRO PARA FAZER AS OPÇÕES APARECEREM (SIM EU SEI QUE ISSO É MUITO IDIOTA).
                    time.sleep(2)
                except:
                    pass
                
                try:
                    img_ferramenta = self.driver.find_element(By.XPATH, "//img[contains(@src, 'ferramenta-de-apoio.png')]")
                    img_ferramenta.click() # CLICA NA IMAGEM PARA CONSEGUIR CLICAR NO LOCAL CERTO PARA AI SIM ABRIR O BAGULHO DO FORM (ISSO AQUI É FEIO).
                    time.sleep(2)
                except:
                    pass
                
                try:
                    textarea = self.driver.find_element(By.XPATH, "//textarea[@data-test-id='84']")
                    textarea.send_keys(self.dados["titulo_rnc"]) # CLICA NA ÁREA DE TEXTO PARA COLOCAR O TITULO.
                    time.sleep(2)
                except:
                    pass
                
                try:
                    btn_iniciar = self.driver.find_element(By.XPATH, "//span[text()='Iniciar']")
                    btn_iniciar.click()     # CLICA EM INICIAR E ABRE O FORMS.
                    time.sleep(8)
                    return True
                except:
                    return False
                
            except Exception as e:
                return False
            
    def acessar_formulario(self):   # FUNÇÃO PARA ACESSAR O FORMS 
        try:
            if not self.janela_principal:
                self.janela_principal = self.driver.current_window_handle
                print(f"Janela principal guardada: {self.janela_principal}")
        
            # Aguarda nova janela abrir
                self.wait.until(lambda d: len(d.window_handles) > 1)
                time.sleep(2)
            
            for janela in self.driver.window_handles:              # VERIFICA SE A JANELA ORIGINAL É DIFERENTE DA ATUAL.
                if janela != self.janela_principal:
                    self.driver.switch_to.window(janela)
                    self.form_window_handle = janela               # TROCA DE JANELA.
                    print(f"Janela do formulário: {self.form_window_handle}")
                    break
            
            time.sleep(3)
            
            try:
                iframe1 = self.wait.until(EC.presence_of_element_located((By.ID, "ribbonFrame")))  # TENTA ACESSAR O PRIMEIRO IFRAME.
                self.driver.switch_to.frame(iframe1)
                time.sleep(2)
            except Exception as e:
                pass
            
            try:
                iframe2 = self.wait.until(EC.presence_of_element_located(                          # TENTA ACESSAR O SEGUNDO IFRAME.
                    (By.XPATH, "//iframe[contains(@id, 'frame_form_')]")
                ))
                self.driver.switch_to.frame(iframe2)
                time.sleep(3)
                return True
            except Exception as e:
                return False
            
        except Exception as e:
            return False
    
    def restaurar_contexto_formulario(self):
        """Restaura o contexto dos iframes após voltar para o formulário"""
        try:
            self.driver.switch_to.default_content()
            time.sleep(1)
            
            iframe1 = self.wait.until(EC.presence_of_element_located((By.ID, "ribbonFrame")))
            self.driver.switch_to.frame(iframe1)
            time.sleep(1)
            
            iframe2 = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, "//iframe[contains(@id, 'frame_form_')]")
            ))
            self.driver.switch_to.frame(iframe2)
            time.sleep(1)
            return True
        except Exception as e:
            return False 
        
######################################## FUNÇÕES PREENCHER FORMS ################################################

    def preencher_campos_sem_responsavel(self):   
            try:
                campos_select = {  # PUXA TODOS OS ID DOS CAMPOS E VINCULA UMA RESPOSTA A ELES.
                    self.ids_campos["reincidente"]: self.dados["reincidente"],
                    self.ids_campos["unidade"]: self.dados["unidade"],
                    self.ids_campos["origem"]: self.dados["origem"],
                    self.ids_campos["formulario_cliente"]: self.dados["formulario_cliente"],
                    self.ids_campos["unidade_medida_1"]: self.dados["unidade_medida_1"],
                    self.ids_campos["unidade_medida_2"]: self.dados["unidade_medida_2"],
                    self.ids_campos["unidade_medida_3"]: self.dados["unidade_medida_3"]
                }
                
                campos_texto = {
                    self.ids_campos["procedencia"]: self.dados["procedencia"],
                    self.ids_campos["data_ocorrencia"]: self.dados["data_ocorrencia"],
                    self.ids_campos["descricao_detalhada"]: self.dados["descricao_detalhada"],
                    self.ids_campos["porque_falha"]: self.dados["porque_falha"],
                    self.ids_campos["quem_detectou"]: self.dados["quem_detectou"],
                    self.ids_campos["como_detectada"]: self.dados["como_detectada"],
                    self.ids_campos["item_suspeito_codigo"]: self.dados["item_suspeito_codigo"],
                    self.ids_campos["item_suspeito_descricao"]: self.dados["item_suspeito_descricao"],
                    self.ids_campos["fornecedor"]: self.dados["fornecedor"],
                    self.ids_campos["lote"]: self.dados["lote"],
                    self.ids_campos["quantas_pecas_falha"]: self.dados["quantas_pecas_falha"],
                    self.ids_campos["quantidade_total_verificada"]: self.dados["quantidade_total_verificada"],
                    self.ids_campos["quantidade_afetada"]: self.dados["quantidade_afetada"]
                }
                
                for campo_id, valor in campos_select.items():   # REPETE ESSE PROCESSO PARA CADA UM DOS CAMPOS DO TIPO "SELECT".
                    try:
                        campo = self.driver.find_element(By.ID, campo_id)   # ENCONTRA O CAMPO PELO ID.
                        select = Select(campo)  # SELECIONA O CAMPO
                        for option in select.options:   # PEGA A OPÇÃO A SER SELECIONADA. ENCONTRA ELA NO DROPDOWN E A SELECIONA.
                            if valor.lower() in option.text.lower():
                                select.select_by_visible_text(option.text)
                                break
                    except:
                        pass
                    time.sleep(0.3)
                
                for campo_id, valor in campos_texto.items():   # REPETE ESSE PROCESSO PARA CADA UM DOS CAMPOS DO TIPO "TEXT".
                    try:
                        campo = self.driver.find_element(By.ID, campo_id)   # ENCONTRA O CAMPO PELO ID.
                        campo.clear()   # LIMPA O CAMPO.
                        campo.send_keys(valor)  # DA INPUT NA RESPOSTA.
                    except:
                        pass
                    time.sleep(0.3)
                
                return True
                
            except Exception as e:
                return False
            
######################################## FUNÇÕES SELECIONAR O RESPONSÁVEL ################################################

    def selecionar_responsavel(self): # VERIFICAR SE FUNCIONA :)
        try:
            # USA A REFERÊNCIA DA PAGINA.
            janela_formulario = self.form_window_handle
            
            time.sleep(2)
            
            try:
                zoom_icon = self.driver.find_element(By.XPATH, "//img[contains(@src, 'zoom.gif')]")
                zoom_icon.click()
                time.sleep(3)
            except Exception as e:
                return False
            
            time.sleep(2)
            janelas_agora = self.driver.window_handles
            
            if len(janelas_agora) <= 1:
                return False
            
            nova_janela = janelas_agora[-1]
            self.driver.switch_to.window(nova_janela)
            time.sleep(2)
            
            try:
                campo_pesquisa = self.driver.find_element(By.ID, "searchtext")
                campo_pesquisa.send_keys(self.dados["nome_responsavel"])
                campo_pesquisa.send_keys(Keys.RETURN)
                time.sleep(4)
            except Exception as e:
                return False
            
            try:
                self.driver.find_element(
                    By.XPATH, "//table//tr[contains(@class, 'row')][1]//td[contains(@onclick, 'show_hide')]"
                ).click()
                time.sleep(2)
            except:
                pass
            
            try:
                try:
                    self.driver.find_element(By.ID, "btnsave_exit").click()
                except:
                    try:
                        self.driver.find_element(
                            By.XPATH, "//button[.//label[contains(text(), 'Salvar e sair')]]"
                        ).click()
                    except:
                        try:
                            self.driver.find_element(
                                By.XPATH, "//button[.//img[contains(@src, 'save_exit.png')]]"
                            ).click()
                        except:
                            pass
                
                for i in range(10):
                    if len(self.driver.window_handles) == 1:
                        break
                    time.sleep(1)
            except Exception as e:
                pass
            
            # VOLTA PARA A JANELA DO FORMULÁRIO USANDO A REFERÊNCIA
            self.driver.switch_to.window(janela_formulario)
            self.restaurar_contexto_formulario()
            
            return True  
            
        except Exception as e:
            # SE DER PAU TENTA VOLTAR NA JANELA ANTERIOR COM A REF.
            try:
                self.driver.switch_to.window(self.form_window_handle)
                self.restaurar_contexto_formulario()
            except:
                pass
            return False  

######################################## FUNÇÕES EXECUÇÃO ################################################

    def executar(self):
        """Método principal que executa toda a automação"""
        try:
            print("Iniciando login...")
            if not self.login():
                print("Falha no login")
                return False

            caminho = r"C:\Users\igor.silva\OneDrive - Steck Indústria Elétrica Ltda\Área de Trabalho\excel_enviar_se.xlsx"
            
            # CHAMA A FUNÇÃO PARA VER SE A COLUNA STATUS EXISTE.
            try:
                self.garantir_coluna_status()
            except Exception:
                # SE DER PAU ELE DEFINE A COLUNA AGORA
                df_temp = pd.read_excel(caminho, engine='openpyxl')
                if 'status' not in df_temp.columns:
                    df_temp['status'] = 'pendente'
                    df_temp.to_excel(caminho, index=False, engine='openpyxl')
            
            df = pd.read_excel(caminho, engine='openpyxl')
            
            # ARRUMA LINHA PARA PEGAR APENAS AS LINHAS QUE ESTÃO COM O STATUS "PENDENTE" PARA PROCESSAR.
            linhas_pendentes = df[df['status'] == 'pendente']
            if len(linhas_pendentes) == 0:
                print("Nada pendente")
                return True

            self.janela_principal = self.driver.current_window_handle
            print(f"Janela principal guardada: {self.janela_principal}")
            print("Login OK")

            for index, linha in linhas_pendentes.iterrows():
                self.atualizar_dados_excel(linha.to_dict())
                
                print("Navegando para formulário...")
                if not self.navegar_para_formulario():
                    print("Falha na navegação")
                    return False
                print("Navegação OK")

                try:
                    print(f"\n--- Processando linha {index} ---")

                    print("Acessando formulário...")
                    if not self.acessar_formulario():
                        print("Falha ao acessar processar linha")
                        df.loc[index, 'status'] = 'erro'
                        continue
                    print("Acesso OK")
                    
                    print("Preenchendo campos básicos usando dados da instância...")
                    if not self.preencher_campos_sem_responsavel():
                        print("Alguns campos não foram preenchidos")
                    print("Campos básicos OK")

                    print("Selecionando responsável principal...")
                    if not self.selecionar_responsavel():
                        print("Falha no responsável principal, aguardando manual...")
                        sleep(3)
                    print("Responsável principal OK")

                    try:
                        print("Procurando botão Executar...")

                        self.driver.switch_to.default_content()
                        time.sleep(1)

                        botao_executar = self.wait.until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//button[.//*[contains(text(),'Executar')]]")
                            )
                        )

                        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_executar)
                        time.sleep(1)

                        botao_executar.click()

                        print("Botão Executar clicado com sucesso!")

                    except Exception as e:
                        print("Erro ao clicar no botão Executar:", e)

                    # ATUALIZAR O STATUS DA LINHA E GRAVAR NO MESMO ARQUIVO
                    df.loc[index, 'status'] = 'concluido'
                    try:
                        df.to_excel(caminho, index=False, engine='openpyxl')
                    except Exception as e:
                        print(f"Erro ao salvar Excel: {e}")

                    print("Processo concluído")
                    time.sleep(2)

                    # AGUARDA O FORMULÁRIO FECHAR
                    print("Aguardando formulário fechar...")
                    time.sleep(5)

                    # VOLTA PARA A JANELA PRINCIPAL
                    try:
                        # Mostra quantas janelas existem
                        print(f"Janelas abertas: {self.driver.window_handles}")

                        # Volta para a janela principal
                        if self.janela_principal in self.driver.window_handles:
                            self.driver.switch_to.window(self.janela_principal)
                            print("Voltou para janela principal")
                        else:
                            # Se perdeu referência, pega a primeira
                            self.driver.switch_to.window(self.driver.window_handles[0])
                            self.janela_principal = self.driver.current_window_handle
                            print("Reconfigurou janela principal")
                    except Exception as e:
                        print(f"Erro ao voltar para janela principal: {e}")

                except Exception as e:
                    print(f"Erro ao processar linha {index}: {e}")
                    return False

            return True

        except Exception as e:
            print(f"Erro crítico: {e}")
            return False
        finally:
            try:
                self.driver.quit()
                sys.exit(0)
            except:
                pass

if __name__ == "__main__":
    processadas, erros, msg = SoftwareExpertRNC.processar_arquivo_excel()
    print(f"Processadas: {processadas}, Erros: {erros}, Mensagem: {msg}")

    




