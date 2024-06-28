from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas as pd
import openpyxl
import datetime
import calendar
import locale
import os
from collections import defaultdict


locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


class ProdutividadeLexio:
    '''
    Essa classe serve para realizarmos o processo de preenchimento dos valores de acordo com o prestador e a sua unidade. Essa é uma automacao de processos roboticos que visa diminuir o tempo de execucao do colaborador,
    possibilitando que o mesmo possa destinar mais tempo do seu horario para questoes mais estrategicas.
    '''
    def __init__(self, url, login, senha, dataframe):
        self.service = Service(ChromeDriverManager().install())
        self.options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(service=self.service, options=self.options)
        self.driver.maximize_window()
        self.timeout = 20
        self.wait = WebDriverWait(self.driver, self.timeout)
        self.url = url
        self.login = login
        self.senha = senha
        self.df = dataframe
        
    def entrar_na_lexio(self):
        # Entramos na url da lexio
        self.driver.get(self.url)
        sleep(3)

        # Inserimos o email do usuario
        username = self.login
        username_input = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//form//input[@autocomplete='username']")))
        username_input.send_keys(username)
        sleep(1)

        # Inserimos a senha do usuario
        senha = self.senha
        senha_input = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//form//input[@name='_password']")))
        senha_input.send_keys(senha)
        sleep(1)

        # Clicamos no botao de confirmar o login para entrar na plataforma
        login_confirmar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='botoesSubmit']//input")))
        login_confirmar.click()
        sleep(2)
        
        
    def entrar_pasta_de_trabalho(self):
        # Entramos diretamente na pasta de trabalho atraves da url
        self.driver.get('https://app.lexio.legal/group/')
    
    
    def procurar_prestador(self, nome_do_prestador):
        # Inserimos o nome do prestador para acha-lo na pasta de trabalho
        self.nome_do_prestador = nome_do_prestador
        nome_prestador = self.driver.find_element(By.XPATH, "/html/body/div[2]/div[3]/div[2]/div/div/div/div/div/div/input")
        nome_prestador.send_keys(self.nome_do_prestador)
        sleep(2)
        pastas_grupos = self.driver.find_element(By.XPATH, "//details//summary//div")
        pastas_grupos.click()
        sleep(1)
        selecionar_colaborador = self.driver.find_elements(By.XPATH, "//details//div//ul[@class='kGibhRl6Doson_63wUPD__folder-list']//li//a")
        if selecionar_colaborador[0].text == self.nome_do_prestador:
            selecionar_colaborador[0].click()
            sleep(2)
          
          
    def alterar_titulo_documento(self, nome_do_prestador):
        self.nome_do_prestador = nome_do_prestador
        data_atual = datetime.date.today()
        ultimo_mes = data_atual.replace(day=1) - datetime.timedelta(days=1)
        primeiro_dia_mes_anterior = ultimo_mes.replace(day=1)
        ultimo_dia_mes_anterior = primeiro_dia_mes_anterior.replace(day=calendar.monthrange(primeiro_dia_mes_anterior.year, primeiro_dia_mes_anterior.month)[1])
        primeiro_dia_formatado = primeiro_dia_mes_anterior.strftime('%d.%m')
        ultimo_dia_formatado = ultimo_dia_mes_anterior.strftime('%d.%m')
        return (f"Produtividade de {self.nome_do_prestador} de {str(primeiro_dia_formatado)} até {str(ultimo_dia_formatado)}")
    
    
    def alterar_produtividade(self):
        # try:
        # Clicamos na pasta de produtividade
        produtividade = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='basBURnp9_MxdYpbhW_9__contrato']//a//span[text()='Produtividade']")))
        produtividade.click()
        sleep(1)
        
        # Clicamos no botao de editar o titulo do documento
        botao_editar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//small//a[@data-original-title='Renomear contrato']")))
        botao_editar.click()
        sleep(1)
        
        # Limpamos o nome anterior e colocamos o retorno da funcao alterar_titulo_documento()
        renomear_produtividade = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='nome_contrato']")))
        renomear_produtividade.clear()
        sleep(.5)
        # Aplicamos a funcao para retornar o titulo e escrevemos para substituir o nome anterior
        novo_titulo = self.alterar_titulo_documento(self.nome_do_prestador)
        renomear_produtividade.send_keys(novo_titulo)
        sleep(.5)
        
        # Salvamos as alteracoes feitas no nome do titulo
        botao_confirmar = self.driver.find_element(By.XPATH, "//*[@id='confirmarContrato']")
        botao_confirmar.click()
        sleep(2)
        # except TimeoutException:
        #     print("Tempo limite excedido ao esperar por um elemento")
        # except Exception as e:
        #     print(f"Ocorreu uma exceção: {e}")
        #     pass
    
    
    def criar_fluxo(self):
        # try:
        # Clicamos no botao "+ novo fluxo" para iniciar um novo fluxo
        botao_novo_fluxo = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='loading-page']/div[5]/div/div[2]/div[6]/div[1]/div/a")))
        botao_novo_fluxo.click()
        sleep(1)
        
        # Clicamos no botao para alterar o tipo de template
        botao_template = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='set-radio-list']/label[2]/span")))
        botao_template.click()
        sleep(1)
        
        # Clicamos no botao para criar o fluxo
        criar_fluxo = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='form-createWorkflow']/div[2]/div[2]/div/input")))
        criar_fluxo.click()
        sleep(5)
        # except TimeoutException:
        #     print("Tempo limite excedido ao esperar por um elemento")
        # except Exception as e:
        #     print(f"Ocorreu uma exceção: {e}")
        #     pass
    
    
    def criar_lista_ufs_ordenadas(self):
        listaUfs = self.driver.find_elements(By.XPATH, "//input[contains(@value, 'Valor CNPJ')]")
        listaUfs_nomes = [item.get_attribute('value').split(' ')[-1] for item in listaUfs]        
        return listaUfs_nomes
    
    
    def digitar_uf(self):
        self.lista_nomes_ufs = list(self.criar_lista_ufs_ordenadas())
        for uf in self.lista_nomes_ufs:
            # Encontre o elemento <input> com o texto "Valor CNPJ UF"
            elemento_cnpj_uf = self.driver.find_element(By.XPATH, f"//input[starts-with(@value, 'Valor CNPJ {uf}')]")
            # Encontre o próximo <input> irmão
            elemento_seguinte = elemento_cnpj_uf.find_element(By.XPATH, "./following-sibling::input[1]")
            # Limpe o campo do próximo <input>
            elemento_seguinte.clear()
            elemento_seguinte.send_keys(str(uf))
    
    
    def digitar_valores(self, dataframe):
        while self.lista_nomes_ufs:
            uf = self.lista_nomes_ufs[0]
            print(f'inserindo informacao da uf {uf}')
            elemento_valor_bruto = self.driver.find_element(By.XPATH, f"//input[@value = '{uf}']")
            valor_bruto_correspondente = dataframe[dataframe["Unidade"] == uf]['Valor Bruto'].values
            if len(valor_bruto_correspondente) > 0:
                print(valor_bruto_correspondente[0])
                valor_bruto_formatado = locale.currency(valor_bruto_correspondente[0], grouping=True, symbol=True)
                elemento_valor_bruto.clear()
                elemento_valor_bruto.send_keys(str(valor_bruto_formatado))
                self.lista_nomes_ufs.remove(uf)
            else:
                elemento_valor_bruto.clear()
                elemento_valor_bruto.send_keys('R$ 0,00')
                self.lista_nomes_ufs.remove(uf)
                
            sleep(0.8)
    
    
    def anexar_documento(self, documento_para_anexar):
        # Clicamos no botao de anexo
        _ = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[1]/div[2]/span[4]")))
        anexo = self.driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[1]/div[2]/span[4]")
        anexo.click()
        sleep(0.8)
        
        # Clicamos para anexar um novo documento
        _ = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button")))
        anexo_novo = self.driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button")
        anexo_novo.click()
        sleep(0.8)
        
        # Enviamos o diretorio do arquivo correspondente para ser salvo dentro do sistema
        _ = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='file']")))
        arquivo = self.driver.find_element(By.XPATH, "//input[@type='file']")
        arquivo.send_keys(rf"{documento_para_anexar}")
        sleep(0.8)
        # Escrevemos uma descrição para o anexo
        _ = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/main/textarea")))
        descricao = self.driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/main/textarea")
        descricao.send_keys("Segue produtividade para validação.")
        sleep(0.8)
        
        # Subimos esse arquivo em anexo
        _ = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button[2]")))
        subir_botao = self.driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button[2]")
        subir_botao.click()
        sleep(0.8)
        
        # Fechamos a janela referente ao anexo
        _ = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "svg[data-testid='CancelIcon']")))
        botao_fechar = self.driver.find_element(By.CSS_SELECTOR, "svg[data-testid='CancelIcon']")
        botao_fechar.click()
        sleep(1)
    
    
    def identificar_documentos_para_anexar(self, dataframe):
        # Organizando os diretórios
        BASE_DIR = os.getcwd()
        PRODUTIVIDADE_DIR = os.path.join(BASE_DIR, 'Produtividade')

        # Criar uma lista com todos os arquivos que podem ser anexados
        files = os.listdir(PRODUTIVIDADE_DIR)

        nomes = list(set(dataframe['Correspondente']))

        arquivos_por_nome = defaultdict(list)

        for file in files:
            for nome in nomes:
                if nome in file:
                    caminho_completo = os.path.join(PRODUTIVIDADE_DIR, file)
                    arquivos_por_nome[nome].append(caminho_completo)

        quantidade_arquivos_por_nome = {nome: {'quantidade': len(arquivos), 'arquivos': arquivos} for nome, arquivos in arquivos_por_nome.items()}

        for arquivo in quantidade_arquivos_por_nome[str(self.nome_do_prestador)]['arquivos']:
            print(f'Inserindo o {arquivo}')
            self.anexar_documento(arquivo)
        
        print('Processo de anexar documentos finalizado.')
    
    
    def salvar_fluxo(self):
        # try:
        botao_salvar = self.wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[3]/div[2]/div/div/div/div[2]/div/button[2]")))
        sleep(1)
        botao_salvar.click()
        botao_iniciar_fluxo = self.wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div[6]/button[3]")))
        sleep(1)
        botao_iniciar_fluxo.click()
        sleep(5)
        # except TimeoutException:
        #     print("Tempo limite excedido ao esperar por um elemento")
        # except Exception as e:
        #     print(f"Ocorreu uma exceção: {e}")
        #     pass

    
    def fluxo_preenchimento(self):
        # Entramos na plataforma
        self.entrar_na_lexio()
        sleep(1)
        
        print("Iniciando o preenchimento para cada colaborador selecionado")
        
        # Colocar um try-except para verificar posteriormente se deu erro ou não. Todas essas informações salvar em uma planilha para exportar e ter maior controle daquilo que foi feito.
        controle_sucesso = []
        controle_falha = []
        # --- PARA CADA PRESTADOR --- utilizando estrutura de repeticao
        try:
            for indice, prestador in enumerate(self.df['Correspondente'].unique()):
                try:
                    # Selecionamos os dados referentes a um prestador especifico
                    dados = self.df[self.df['Correspondente'] == prestador]
                    # Fazemos um agrupamento para somar o valor bruto de acordo com o prestador e a sua unidade (uf)
                    dados_group = dados.groupby(['Correspondente', 'Unidade'])['Valor Bruto'].sum()
                    
                    # Transformamos em um dataframe e resetamos os indices para ficar de acordo com o formato que trabalhamos
                    df_final = pd.DataFrame(dados_group).reset_index()
                    print(f"Iniciando com o prestador: {indice+1} - {prestador}")
                    print(df_final)
                    sleep(0.5)
                    # Entramos na aba de pasta de trabalho
                    self.entrar_pasta_de_trabalho()
                    sleep(0.8)
                    # Procuramos o nome do prestador e selecionamos
                    self.procurar_prestador(nome_do_prestador = prestador)
                    sleep(0.8)
                    # Entramos na pasta de trabalho que se chama "Produtividade" e alteramos o nome da pasta
                    self.alterar_produtividade()
                    sleep(0.8)
                    # Iniciamos um novo fluxo
                    self.criar_fluxo()
                    sleep(0.8)
                    # Inserimos o documento para ser anexado
                    self.identificar_documentos_para_anexar(dataframe = self.df)
                    sleep(0.8)
                    # Inserimos os dados de pagamento nos locais corretos
                    ## Precisamos colocar as UFs
                    self.digitar_uf()
                    sleep(0.8)
                    ## Precisamos colocar os valores de acordo com as UFs especificas
                    self.digitar_valores(dataframe = df_final)
                    sleep(0.8)
                    # Salvamos todo o procedimento
                    self.salvar_fluxo()
                    print(f'{indice+1} - Prestador {prestador} finalizado')
                    controle_sucesso.append({'Prestador': prestador, 'Status': 'Sucesso'})
                except TimeoutException as e:
                    print(f'Erro ao processar o prestador {prestador}')
                    controle_falha.append({'Prestador': prestador, 'Status': f'Erro: {str(e)}'})
                except Exception as e:
                    print(f'Erro inesperado ao processar o prestador {indice}')
                    controle_falha.append({'Prestador': prestador, 'Status': f'Erro inesperado: {str(e)}'})
                    continue
        finally:
            df_controle_sucesso = pd.DataFrame(controle_sucesso)
            df_controle_falha = pd.DataFrame(controle_falha)
            
            df_controle_sucesso.to_excel('planilha_controle_sucesso.xlsx', index=False)
            df_controle_falha.to_excel('planilha_controle_falha.xlsx', index=False)
            

    def fechar_chrome(self):
        self.driver.quit()
            

def carregar_base_dados(caminho_diretorio):
    lista_arquivos = os.listdir(caminho_diretorio)
    for arquivo in lista_arquivos:
        if arquivo.endswith(".xlsx"):
            caminho_arquivo = os.path.join(caminho_diretorio, arquivo)
            print(caminho_arquivo)
            df = pd.read_excel(caminho_arquivo, engine="openpyxl", sheet_name="main")
    return df
            
# if __name__ == '__main__':
    # Pegar planilha excel independente do nome mas que esteja no formato excel
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
df = carregar_base_dados(DATA_DIR)
print('Base de dados carregada', df)
print(df.columns)
lexio = ProdutividadeLexio(url="https://app.lexio.legal/login",
                            login="daianasilva@queirozcavalcanti.adv.br", # ALTERAR PARA O SEU E-MAIL
                            senha="Qca@86704977", # ALTERAR PARA A SUA SENHA
                            dataframe=df)
sleep(3)    
print("Iniciando o fluxo geral")    
lexio.fluxo_preenchimento()
sleep(3)    
lexio.fechar_chrome()    
print("O procedimento foi executado!")
