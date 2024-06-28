import os
import pandas as pd
import openpyxl

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
import datetime


# --- ORGANIZANDO DIRETORIO ---
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
DOCS_DIR = os.path.join(BASE_DIR, 'docs')

# --- ABRINDO A BASE DE DADOS ---
file_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR)][0]
df = pd.read_excel(file_path)
print(df.head())
print(df.shape)
lista_arquivos_sem_extensao = df.iloc[:,0].tolist()

# --- INICIALIZAR O CHROME ---
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 20
wait = WebDriverWait(driver, timeout)

# --- USAR AS CREDENCIAIS PARA ENTRAR NA PLATAFORMA ---
driver.get('https://performaqca.seven.adv.br/')
print('aguardando 30 segundos')
sleep(30)

# lista_arquivos_sem_extensao[:2]
print('iniciando loop')
caso_sucesso = []
caso_fracasso = []
for indice, id in enumerate(lista_arquivos_sem_extensao):
    driver.get(f'https://performaqca.seven.adv.br/compromisso/details/{id}')
    sleep(5)
    try:
        # --- INSERIR OS DOCUMENTOS ---
        # Clicar em "Documentos"
        botao_documentos = driver.find_element(By.XPATH, '//ul[@class="nav nav-tabs tab-red"]//li//a[@href="#box-documentos"]')
        botao_documentos.click()
        sleep(3)

        # Clicar em "Ações"
        botao_acoes = driver.find_elements(By.XPATH, '//div[@class="btn-group btn-group-xs pull-right"]//button[@data-toggle="dropdown"]')[-1]
        botao_acoes.click()
        sleep(1)

        # Clicar em "Novo Documento"
        novo_documento_botao = driver.find_element(By.XPATH, '//*[@id="buttonNewDocumento"]')
        novo_documento_botao.click()
        sleep(3)
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)

        # Inserir o caminho do arquivo
        caminho_arquivo = os.path.join(DOCS_DIR, str(id) + '.pdf')
        substabelecimento_arquivo = rf'C:\Users\victoramarante\Documents\9.automacao_diogo\substabelecimento.pdf'
        upload_arquivo = driver.find_element(By.XPATH, '//*[@id="ufile"]')
        upload_arquivo.send_keys(caminho_arquivo)
        upload_arquivo.send_keys(substabelecimento_arquivo)
        sleep(1)

        # PARTE 1 ---------------------------------------------------------------------------------------------------------
        ## Selecionar as opções necessárias
        # Clicar em "Tipo de Documento"
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/button')
        tipo_documento.click()
        # Escrever o "Tipo de Documento"
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        # Clicar em "Documento Interno = Sim"
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        # Clicar em "Foi utilizado Legal Design = Sim"
        legal_design_sim = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[2]/div/ins')[-1]
        legal_design_sim.click()
        # Clicar em "Integrar protocolo? = Não"
        integrar_protocolo_nao = driver.find_elements(By.XPATH, '//*[@id="radio0"]/div/label[1]/div/ins')[1]
        integrar_protocolo_nao.click()
        # Digitar a data atual no formato brasileiro
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[1]/td[2]/div[3]/div/input')
        data_atual.send_keys(datetime.date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)

        # PARTE 2 ---------------------------------------------------------------------------------------------------------
        # Clicar em "Tipo de Documento"
        tipo_documento = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/button')
        tipo_documento.click()
        # Escrever o "Tipo de Documento"
        documento_input = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[2]/div/div/div[1]/input')
        documento_input.send_keys('Protocolo - Habilitação', Keys.ENTER)
        # Clicar em "Documento Interno = Sim"
        doc_interno_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[0]
        doc_interno_sim.click()
        # Clicar em "Foi utilizado Legal Design = Sim"
        legal_design_sim = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[2]/div/ins')[-1]
        legal_design_sim.click()
        # Clicar em "Integrar protocolo? = Não"
        integrar_protocolo_nao = driver.find_elements(By.XPATH, '//*[@id="radio1"]/div/label[1]/div/ins')[1]
        integrar_protocolo_nao.click()
        # Digitar a data atual no formato brasileiro
        data_atual = driver.find_element(By.XPATH, '//*[@id="gridDocTemp"]/table/tbody/tr[2]/td[2]/div[3]/div/input')
        data_atual.send_keys(datetime.date.today().strftime("%d/%m/%Y"), Keys.ENTER)
        sleep(1)
        
        botao_salvar = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]')
        botao_salvar.click()
        # ----------------------------------------------------------------------------------------------------------------------
        sleep(3)
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '//*[@id="conclusaoRevisaoForm"]/div[1]/div/label[2]/div/ins')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)
        
        # Clicar em "Aprova revisão? = Sim"
        aprova_revisao_sim = driver.find_element(By.XPATH, '//*[@id="conclusaoRevisaoForm"]/div[1]/div/label[2]/div/ins')
        aprova_revisao_sim.click()
        sleep(1)
               
        botao_auditar = driver.find_element(By.XPATH, '//*[@id="buttonConcluirRevisao"]')
        botao_auditar.click()
        sleep(3)
        print(f'{indice+1}: ID {id} preenchido')
        caso_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
    except TimeoutException as e:
        print(f'{indice+1}: Erro ao processar o processo {id}')
        caso_fracasso.append({'Processo': id, 'Status': f'Erro: {str(e)}'})
    except Exception as e:
        print(f'{indice+1}: Erro inesperado ao processar o processo {id}')
        caso_fracasso.append({'Processo': id, 'Status': f'Erro inesperado: {str(e)}'})

driver.quit()

# exportando as bases para controle do usuario
df_sucesso = pd.DataFrame(caso_sucesso)
df_fracasso = pd.DataFrame(caso_fracasso)
df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

sleep(1)

print('Processo finalizado.')