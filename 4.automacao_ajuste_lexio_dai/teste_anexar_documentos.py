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

service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 20
wait = WebDriverWait(driver, timeout)

def anexar_documento(documento_para_anexar):
    # Clicamos no botao de anexo
    _ = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[1]/div[2]/span[4]")))
    anexo = driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[1]/div[2]/span[4]")
    anexo.click()
    sleep(0.8)
    
    # Clicamos para anexar um novo documento
    _ = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button")))
    anexo_novo = driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button")
    anexo_novo.click()
    sleep(0.8)
    
    # Enviamos o diretorio do arquivo correspondente para ser salvo dentro do sistema
    _ = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='file']")))
    arquivo = driver.find_element(By.XPATH, "//input[@type='file']")
    arquivo.send_keys(rf"{documento_para_anexar}")
    sleep(0.8)
    # Escrevemos uma descrição para o anexo
    _ = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/main/textarea")))
    descricao = driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/main/textarea")
    descricao.send_keys("Segue produtividade para validação.")
    sleep(0.8)
    
    # Subimos esse arquivo em anexo
    _ = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button[2]")))
    subir_botao = driver.find_element(By.XPATH, "//*[@id='workflowAdmin']/div/div[1]/div/div[3]/div/div/footer/button[2]")
    subir_botao.click()
    sleep(0.8)
    
    # Fechamos a janela referente ao anexo
    _ = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "svg[data-testid='CancelIcon']")))
    botao_fechar = driver.find_element(By.CSS_SELECTOR, "svg[data-testid='CancelIcon']")
    botao_fechar.click()
    sleep(1)

def identificar_documentos_para_anexar():
    BASE_DIR = os.getcwd()
    PRODUTIVIDADE_DIR = os.path.join(BASE_DIR, 'Produtividade')

    files = os.listdir(PRODUTIVIDADE_DIR)

    df = pd.read_excel('data/base.xlsx', sheet_name='main', engine="openpyxl")

    nomes = list(set(df['Correspondente']))

    arquivos_por_nome = defaultdict(list)

    for file in files:
        for nome in nomes:
            if nome in file:
                caminho_completo = os.path.join(PRODUTIVIDADE_DIR, file)
                arquivos_por_nome[nome].append(caminho_completo)

    quantidade_arquivos_por_nome = {nome: {'quantidade': len(arquivos), 'arquivos': arquivos} for nome, arquivos in arquivos_por_nome.items()}

    for arquivo in quantidade_arquivos_por_nome[str(self.nome_do_prestador)]['arquivos']:
        print(f'Inserindo o {arquivo}')
        anexar_documento(arquivo)
    
    print('Processo de anexar documentos finalizado.')