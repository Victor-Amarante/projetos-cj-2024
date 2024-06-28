#### ---- Essa automação será referente ao processo de preenchimento do Follow Up de Acordos ----

# --- CARREGAR AS BIBLIOTECAS NECESSÁRIAS --- [ X ]
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
import numpy as np
import openpyxl
import os
import shutil

# --- ORGANIZAR OS CAMINHOS DE DIRETÓRIO --- [ X ]
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')

# --- SOLICITAR AS CREDENCIAIS PARA QUEM ESTÁ UTILIZANDO O PROGRAMA --- [ X ]
nome_usuario = 'Denise Ferreira dos Santos'
email_usuario = 'denisesantos@queirozcavalcanti.adv.br'
senha_usuario = 'Qca123456##'

# --- CARREGAR A BASE DE DADOS COM AS INFORMAÇÕES DOS FOLLOW UPs --- [ X ]
file_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR)][0]
df_initial = pd.read_csv(file_path, encoding='utf-8', sep=';')
# df_initial = pd.read_excel(file_path)
# df = df_initial.copy()
df = df_initial[['Tags', 'Descrição']]
df.loc[:, 'Numero Processo'] = df['Descrição'].apply(lambda x: str(x).split(' ')[0])
df.loc[:, 'Follow Up de Acordo'] = df['Tags'] + ': ' + df['Descrição']

# --- INICIALIZAR O BROWSER --- [ X ]
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 60
wait = WebDriverWait(driver, timeout)

# --- ENTRAR NO PERFORMA --- [ X ]
print('Acessando a plataforma do Perfoma. Inserindo as credenciais necessárias')
# Acessar o link de entrada
url_performa = 'https://performaqca.seven.adv.br/login'
driver.get(url= url_performa)
sleep(2)

# --- CREDENCIAIS RELACIONADAS A REDE ---
# Acessar a platafora com as credenciais de rede
logar_credenciais_rede = wait.until(EC.element_to_be_clickable((By.XPATH, '//a[@id="buttonVerificarUsuario"]')))
logar_credenciais_rede.click()
sleep(2)

# Inserir o email corporativo
user_email = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="i0116"]')))
user_email.send_keys(email_usuario)
user_email.send_keys(Keys.ENTER)
sleep(2)

# Inserir a senha do email corporativo
user_password = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="i0118"]')))
user_password.send_keys(senha_usuario)
user_password.send_keys(Keys.ENTER)

sleep(2)
botao_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
botao_confirmar.click()

wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[2]/div/div[1]/div[1]/ol/li/a')))
sleep(1)

print('Entramos dentro do Perfoma. Seguindo para o próximo passo...\n')

# --- PESQUISAR PELO NÚMERO DO PROCESSO --- [ X ]
print('Iniciando as iterações sobre os processos\n')

for index, processo in enumerate(df['Numero Processo']):
    follow_up = df.loc[index, 'Follow Up de Acordo']
    try:
        pesquisar_processo = driver.find_element(By.XPATH, '//*[@id="textFinder"]')
        sleep(1)
        pesquisar_processo.send_keys(processo)
        pesquisar_processo.send_keys(Keys.ENTER)
        sleep(2)

        # --- ENTRAR EM NEGOCIAÇÃO --- [ X ]
        negociacao = wait.until(EC.element_to_be_clickable((By.XPATH, '//ul//li//a[@href = "#box-negociacao"]')))
        negociacao.click()
        sleep(1)

        # --- VERIFICAR O STATUS DE ACORDO PARA PODER INSERIR AS INFORMAÇÕES --- [ X ]
        status_acordo = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="negociacaoList"]/div[1]/table/tbody/tr[1]/td[2]/span')))
        status_text = status_acordo.text
        responsavel_acordo = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="negociacaoList"]/div[1]/table/tbody/tr[8]/td[2]/span')))
        responsavel_text = responsavel_acordo.text
        if (status_text == 'Em Negociação' or status_text == 'Em Negociação Pós Sentença') and (responsavel_text == nome_usuario):
            print(f"O processo está liberado para inserir o Follow Up de Acordo.")
            sleep(1)
            
            # --- INSERIR O FOLLOW UP --- [ X ]
            # clicar no botão follow up
            botao_follow_up = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="buttonNovoFollowupAcordo"]')))
            botao_follow_up.click()
            sleep(1)
            
            # selecionar o tipo de advogado - como adv da parte
            adv_parte = driver.find_elements(By.XPATH, '//ins[@class="iCheck-helper"]')
            adv_parte[1].click()
            # sleep(0.8)
            # escrever no campo de observacao a descrição do follow up de acordo
            observacao = driver.find_element(By.XPATH, '//*[@id="Observacao"]')
            observacao.send_keys(follow_up)
            sleep(0.8)
            
            # clicar no botão salvar para salvar as informações do acordo do processo
            botao_salvar = driver.find_element(By.XPATH, '//*[@id="dialog-modal"]/div/div/div[3]/button[1]')
            botao_salvar.click()
            print(f'Processo {index+1}: {processo}\n')
            sleep(3)
        else:
            print('Bloqueado')
    except Exception as e:
        print(f'Erro ao processar o processo {processo}\n')
        continue

driver.quit()
print('O processo foi finalizado.')
