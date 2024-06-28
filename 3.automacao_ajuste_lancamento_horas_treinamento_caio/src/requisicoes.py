import requests

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
import os


# organizar o diretorio
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')

file_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR)][0]
df = pd.read_excel(file_path)

# inicializando chrome
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 15
wait = WebDriverWait(driver, timeout)

# credenciais
url_sharepoint = r'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx'
email = 'victoramarante@queirozcavalcanti.adv.br'
senha = '692Qca5757'
# entrar no site
driver.get(url_sharepoint)
sleep(3)
# email
email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0116']")))  
email_input.send_keys(email, Keys.ENTER)

# senha
senha_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0118']")))
senha_input.send_keys(senha, Keys.ENTER)

sleep(3)
# "Continuar Conectado?"
botao_sim = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='idSIButton9']")))
botao_sim.click()
print('Entramos no Sharepoint. Aguarde para iniciar o procedimento de preenchimento das informações dos treinamentos.')

# Obter cookies do navegador
cookies = driver.get_cookies()

# Fechar o navegador
driver.quit()

# Converter cookies para o formato utilizado pelo requests
session = requests.Session()
for cookie in cookies:
    session.cookies.set(cookie['name'], cookie['value'])


url = "https://msmanaged-na.azure-apim.net/invoke"

for index, row in df.iterrows():
    email = row['Email']
    departamento = row['Departamento']
    nome = row['Nome']
    cargo = row['Cargo']
    carga_horaria = row['CARGA HORÁRIA']
    unidade = row['UNIDADE']
    categoria = row['CATEGORIA']
    instituicao = row['INSTITUIÇÃO/INSTRUTOR']
    inicio = row['INICIO DO TREINAMENTO']
    conclusao = row['TERMINO DO TREINAMENTO']
    equipe = row['EQUIPE']
    tipo_treinamento = row['TIPO DO TREINAMENTO']
    nome_treinamento = row['TREINAMENTO']
    
    payload = {
        "NOMEDOINTEGRANTE": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser",
            "Claims": f"i:0#.f|membership|{email}",
            "Department": f"{departamento}",
            "DisplayName": f"{nome}",
            "Email": f"{email}",
            "JobTitle": f"{cargo}",
            "Picture": f"https://queirozcavalcanti.sharepoint.com/sites/qca360/_layouts/15/UserPhoto.aspx?SizeL&AccountName{email}"
        },
        "CARGAHORARIA": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
            "Id": 1,
            "Value": f"{carga_horaria}"
        },
        "UNIDADE": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
            "Id": 0,
            "Value": f"{unidade}"
        },
        "TIPO_": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
            "Id": 0,
            "Value": f"{categoria}"
        },
        "INSTITUI_x00c7__x00c3_O_x002f_IN": f"{instituicao}",
        "INICIO_x0020_DO_x0020_TREINAMENT": f"{inicio}",
        "TERMINO_x0020_DO_x0020_TREINAMEN": f"{conclusao}",
        "EQUIPE_x002e_": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
            "Id": 1,
            "Value": f"{equipe}"
        },
        "TIPO_x0020_DO_x0020_TREINAMENTO_": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
            "Id": 0,
            "Value": f"{tipo_treinamento}"
        },
        "E_x002d_MAIL": {
            "@odata.type": "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser",
            "Claims": f"i:0#.f|membership|{email}",
            "Department": f"{departamento}",
            "DisplayName": f"{nome}",
            "Email": f"{email}",
            "JobTitle": f"{cargo}",
            "Picture": f"https://queirozcavalcanti.sharepoint.com/sites/qca360/_layouts/15/UserPhoto.aspx?SizeL&AccountName{email}"
        },
        "Title": f"{nome_treinamento}"
        }

    response = session.post(url, json=payload)

    if response.status_code == 201:
        print(f'Lançamento {index + 1} criado. Envio concluído')
    else:
        print(f'Falha ao enviar o formulário para o lançamento {index + 1}')

    # Atraso para evitar sobrecarga do servidor
    sleep(1)

print("Processo concluído.")
