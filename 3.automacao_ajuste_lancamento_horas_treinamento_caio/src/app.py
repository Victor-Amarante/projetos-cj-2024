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

# organizando pastas
BASE_DIR = os.getcwd()
SRC_DIR = os.path.join(BASE_DIR, 'src')
DATA_DIR = os.path.join(BASE_DIR, 'data')

# importando base de dados
for file in os.listdir(DATA_DIR):
    if file.endswith('.xlsx'):
        df = pd.read_excel(os.path.join(DATA_DIR, file))

# inicializando chrome
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 15
wait = WebDriverWait(driver, timeout)

# credenciais
url_sharepoint = r'https://queirozcavalcanti.sharepoint.com/sites/qca360/Lists/treinamentos_qca/AllItems.aspx?xsdata=MDV8MDF8fGVhZjhlOTRjNWU4NDRmZTdmZjFiMDhkYmNlM2JhZmQ3fDk1NGQ1NjllZjk3NTRmNGFiZWQwOWViMzBkMTI3NjRifDB8MHw2MzgzMzA1Mjc5NTE2MTc4MTl8VW5rbm93bnxWR1ZoYlhOVFpXTjFjbWwwZVZObGNuWnBZMlY4ZXlKV0lqb2lNQzR3TGpBd01EQWlMQ0pRSWpvaVYybHVNeklpTENKQlRpSTZJazkwYUdWeUlpd2lWMVFpT2pFeGZRPT18MXxMMk5vWVhSekx6RTVPakUzT1dZMU1qTmpMVFl6TnpjdE5ERmxZUzFpTTJFMExXRmtNRGxtWkdSbU1HVm1NbDg1TmpZME1qZzVaaTA1WmpObUxUUm1PR1F0WW1NelpTMWhOemMxTmpBMk5UUXhaRFJBZFc1eExtZGliQzV6Y0dGalpYTXZiV1Z6YzJGblpYTXZNVFk1TnpRMU5UazVNemd6Tnc9PXwxMjNjZDQ4MTVjOWE0ODJlZmYxYjA4ZGJjZTNiYWZkN3wyNjZhY2Q1NzEwMzQ0NjJjOTFhNmVhMGI5NzliZmQyZg%3D%3D&sdata=WDFmalJYeFFEUU9oU1lIRzh2eWp5MWw3cnUrWGx0WmlFQzlZbFo0N3ZWND0%3D&ovuser=954d569e%2Df975%2D4f4a%2Dbed0%2D9eb30d12764b%2Cvictoramarante%40queirozcavalcanti%2Eadv%2Ebr&OR=Teams%2DHL&CT=1697463564021&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiIyNy8yMzA5MDExMjI3OCIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D&SafelinksUrl=https%3A%2F%2Fqueirozcavalcanti%2Esharepoint%2Ecom%2Fsites%2Fqca360%2FLists%2Ftreinamentos%5Fqca%2FAllItems%2Easpx'
email = 'victoramarante@queirozcavalcanti.adv.br'
senha = 'Qca%6925757'

# entrar no site
driver.get(url_sharepoint)
sleep(3)

# email
email_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0116']")))  
email_input.send_keys(email)
sleep(.5)
avancar_button = driver.find_element(By.XPATH, "//input[@id='idSIButton9']")   
avancar_button.click() 

# senha
senha_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='i0118']")))
senha_input.send_keys(senha)
sleep(.5)
avancar_button = driver.find_element(By.XPATH, "//input[@id='idSIButton9']")
avancar_button.click()
sleep(3)

# "Continuar Conectado?"
botao_sim = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@id='idSIButton9']")))
botao_sim.click()

print('Entramos no Sharepoint. Aguarde para iniciar o procedimento de preenchimento das informações dos treinamentos.')
sleep(10)

print('Iniciando preenchimento')

casos_sucesso = []
casos_fracasso = []
try:
    # --- INICIAR FLUXO DE ITERACAO PARA PREENCHIMENTO DE INFORMACOES SOBRE O TREINAMENTO DOS COLABORADORES ---
    for index, id in enumerate(df['ID']):
        colaborador = df.loc[index, 'Nome']
        email_colaborador = df.loc[index, 'Email']
        equipe = df.loc[index, 'EQUIPE']
        unidade = df.loc[index, 'UNIDADE']
        treinamento = df.loc[index, 'TREINAMENTO']
        tipo_de_treinamento = df.loc[index, 'TIPO DO TREINAMENTO']
        categoria = df.loc[index, 'CATEGORIA']
        instituicao_instrutor = df.loc[index, 'INSTITUIÇÃO/INSTRUTOR']
        carga_horaria = df.loc[index, 'CARGA HORÁRIA']
        inicio_do_treinamento = df.loc[index, 'INICIO DO TREINAMENTO']
        termino_do_treinamento = df.loc[index, 'TERMINO DO TREINAMENTO']
        
        print(f'Adicionando informações do colaborador: {colaborador}')

        # adicionar um novo treinamento
        botao_novo_elemento_esperado = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//*[@id='appRoot']/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button")))
        botao_novo = driver.find_element(By.XPATH, "//*[@id='appRoot']/div[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div/div[1]/div[1]/button")
        botao_novo.click()
        
        sleep(10)

        # ----- MUDAR PARA O IFRAME -----
        iframe = driver.find_elements(By.XPATH, "//iframe")
        driver.switch_to.frame(iframe[0])

        iframe2 = driver.find_element(By.XPATH, "//iframe[@class='player-app-frame']")
        driver.switch_to.frame(iframe2)
        sleep(3)

        # --- FUNCAO PARA PREENCHER INFORMACOES QUE CLICAR, ESCREVE E SELECIONA ---
        def clica_seleciona_informacao(endereco1, endereco2, valor2, endereco3):
            elemento1 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco1)))
            elemento1.click()
            sleep(3)
            elemento2 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco2)))
            elemento2.send_keys(valor2)
            sleep(3)
            elemento3 = wait.until(EC.element_to_be_clickable((By.XPATH, endereco3)))
            elemento3.click()
            
        # --- NOME DO INTEGRANTE ---
        clica_seleciona_informacao(endereco1='//div[@title="NOME DO INTEGRANTE"]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-0"]/div/div/div/div/input', valor2=str(colaborador),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-0"]/div/ul/li/div/div/span[1][text() = "{str(colaborador)}"]')
        sleep(2)
        # --- EMAIL ---
        clica_seleciona_informacao(endereco1='//div[@title="E-MAIL"]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-1"]/div/div/div/div/input', valor2={email_colaborador},
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-1"]/div/ul/li/div/span[text() = "{str(email_colaborador)}"]')
        sleep(2)
        # --- EQUIPE ---
        clica_seleciona_informacao(endereco1='//div[@title="EQUIPE."]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-2"]/div/div/div/div/input', valor2=str(equipe),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-2"]/div/ul/li/div/span[text() = "{str(equipe)}"]')
        sleep(2)
        # --- UNIDADE ---
        clica_seleciona_informacao(endereco1='//div[@title="UNIDADE"]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-3"]/div/div/div/div/input', valor2=str(unidade),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-3"]/div/ul/li/div/span[text() = "{str(unidade)}"]')
        sleep(2)
        
        # --- TREINAMENTO ---
        driver.find_element(By.XPATH, '//input[@title="TREINAMENTO"]').send_keys(str(treinamento))
        sleep(2)
        
        # --- TIPO DO TREINAMENTO ---
        clica_seleciona_informacao(endereco1='//div[@title="TIPO DO TREINAMENTO."]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-4"]/div/div/div/div/input', valor2=str(tipo_de_treinamento),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-4"]/div/ul/li/div/span[text() = "{str(tipo_de_treinamento)}"]')
        sleep(2)
        # --- CATEGORIA ---
        clica_seleciona_informacao(endereco1='//div[@title="CATEGORIA"]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-5"]/div/div/div/div/input', valor2=str(categoria),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-5"]/div/ul/li/div/span[text() = "{str(categoria)}"]')
        sleep(2)
        
        # --- SCROOL PARA O ELEMENTO ---
        elemento_scroll = driver.find_element(By.XPATH, '//input[@title="INSTITUIÇÃO/INSTRUTOR"]')
        driver.execute_script("arguments[0].scrollIntoView(true);", elemento_scroll)
        sleep(.5)
        
        # --- INSTITUICAO/INSTRUTOR ---
        driver.find_element(By.XPATH, '//input[@title="INSTITUIÇÃO/INSTRUTOR"]').send_keys(str(instituicao_instrutor))
        sleep(2)
        
        # --- CARGA HORARIA ---
        clica_seleciona_informacao(endereco1='//div[@title="CARGA HORARIA"]',
                                endereco2='//*[@id="powerapps-flyout-react-combobox-view-6"]/div/div/div/div/input', valor2=str(carga_horaria),
                                endereco3=f'//*[@id="powerapps-flyout-react-combobox-view-6"]/div/ul/li/div/span[text() = "{str(carga_horaria)}"]')
        sleep(2)
        
        # --- DATA DE INICIO ---
        data_inicio = driver.find_element(By.XPATH, '//input[@title="INICIO DO TREINAMENTO"]')
        data_inicio.send_keys(str(inicio_do_treinamento))
        sleep(2)
        
        # --- DATA FINAL ---
        data_final = driver.find_element(By.XPATH, '//input[@title="TERMINO DO TREINAMENTO"]')
        data_final.send_keys(str(termino_do_treinamento))
        sleep(2)
        
        # --- SALVAR ---
        driver.switch_to.default_content()
        salvar = driver.find_element(By.XPATH, '//span[text() = "Salvar"]')
        salvar.click()
        
        print(f'{index+1} - ID {id} - {colaborador} - {treinamento} - finalizado')
        
        casos_sucesso.append({'Caso': id, 'Status': 'Sucesso'})
        sleep(3)
        
except TimeoutException as e:
    print(f'Erro ao processar o treinamento {id}')
    casos_fracasso.append({'Treinamento': id, 'Status': f'Erro: {str(e)}'})
    
except Exception as e:
    print(f'Erro inesperado ao processar o treinamento {id}')
    casos_fracasso.append({'Treinamento': id, 'Status': f'Erro inesperado: {str(e)}'})

# exportando as bases para controle do usuario
df_sucesso = pd.DataFrame(casos_sucesso)
df_fracasso = pd.DataFrame(casos_fracasso)
df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

sleep(2)

# fechando o browser para finalizar o procedimento
driver.quit()

print('O processo finalizou! Verificar no Sharepoint as informações editadas.')