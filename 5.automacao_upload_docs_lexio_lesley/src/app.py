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

from dotenv import load_dotenv
import os


# --- ESCOLHER QUAL PASTA/EQUIPE SERÁ UTILIZADA PARA FAZER O UPLOAD ---
nome_da_equipe = 'TESTE' # ALTERAR O NOME DA EQUIPE PARA FICAR IGUAL AO NOME DA PASTA QUE ESTÁ NO ONEDRIVE.

# --- ORGANIZAÇÃO DO DIRETÓRIO --- C:\Users\victoramarante\Queiroz Cavalcanti Advocacia\Equipe Controladoria Jurídica - Núcleo de Operacões\01. Fluxos jurídicos\POP's - equipes\BANCO INTER
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
CENTRAL_DIR = os.path.dirname(os.path.dirname(BASE_DIR))

POPS_DIR = os.path.join(CENTRAL_DIR, "Queiroz Cavalcanti Advocacia", "Equipe Controladoria Jurídica - Núcleo de Operacões", "01. Fluxos jurídicos", "POP's - equipes", str(nome_da_equipe))

# --- CARREGAR E TRATAR A BASE DE DADOS ---
file_path = [os.path.join(DATA_DIR, file) for file in os.listdir(DATA_DIR)][0]
df = pd.read_excel(file_path)
df['Arquivos'] = os.listdir(POPS_DIR)
df['Data Inicio'] = df['Data Inicio'].dt.strftime('%d/%m/%Y')
df['Data Conclusao'] = df['Data Conclusao'].dt.strftime('%d/%m/%Y')

# --- EXTRAIR AS CREDENCIAIS DA LEXIO ---
load_dotenv()

email = os.getenv('USER_EMAIL')
senha = os.getenv('USER_PASSWORD')

# --- INICIALIZAR O CHROME ---
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 20
wait = WebDriverWait(driver, timeout)

# --- USAR AS CREDENCIAIS PARA ENTRAR NA PLATAFORMA ---
# Entramos na url da Lexio
driver.get('https://app.lexio.legal/login')
sleep(5)

# Inserimos o email do usuario
username_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//form//input[@id="_username"]')))
username_input.send_keys(email)
sleep(1)

# Inserimos a senha do usuario
senha_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//form//input[@id="_password"]')))
senha_input.send_keys(senha)
sleep(1)

# Clicamos no botao de confirmar o login para entrar na plataforma
login_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="login"]/div/div[2]/form/div[2]/input')))
login_confirmar.click()
sleep(10)

casos_sucesso = []
casos_fracasso = []
# --- ENTRAR NA PASTA DE TRABALHO DE PROCEDIMENTOS OPERACIONAIS PADRÕES ---
for idx, linha in df.iterrows():
    try:
        data_de_inicio = linha['Data Inicio']
        data_de_conclusao = linha['Data Conclusao']
        ano = linha['Data Conclusao'].split('/')[-1]
        mes = linha['Data Conclusao'].split('/')[-2]
        dia = linha['Data Conclusao'].split('/')[0]
        contrato = linha['Arquivos']
        caminho_arquivo = os.path.join(POPS_DIR, linha['Arquivos'])

        driver.get('https://app.lexio.legal/group/')
        sleep(2)

        # Clicar para abrir as opções das pastas disponíveis
        wait.until(EC.element_to_be_clickable((By.XPATH, "//details//summary//div")))
        sleep(.5)
        driver.find_element(By.XPATH, "//details//summary//div").click()
        sleep(.8)

        # Selecionar a pasta referente a Procedimentos Operacionais Padrões
        wait.until(EC.element_to_be_clickable((By.XPATH, "//details//div//ul[@class='kGibhRl6Doson_63wUPD__folder-list']//li//a")))
        sleep(.5)
        pastas = driver.find_elements(By.XPATH, "//details//div//ul[@class='kGibhRl6Doson_63wUPD__folder-list']//li//a")
        if pastas[0].text == 'Procedimentos Operacionais Padrões':
            pastas[0].click()
        sleep(2)
        
        # --- UPLOAD DE DOCUMENTO ---
        # Inserir o caminho do documento
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dashboard"]/div/div[2]/div/div/div[2]')))
        sleep(.5)
        driver.find_element(By.XPATH, '//*[@id="dashboard"]/div/div[2]/div/div/div[2]/input').send_keys(caminho_arquivo)
        sleep(4)

        # Clicar em "Sim"
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="document-upload"]/form/div/button[2]')))
        sleep(.5)
        driver.find_element(By.XPATH, '//*[@id="document-upload"]/form/div/button[2]').click()
        sleep(2)

        # Digitar o Nome do Contrato
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="document-name"]')))
        sleep(.5)
        nome_contrato = driver.find_element(By.XPATH, '//*[@id="document-name"]')
        nome_contrato.clear()
        nome_contrato.send_keys(f'{contrato}')
        sleep(1)

        # Digitar o tipo do documento
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="document-type-list-name"]')))
        sleep(.5)
        driver.find_element(By.XPATH, '//*[@id="document-type-list-name"]').send_keys('Queiroz Cavalcanti | Procedimento Operacional Padrão')
        sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="select-document-type-list-name-item-Queiroz Cavalcanti | Procedimento Operacional Padrão"]')))
        sleep(.5)
        driver.find_element(By.XPATH, '//*[@id="select-document-type-list-name-item-Queiroz Cavalcanti | Procedimento Operacional Padrão"]').click()
        sleep(1)

        # Clicar no botão de próximo
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="document-upload"]/form/div/button[1]')))
        sleep(.5)
        botao_proximo = driver.find_element(By.XPATH, '//*[@id="document-upload"]/form/div/button[1]')
        botao_proximo.click()
        sleep(1)

        # Clicar no botão de enviar
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="document-upload"]/form/div/button[1]')))
        sleep(.5)
        botao_enviar = driver.find_element(By.XPATH, '//*[@id="document-upload"]/form/div/button[1]')
        botao_enviar.click()
        sleep(2)
        
        # --- INSERIR AS TAGS NOS DOCUMENTOS ---
        # Clicar no botão para inserir a tag
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="etiquetasFlex"]/div[1]/a')))
        sleep(.5)
        driver.find_element(By.XPATH, '//*[@id="etiquetasFlex"]/div[1]/a').click()
        sleep(12)

        # Verificar quais etiquetas já existem
        tags = [tag.text for tag in driver.find_elements(By.XPATH, '//div[@class="containerEtiqueta"]//div')]
        sleep(.8)

        # Selecionar a etiqueta de acordo com o nome da nossa pasta
        escolher_etiqueta = driver.find_element(By.XPATH, f'//div[@class="containerEtiqueta"]//div//span//b[text() = "{nome_da_equipe}"]')
        sleep(.5)
        escolher_etiqueta.click()
        sleep(1)

        # Fechar a etiqueta
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="etqStore"]/i')))
        sleep(.5)
        fechar_etiqueta = driver.find_element(By.XPATH, '//*[@id="etqStore"]/i')
        fechar_etiqueta.click()
        sleep(4)
        
        # --- INSERIR AS DATAS DE INICIO E CONCLUSAO

        data_inicio = driver.find_element(By.XPATH, '//*[@id="infoIncDate"]')
        sleep(1)
        data_inicio.send_keys(data_de_inicio)

        data_conclusao = driver.find_element(By.XPATH, '//*[@id="infoFimDate"]')
        sleep(1)
        data_conclusao.send_keys(data_de_conclusao)
        sleep(5)

        # --- ADICIONAR EVENTO ---
        # Clicar em adicionar evento
        botao_add_evento = driver.find_element(By.XPATH, '//*[@id="loading-page"]/div[5]/div/div[2]/div[10]/div[1]/button')
        botao_add_evento.click()
        sleep(3)

        # Colocar na data do evento a data de conclusao
        meses = {
            '01': 'Jan',
            '02': 'Fev',
            '03': 'Mar',
            '04': 'Abr',
            '05': 'Mai',
            '06': 'Jun',
            '07': 'Jul',
            '08': 'Ago',
            '09': 'Set',
            '10': 'Out',
            '11': 'Nov',
            '12': 'Dez'
        }

        sleep(3)
        driver.find_element(By.XPATH, '//*[@id="modal_novo_evento"]/div[1]/div/span/button').click()
        sleep(.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[1]/table/thead/tr[2]/th[@class="datepicker-switch"]')))
        driver.find_element(By.XPATH, '/html/body/div[4]/div[1]/table/thead/tr[2]/th[@class="datepicker-switch"]').click()
        sleep(.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div[2]/table/thead/tr[2]/th[@class="datepicker-switch"]')))
        driver.find_element(By.XPATH, '/html/body/div[4]/div[2]/table/thead/tr[2]/th[@class="datepicker-switch"]').click()
        sleep(.5)
        # driver.find_element(By.XPATH, f'/html/body/div[4]/div[3]/table/tbody/tr/td/span[6][text()="{ano}"]').click()
        driver.find_element(By.XPATH, f'/html/body/div[4]/div[3]/table/tbody/tr/td/span[6][text()="2024"]').click()
        sleep(.5)
        # driver.find_element(By.XPATH, f'/html/body/div[4]/div[2]/table/tbody/tr/td/span[contains(@class, "month")][text()="{meses[mes]}"]').click()
        driver.find_element(By.XPATH, f'/html/body/div[4]/div[2]/table/tbody/tr/td/span[contains(@class, "month")][text()="Jul"]').click()
        sleep(.5)
        # driver.find_element(By.XPATH, f'//td[@class="day" and text()="{dia}"]').click()
        driver.find_element(By.XPATH, f'//td[@class="day" and text()="24"]').click()
        sleep(.8)

        # Colocar título padrão: Vigência do documento
        titulo = driver.find_element(By.XPATH, '//*[@id="modal_novo_evento"]/div[2]/input')
        titulo.send_keys('Vigência do documento')
        sleep(.8)

        # Colocar descrição padrão: Prazo de verificação para vigência do documento.
        descricao = driver.find_element(By.XPATH, '//*[@id="modal_novo_evento"]/div[3]/textarea')
        descricao.send_keys('Prazo de verificação para vigência do documento.')
        sleep(.8)

        # Clicar no botão de salvar
        botao_salvar = driver.find_element(By.XPATH, '//*[@id="modal-conteudo"]/div[3]/button[4]')
        botao_salvar.click()
        print(f'{contrato} - finalizado')
        
        casos_sucesso.append({'Caso': contrato, 'Status': 'Sucesso'})
        sleep(5)
        
    except TimeoutException as e:
        print(f'Erro ao processar o documento {contrato}')
        casos_fracasso.append({'Caso': contrato, 'Status': f'Erro: {str(e)}'})
    
    except Exception as e:
        print(f'Erro inesperado ao processar o documento {contrato}')
        casos_fracasso.append({'Caso': contrato, 'Status': f'Erro inesperado: {str(e)}'})

# exportando as bases para controle do usuario
df_sucesso = pd.DataFrame(casos_sucesso)
df_fracasso = pd.DataFrame(casos_fracasso)
df_sucesso.to_excel('casos_sucesso.xlsx', index=False)
df_fracasso.to_excel('casos_fracasso.xlsx', index=False)

sleep(2)

driver.quit()

print('Processo finalizado!')
