import os
import pandas as pd
from datetime import date
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager


# Organizando o diretorio e pegando o file path que sera utilizado para o carregamento da base da dados
BASE_DIR = os.getcwd()
file_path = [os.path.join(BASE_DIR, file) for file in os.listdir(BASE_DIR) if file.endswith('.xlsx')][0]

df = pd.read_excel(file_path, engine='openpyxl')

# config de login
link = 'https://performaqca.seven.adv.br/main'

# config e inicializacao de browser
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
timeout = 30
wait = WebDriverWait(driver, timeout)

# entrar na plataforma performa
driver.get(link)
print('Esperando o tempo de login e carregamento da pagina')
sleep(40)

print('Iniciando o preenchimento')
for i, email_adv in enumerate(df['Email Advogado']):
    email_adv = df.loc[i, 'Email Advogado']
    telefone_adv = df.loc[i, 'Telefones Advogado']
    wpp_adv = df.loc[i, 'Whatsapp Advogado']
    processo = df.loc[i, 'Pasta Cliente'] # numero do processo
    nome_adv = df.loc[i, 'Nome Advogado']
    proposta_ofertada_valor = df.loc[i, 'Ultima proposta']  # foi adicionado essa variavel para inserir no campo de Proposta Ofertada na atualizacao do Follow Up de Acordo
    
    link = f"https://performaqca.seven.adv.br/processo/details/{int(processo)}"

    driver.get(link)
    sleep(5)
    try:
        # clicar em negociacao
        negociacao_button = driver.find_element(By.XPATH, "//a[@data-toggle='tab']//i[@class='fa fa-money']")
        negociacao_button.click()
        sleep(3)
    
        # verifica se eh duda sena e se o status está como negociacao ou negociacao pos sentenca
        status_acordo = driver.find_element(By.XPATH, '//*[@id="negociacaoList"]/div[1]/table/tbody/tr[1]/td[2]/span')
        status_text = status_acordo.text
        responsavel_acordo = driver.find_elements(By.XPATH, "//table[@class='table  table-striped']//tbody//tr//td//span")[3]
        responsavel_acordo = responsavel_acordo.text
        if (status_text == 'Em Negociação' or status_text == 'Em Negociação Pós Sentença') and (responsavel_acordo == "Maria Eduarda de Sena (Alterar Responsável)"):
            # clicar em "NOVO FOLLOW UP ACORDO"
            followup_button = driver.find_element(By.XPATH, "//button[@class='btn btn-export']")
            followup_button.click()
            sleep(3)
    
            # clicar em adv da parte
            advogado_parte = driver.find_elements(By.XPATH, "//div[@class='iradio_minimal-red']//ins[@class='iCheck-helper']")[1]
            advogado_parte.click()
            sleep(1)
            
            # inserir valor de proposta ofertada
            proposta_ofertada = driver.find_element(By.XPATH, "//div[@class='row']//div[@class='col-md-6']//div[@class='form-group']//input[@id='PropostaOfertada']")
            proposta_ofertada.send_keys(proposta_ofertada_valor)
            sleep(2)
    
            # digitar a observacao
            campo_observacao = driver.find_element(By.XPATH, "//textarea[@class='form-control']")
            campo_observacao.send_keys(f"AF em negociação\nNome do Advogado: {nome_adv}\nEmail: {email_adv}\nTelefones: {telefone_adv}\nWhatsApp: {wpp_adv}") # ISSO AQUI QUE VAI MUDAR
            sleep(1)
    
            # clicar em salvar
            botao_salvar = driver.find_element(By.XPATH, "//button[@class='btn btn-primary salvar']")
            botao_salvar.click()
    
            print(f"Foi realizado o preenchimento das informações do processo: {int(processo)}")
            sleep(3)
    except TimeoutException:
        print(f'Erro no processo {int(processo)}')
    except Exception:
        print(f'Erro no processo {int(processo)}')

sleep(2)

driver.quit()
print("\nO processo foi concluído!")