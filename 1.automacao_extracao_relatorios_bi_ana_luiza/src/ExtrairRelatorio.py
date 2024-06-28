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


def extrair_relatorio(): # QUANDO FINALIZAR O PROCEDIMENTO COLOCAR TUDO DENTRO DE UMA UNICA FUNCAO
    # organizando os diretórios
    BASE_DIR = os.getcwd()
    CORRESP_DIR = os.path.join(BASE_DIR, 'email')
    
    # pegando o caminho exato do arquivo para ler os dados 
    file_path = [os.path.join(CORRESP_DIR, file) for file in os.listdir(CORRESP_DIR)][0]

    # ---- CARREGAR A BASE DE DADOS ----
    df = pd.read_excel(file_path, sheet_name='Planilha2')
    lista_correspondentes = list(df['CORRESPONDENTE'])
    
    # inicializar o chrome
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    timeout = 60
    wait = WebDriverWait(driver, timeout)

    def filtrar_correspondente(nome_correspondente):
        # clicar em filtro
        botao_filtro = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-group/transform/div/div[2]/visual-container[1]/transform/div/div[3]/div/div/visual-modern/div/div')))
        botao_filtro.click()
        sleep(2)

        # clicar em selecao de correspondente
        botao_selecionar_corresp = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-group[1]/transform/div/div[2]/visual-container[3]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div/div')))
        botao_selecionar_corresp.click()
        sleep(2)

        # digitar o nome do correspondente
        digitar_corresp = driver.find_elements(By.XPATH, '//input[@class="searchInput"]')[-1]
        digitar_corresp.clear()
        digitar_corresp.send_keys(nome_correspondente)
        sleep(2)

        # selecionar o correspondente
        selecionar_corresp = wait.until(EC.element_to_be_clickable((By.XPATH, f'//div[@title="{nome_correspondente}"]')))
        selecionar_corresp.click()
        sleep(1)

        # apagar o correspondente que foi escrito
        digitar_corresp.clear()

        # fechar o botao de filtro
        botao_fechar_filtro = driver.find_elements(By.XPATH, '//visual-modern//div[@class="imageBackground"]')[0]
        botao_fechar_filtro.click()
        sleep(1)

    def baixar_arquivo_pdf():
        # baixar arquivo PDF
        botao_exportar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportMenuBtn"]')))
        botao_exportar.click()
        sleep(1)
        botao_pdf = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@data-testid="export-to-pdf-btn"]//span[text()="PDF"]')))
        botao_pdf.click()
        sleep(1)
        botao_exportar_pagina_atual = driver.find_elements(By.XPATH, "//div[@class='pbi-checkbox-checkbox']")[1]
        botao_exportar_pagina_atual.click()
        sleep(1)
        botao_exportar = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="okButton"]')))
        botao_exportar.click()


    # entrar no power bi
    sleep(3)
    url = "https://app.powerbi.com/"
    driver.get(url= url)

    # inserir email
    login_email = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="i0116"]')))
    sleep(.5)
    login_email.send_keys('victoramarante@queirozcavalcanti.adv.br')
    avancar_botao = driver.find_element(By.XPATH, '//*[@id="idSIButton9"]')
    avancar_botao.click()
    sleep(2)

    # inserir senha
    senha_email = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="i0118"]')))
    sleep(.5)
    senha_email.send_keys('Qca%6925757')
    avancar_botao = driver.find_element(By.XPATH, '//*[@id="idSIButton9"]')
    avancar_botao.click()
    sleep(2)

    # continuar conectado
    continuar_conectado_botao =  wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
    continuar_conectado_botao.click()
    sleep(5)

    # entrar no workspace BI Acompanhamento de Serviços de Correspondentes
    wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Novo relatório"]')))
    url_bi_acompanhamento = 'https://app.powerbi.com/groups/f006acbd-5907-4ec8-8c79-53a674a74586/reports/79893444-95f1-4c10-b767-c3272c0f10d8/ReportSection007656d32289422853d7?experience=power-bi'
    driver.get(url_bi_acompanhamento)
    sleep(5)

    # entrar na aba padrao qca correspondente
    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fileMenuBtn"]')))    # esperamos até que um elemento fixo na pagina carregue. Esse botao em específico é o que "arquivar"
    aba_padrao_qca = driver.find_elements(By.XPATH, "//visual-container[@class='visual-container-component ng-star-inserted']//transform//div//div[@class='visualContent']//div//div//visual-modern//div")[6]
    aba_padrao_qca.click()

    sleep(10)

    wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="fileMenuBtn"]'))) # esperamos um elemento ficar clicável para iniciar o procedimento de extração

    # ---- SELECIONAR CORRESPONDENTE ----
    for idx, corresp in enumerate(lista_correspondentes):
        filtrar_correspondente(corresp)
        sleep(2)
        baixar_arquivo_pdf()
        sleep(35)
        print(f'Correspondente {idx+1}: {corresp}')

    sleep(3)

    driver.quit()

    print('O processo de extração dos relatórios foi finalizado.')
