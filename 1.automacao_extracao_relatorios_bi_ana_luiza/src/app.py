# ---- CARREGAMENTO DE BIBLIOTECAS NECESSÁRIAS PARA O APP FUNCIONAR ----
from ExtrairRelatorio import extrair_relatorio
from RenomearRelatorio import renomear_relatorios
from EnviarEmail import enviar_email

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
import openpyxl
import shutil
import time
import win32com.client as win32
from datetime import datetime

print('Iniciando o fluxo automatizado')

print('Acionando o fluxo de extração do relatório')
extrair_relatorio()

print('Aguarde para entrar no próximo fluxo...')
sleep(5)

print('Acionando o fluxo de redistribuição e renomeação')
renomear_relatorios()

print('Aguarde para entrar no próximo fluxo...')
sleep(5)

print('Acionando o fluxo de envio de e-mails automáticos dos relatórios')
enviar_email()

sleep(3)

print('Processo finalizado!')
