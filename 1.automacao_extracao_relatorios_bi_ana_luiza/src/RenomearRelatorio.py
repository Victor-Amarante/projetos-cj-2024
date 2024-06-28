import os
import pandas as pd
import openpyxl
import shutil
import time


def renomear_relatorios():
    '''A função tem como objetivo reorganizar os relatorios e renomea-los corretamente'''
    
    # extrair os diretorios dentro da pasta de downloads
    print('Encontrando os relatórios baixados')
    BASE_DIR = os.getcwd()
    CORRESP_DIR = os.path.join(BASE_DIR, 'email')
    RELATORIO_DIR = os.path.join(BASE_DIR, 'relatorios')
    DOWNLOADS_DIR = os.path.join(os.path.dirname(os.path.dirname(BASE_DIR)), 'Downloads')
    lista_relatorios = [arquivo for arquivo in os.listdir(DOWNLOADS_DIR) if 'BI Acompanhamento de Serviços de Correspondentes' in arquivo]
    time.sleep(2)

    # mover para a pasta de relatorios
    print('Movendo os arquivos para a pasta Relatórios')
    for relatorio in lista_relatorios:
        origem = os.path.join(DOWNLOADS_DIR, relatorio)
        destino = os.path.join(RELATORIO_DIR, relatorio)
        shutil.move(origem, destino)
    time.sleep(2)

    # renomear os relatorios de acordo com o nome do Correspondente
    print('Modificando os nomes dos arquivos de acordo com o nome do Correspondente')
    for arquivo in os.listdir(CORRESP_DIR):
        if arquivo.endswith('.xlsx'):
            caminho_arquivo = os.path.join(CORRESP_DIR, arquivo)
            df_corresp = pd.read_excel(caminho_arquivo, sheet_name='Planilha2')

    df_corresp = df_corresp.rename(columns={'CORRESPONDENTE': 'correspondente', 'E-MAILS PARA RECEBIMENTO DAS SOLICITAÇÕES': 'emails'})

    dicionario = {}
    for indice, correspondente in enumerate(df_corresp['correspondente']):
        email = df_corresp.loc[indice, 'emails']
        dicionario[indice] = correspondente

    for relatorio in os.listdir(RELATORIO_DIR):
        if relatorio.endswith('.pdf'):
            numero = None
            nome_relatorio, extensao = os.path.splitext(relatorio)
            if '(' in nome_relatorio:
                numero = int(nome_relatorio.split('(')[-1].split(')')[0])

            # Verifica se o número existe no dicionário e renomeia o relatorio
            if numero is not None and numero in dicionario:
                novo_nome = f"{dicionario[numero]}{extensao}"
                os.rename(os.path.join(RELATORIO_DIR, relatorio), os.path.join(RELATORIO_DIR, novo_nome))
            elif numero is None:
                os.rename(os.path.join(RELATORIO_DIR, relatorio), os.path.join(RELATORIO_DIR, f"{dicionario[0]}{extensao}"))
    time.sleep(2)

    print('O processo de redistribuição e renomeação foi finalizado.')
    

if __name__ == '__main__':
    renomear_relatorios()
