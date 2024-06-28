import os
import win32com.client as win32
import pandas as pd
from datetime import datetime


def enviar_email():
    '''A função tem como objetivo enviar emails para os correspondentes de forma automatizada contendo os arquivos PDFs renomeados em anexo.'''
    
    # definir os diretorios
    BASE_DIR = os.getcwd()
    CORRESP_DIR = os.path.join(BASE_DIR, 'email')
    RELATORIO_DIR = os.path.join(BASE_DIR, 'relatorios')
    ASSINATURA_DIR = os.path.join(BASE_DIR, 'assinatura')

    # criar lista de PDFs baixados para envio
    lista_diretorio_pdfs = [os.path.join(RELATORIO_DIR, pdf) for pdf in os.listdir(RELATORIO_DIR)]
    nome_corresp = [diretorio.split('\\')[-1].split('.')[0] for diretorio in lista_diretorio_pdfs]
    df_diretorio = pd.DataFrame({'CORRESPONDENTE': nome_corresp, 'DIRETORIO': lista_diretorio_pdfs}, columns=['CORRESPONDENTE', 'DIRETORIO'])

    # abrir a base de dados com informacoes de nome do correspondente e seus emails
    email_correspondente = pd.read_excel(os.path.join(CORRESP_DIR, os.listdir(CORRESP_DIR)[0]), sheet_name='Planilha2')

    # criar uma estrutura para ter o nome do correspondente e acessar o seu email e anexo
    df = pd.merge(email_correspondente, df_diretorio, how='left', on='CORRESPONDENTE')

    # filtrar a base de dados somente para aqueles correspondentes que tem relatorios baixados (REGRA DE NEGOCIO APLICADA)
    df = df[~df['DIRETORIO'].isna()]

    # acessar cada correspondente e enviar para o email correspondente junto com o anexo referente ao mesmo
    cont = 1
    for index, destino in enumerate(df['E-MAILS PARA RECEBIMENTO DAS SOLICITAÇÕES']):
        correspondente = df.loc[index, 'CORRESPONDENTE']
        diretorio = df.loc[index, 'DIRETORIO']
        
        # inicializar a conexao com o email outlook e definir o titulo do email
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        email.To = destino
        year = datetime.now().year
        titulo = f"Padrão QCA Correspondentes {year} - Ranking - {correspondente}"
        email.Subject = titulo

        # definir a hora atual para saudacao no email
        hora_atual = datetime.now().hour
        if 6 <= hora_atual < 12:
            horario = 'Bom dia'
        else:
            horario = 'Boa tarde'

        # pegar o caminho da imagem de assinatura do email
        diretorio_assinatura = os.path.join(ASSINATURA_DIR, os.listdir(ASSINATURA_DIR)[0])
        
        # construir o corpo do texto utilizado
        email.HTMLBody = f"""
            <p>Prezados,</p>
            <p>{horario}.</p>
            <p>Estou enviando para vocês o resultado do Padrão QCA - Correspondentes referente ao ano de {year}.</p>
            <p>Gostaria de lembrar que o Padrão QCA é uma ferramenta importante para tomada de decisão em relação à qualidade da prestação de serviços, seu desempenho nesse ranking pode impactar positivamente, como possibilitar a ampliação das comarcas de atuação, ou negativamente, com a diminuição do envio de solicitações, por exemplo.</p>
            <p>É crucial que vocês acompanhem essa ferramenta, inclusive pelo desempenho mensal que também estamos enviando, com foco em busca contínua pela melhoria.</p>
            <p>Estou à disposição para qualquer dúvida ou esclarecimento que precisarem.</p>
            <p>Atenciosamente,</p>
            <img src={diretorio_assinatura} alt="Imagem da Assinatura" width="600">
            """
        
        # anexar arquivo de acordo com o diretorio do documento
        email.Attachments.Add(diretorio)

        email.Send()
        print(f"Email {cont} enviado para {destino} - {titulo}")
        cont += 1
    print('O processo de envio automatizado de email foi finalizado.')


if __name__ == '__main__':
    enviar_email()