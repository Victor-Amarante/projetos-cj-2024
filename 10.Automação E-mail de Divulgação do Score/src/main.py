# ----- Automação E-mail de Divulgação do Score -----

import os
import win32com.client as win32
from time import sleep
import pandas as pd
from datetime import datetime, timedelta
import locale
# --- ORGANIZAR O DIRETÓRIO da base de dados--- 
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'data')
Arquivos = os.listdir(DATA_DIR)

# caminho da pasta de Rhuan
BASE_DIR2 = os.path.dirname(os.path.dirname(os.path.dirname(os.getcwd())))

# percorrendo os diretorios da base
caminho = []
for item in Arquivos:
    caminho1 = os.path.join(DATA_DIR, item)
    caminho.append(caminho1)

df1 = pd.read_excel(caminho[0], sheet_name=0)
df1['N° SUB-INDICADOR'].unique()
df2 = pd.read_excel(caminho[0], sheet_name=1)
df2 = df2.fillna(value='N/A')

# --- ORGANIZAR O DIRETÓRIO da base de email--- 
BASE_email = os.getcwd()
DATA_email = os.path.join(BASE_DIR, 'data_email')
Arquivos_email = os.listdir(DATA_email)

caminho_planinha_email = []
for item2 in Arquivos_email:
    caminho2 = os.path.join(DATA_email, item2)
    caminho_planinha_email.append(caminho2)

df_emails = pd.read_excel(caminho_planinha_email[0])

# --- ORGANIZAR O DIRETÓRIO da base de historico--- 
BASE_historico = os.getcwd()
DATA_historico = os.path.join(BASE_DIR, 'data_grafico')
Arquivos_historico = os.listdir(DATA_historico)

caminho_planinha_historico = []
for item2 in Arquivos_historico:
    caminho3 = os.path.join(DATA_historico, item2)
    caminho_planinha_historico.append(caminho3)

df_historico = pd.read_excel(caminho_planinha_historico[0])
# --- Tratamento da base de dados df1--- 
df1 = df1.dropna()
colunas100 = ['PESO','VALOR REAL', 'DESEMPENHO']
df1[colunas100] = df1[colunas100] * 100


#----------------------------------------------------------------------------
definir_mes = '2024-06-01'       # data no formato yyyy-mm-dd 
MES = 'Junho'                    # Formato da sua pasta primeira letra maiuscula 
Numero_mes = '6'                 # numero da pasta eferente ao mes 
ano = 2024                       # ano atual
Numero_ano = 2                   # nummero da pasta referente ao ano
#----------------------------------------------------------------------------

df1_filtro = df1[df1['DATA'] == definir_mes]

# --- Tratamento da base de dados df2 --- 
def multiplica_se_numero(value):

    if isinstance(value, (int, float)):
        return round(value * 100, 2)
    else:
        return value

colunas_para_multiplicar = ['PUBLICAÇÃO', 'PA', 'PRAZO', 'OBRIGAÇÕES', 'GARANTIA', 'ENCERRAMENTO', 'ACORDO', 'SUBSÍDIO', 'CADASTRO', 'ESTRATÉGIA', 'SERVIÇOS', 'VALOR A PROVISIONAR','MÉDIA GERAL','Variação']
df2[colunas_para_multiplicar] = df2[colunas_para_multiplicar].applymap(multiplica_se_numero)


df2_group = df2.groupby('CENTRO DE CUSTO')[['PUBLICAÇÃO', 'PA', 'PRAZO', 'OBRIGAÇÕES', 'GARANTIA', 'ENCERRAMENTO', 'ACORDO', 'SUBSÍDIO', 'CADASTRO', 'ESTRATÉGIA', 'SERVIÇOS', 'VALOR A PROVISIONAR','MÉDIA GERAL','Variação']].sum().reset_index()
df_teste = pd.DataFrame(df2_group)

# --- Tratamento da base de dados df3 --- 
df_selecionado = df_historico[['MÊS', 'CENTRO DE CUSTO', 'MÉDIA GERAL']]
df_selecionado['MÉDIA GERAL'] = df_selecionado['MÉDIA GERAL'] * 100
df_selecionado['MÉDIA GERAL'] = df_selecionado['MÉDIA GERAL'].apply(lambda x: f'{x:.2f}%')

df1_filtro['N° SUB-INDICADOR'].unique() #--------------------------------------------------------------
# ------------------------------------------------------
lista_equipe = df1_filtro['CENTRO DE CUSTO'].unique()

tabelas_historico = {}
# --- Criar a integração com o Outlook ---
outlook = win32.Dispatch('outlook.application')

# Percorrer as equipes e buscar os emails
for equipe in lista_equipe:
    print(f'Equipe: {equipe} \n')
    
    df_selecionado_filtrado = df_selecionado[df_selecionado['CENTRO DE CUSTO'] == equipe]

    # Filtra os dados da equipe específica
    df_equipe = df_teste[df_teste['CENTRO DE CUSTO'] == equipe]

    # Buscar os emails para a equipe
    emails = df_emails[df_emails['CENTRO DE CUSTO'] == equipe]['E-MAILS'].tolist()
    copias = df_emails[df_emails['CENTRO DE CUSTO'] == equipe]['COPIA'].tolist()
    if not emails:
        print(f'Nenhum email encontrado para a equipe: {equipe}')
        continue

    print('email dos integrantes da equipe')
    print(';'.join(emails))
    print('\n')
    print('email da copia')
    print(';'.join(copias))
    print('\n')

    # Personalizar a tabela HTML
    # Criar cabeçalho da tabela HTML
    headers = ''.join(['<th style="padding: 8px; text-align: center; background-color: #9e1c1c; color: white; border: 1px solid black;">' + col + '</th>' for col in df_equipe.columns])

    # Criar linhas da tabela HTML
    rows = ''
    for row in df_equipe.values:
        row_html = ''.join(['<td style="padding: 8px; text-align: center; border: 1px solid black;">' + (str(cell) + '%') if isinstance(cell, (int, float)) else '<td style="padding: 8px; text-align: center; border: 1px solid black;">' + str(cell) + '</td>' for cell in row])
        rows += '<tr>' + row_html + '</tr>'


    # Montar a tabela HTML completa
    tabela_html_personalizada = '''
    <table style="border-collapse: collapse; width: 100%; border: 1px solid black;">
        <thead>
            <tr>{}</tr>
        </thead>
        <tbody>{}</tbody>
    </table>
    '''.format(headers, rows)
    
    # Calcular o score
    valores_numericos = pd.to_numeric(df_equipe['MÉDIA GERAL'], errors='coerce')
    
    score = valores_numericos.mean()
    
    # Definir o locale para português do Brasil ('pt_BR')
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

    # Data limite daqui a 2 dias
    data_limite = datetime.now() + timedelta(days=2)

    # Formatar a data de término para o formato desejado
    data_termino = data_limite.strftime('%d/%m/%Y')

    # Dia da semana da data limite (em português)
    dia_semana_termino = data_limite.strftime('%A')
    
    # Definir a mensagem com base nas condições
    if (score >= 90) and (df_equipe['Variação'] > 2.5).any():
        mensagem = f'<p><b>Parabenizamos a equipe por atingir um SCORE de {score:.2f}%, aumentando em +{df_equipe['Variação'].iloc[0]}%. 🎉👏</b></p>'
    else:
        mensagem = ''
        
    # Criar o email
    email = outlook.CreateItem(0)
    email.To = ';'.join(emails)
    email.CC = ';'.join(copias)
    email.Subject = f'Score Operacional CJ - {MES} {ano} - {equipe}'
    email.HTMLBody = f'''
    <html>
    <head> </head>
    <body>
    <p>Prezados, boa tarde!</p>
    <p>Segue o resultado da equipe no Score Operacional CJ durante o último mês.</p>
    {mensagem}
    <br>
    {tabela_html_personalizada}
    <br>
    <p>Seguem também as planilhas utilizadas na mensuração para que vocês possam observar 
    os casos pontuados e atuar nas falhas dentro da equipe.</p>
    
    <br></br>
    <b>OBS:</b> impugnações só poderão ser feitas até sexta-feira {dia_semana_termino}. ({data_termino})
    </body>
    </html>
    '''

    # Lista para armazenar os caminhos dos arquivos
    arquivos_para_anexar = []
    # Percorrer as linhas do DataFrame filtrado
    for index, linha in df1_filtro[df1_filtro['CENTRO DE CUSTO'] == equipe].iterrows():
        centro_de_custo = linha['CENTRO DE CUSTO']
        indicador = linha['INDICADOR']
        N_sub_indicador = linha['N° SUB-INDICADOR']
        N_sub_indicador_inteiro = str(N_sub_indicador).split('.')[0]
        sub_indicador = linha['SUB-INDICADOR']
 
        # Construir o caminho do arquivo
        print(f'Anexando os arquivos da equipe {centro_de_custo}')
        path = os.path.join(BASE_DIR2, 'Queiroz Cavalcanti Advocacia', 'Equipe Controladoria Jurídica - Núcleo Gestão de Dados', '09 Score das Equipes', 'Bases', f'{Numero_ano} - {ano}', f'{Numero_mes}. {MES}', f'{N_sub_indicador_inteiro}. {indicador}',f'{N_sub_indicador} {sub_indicador.strip()}',f'{N_sub_indicador} {centro_de_custo}.xlsx')
        print(f'{path}')
       
        # Verificar se o arquivo existe antes de adicioná-lo à lista
        if os.path.exists(path):
            print('')
            arquivos_para_anexar.append(path)
        else:
            print('')
            print(f'Arquivo não encontrado: {path}')
 
    # Anexar todos os arquivos da lista ao email
    for arquivo in arquivos_para_anexar:
        email.Attachments.Add(arquivo)
        
    # Salvar o email como rascunho
    email.Save()
    
    print('')
    print(f'Email enviado para a equipe: {equipe} \n')
    sleep(1)
