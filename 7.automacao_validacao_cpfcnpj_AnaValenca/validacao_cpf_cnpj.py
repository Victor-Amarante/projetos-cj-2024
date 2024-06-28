import os
import pandas as pd
import numpy as np
import re
import requests
import json
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from validate_docbr import CPF, CNPJ 


BASE_DIR = os.getcwd()
DATA_DIR = os.path.join( BASE_DIR, 'data' )

cpf_file_path = [ os.path.join( DATA_DIR, file ) for file in os.listdir(DATA_DIR) ][0]
cnpj_file_path = [ os.path.join( DATA_DIR, file ) for file in os.listdir(DATA_DIR) ][1]

df_cpf = pd.read_excel( cpf_file_path )
df_cnpj = pd.read_excel( cnpj_file_path )

print(f'A base referente aos CPFs possui: {df_cpf.shape[0]} linhas.\nA base referente aos CNPJs possui: {df_cnpj.shape[0]} linhas')

# --- CPF
# Identificando os valores ausentes na coluna 'CPF/CNPJ' e substituindo os seus valores por "Não Consta"
df_cpf.loc[df_cpf['CPF/CNPJ'].isna(), 'CPF/CNPJ'] = 'Não Consta'

# Alterando todos os registros da coluna 'CPF/CNPJ' para string
df_cpf['CPF/CNPJ'] = df_cpf['CPF/CNPJ'].astype('str')

cpf = CPF()

def validar_cpf(numero_cpf: str):
    if numero_cpf:
        novo_numero_cpf = re.sub(r'\D', '', numero_cpf).strip()
        
        if len(novo_numero_cpf) == 11:
            if re.match(r'^\d{3}\.\d{3}\.\d{3}-\d{2}$', numero_cpf):
                if cpf.validate(novo_numero_cpf):
                    return {
                        'CPF': numero_cpf,
                        'API': novo_numero_cpf,
                        'Condição': 'Válido',
                        'Descrição': 'Com pontuação especial'
                    }
            elif cpf.validate(novo_numero_cpf):
                cpf_formatado = cpf.mask(novo_numero_cpf)
                return {
                    'CPF': cpf_formatado,
                    'API': novo_numero_cpf,
                    'Condição': 'Válido',
                    'Descrição': 'Sem pontuação especial'
                }
    return {
        'CPF': numero_cpf,
        'API': novo_numero_cpf,
        'Condição': 'Inválido',
        'Descrição': 'CPF inválido ou formato incorreto'
    }

def validar_cpfs_dataframe(df, coluna_cpf):
    resultados = df[coluna_cpf].apply(validar_cpf)
    df_resultados = pd.DataFrame(resultados.tolist())
    return df_resultados

# Valida e formata os CPFs na coluna 'CPFs'
df_resultados = validar_cpfs_dataframe(df_cpf, 'CPF/CNPJ')

# Salvar o arquivo em formato Excel dentro da pasta de 'data'
FILE_PATH_CPF = os.path.join(DATA_DIR, 'validacao_cpf.xlsx')
df_resultados.to_excel(FILE_PATH_CPF, index=False)

# --- CNPJ
# Identificando os valores ausentes na coluna 'CPF/CNPJ' e substituindo os seus valores por "Não Consta"
df_cnpj.loc[df_cnpj['CPF/CNPJ'].isna(), 'CPF/CNPJ'] = 'Não Consta'

# Alterando todos os registros da coluna 'CPF/CNPJ' para string
df_cpf['CPF/CNPJ'] = df_cpf['CPF/CNPJ'].astype('str')

cnpj = CNPJ()

def validar_cnpj(numero_cnpj: str):
    if numero_cnpj:
        novo_numero_cnpj = re.sub(r'\D', '', numero_cnpj).strip()
        
        if len(novo_numero_cnpj) == 14:
            if re.match(r'^\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}$', numero_cnpj):
                if cnpj.validate(novo_numero_cnpj):
                    return {
                        'Documento': numero_cnpj,
                        'API': novo_numero_cnpj,
                        'Condição': 'Válido',
                        'Descrição': 'CNPJ com pontuação especial'
                    }
            elif cnpj.validate(novo_numero_cnpj):
                cnpj_formatado = cnpj.mask(novo_numero_cnpj)
                return {
                    'Documento': cnpj_formatado,
                    'API': novo_numero_cnpj,
                    'Condição': 'Válido',
                    'Descrição': 'CNPJ sem pontuação especial'
                }
    return {
        'Documento': numero_cnpj,
        'API': novo_numero_cnpj,
        'Condição': 'Inválido',
        'Descrição': 'CNPJ inválido ou formato incorreto'
    }

def validar_cnpjs_dataframe(df, coluna_cpf):
    df[coluna_cpf] = df[coluna_cpf].astype(str)
    resultados = df[coluna_cpf].apply(validar_cnpj)
    df_resultados = pd.DataFrame(resultados.tolist())
    return df_resultados

df_resultados_cnpj = validar_cnpjs_dataframe(df_cnpj, 'CPF/CNPJ')

# Salvar o arquivo em formato Excel dentro da pasta de 'data'
FILE_PATH_CPF = os.path.join(DATA_DIR, 'validacao_cnpj.xlsx')
df_resultados_cnpj.to_excel(FILE_PATH_CPF, index=False)