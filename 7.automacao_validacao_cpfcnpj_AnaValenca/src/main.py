import os
import pandas as pd
import requests
import json
from validate_docbr import CPF, CNPJ 
from time import sleep
from tqdm import tqdm


validacao_cnpj = pd.read_excel(r'C:\Users\victoramarante\Documents\7.automacao_validacao_cpfcnpj_AnaValenca\data\validacao_cnpj.xlsx')
validacao_cnpj['API'] = validacao_cnpj['API'].apply(lambda x: f"{int(x):014d}" if pd.notna(x) else pd.NA)
apis = validacao_cnpj[validacao_cnpj['Condição'] == 'Válido']['API'].tolist()

def extrair_infos_cnpj(numero_cnpj):
    url = f"https://brasilapi.com.br/api/cnpj/v1/{numero_cnpj}"
    try:
        response = requests.request("GET", url,)
        response.raise_for_status()  # Raise an exception for 4xx or 5xx errors
        json_res = response.json()
        if 'message' in json_res:
            return None
        elif 'razao_social' in json_res:
            return json_res['razao_social']
        else:
            return None
    except (requests.exceptions.HTTPError, requests.exceptions.RequestException) as e:
        print(f"Error fetching data: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return None

# Inicialização das variáveis
razao_social = []

# Iteração para extração das informações com barra de progresso
print(f'Temos uma base de {len(apis)}')
for index, api in enumerate(tqdm(apis, desc="Processando CNPJs")):
    resp = extrair_infos_cnpj(api)
    razao_social.append(resp)
    tqdm.write(f'{index+1} - CNPJ {api} extraído')
    
    # Salvando a cada 500 iterações
    if (index + 1) % 100 == 0 or (index + 1) == len(apis):
        # Filtrando dados válidos e adicionando razão social
        validos = validacao_cnpj[validacao_cnpj['Condição'] == 'Válido'].copy()
        validos['Razão Social'] = pd.Series(razao_social).reindex(validos.index)
        
        # Salvando em um arquivo Excel
        filename = f'teste_validos_parcial_{index + 1}.xlsx'
        validos.to_excel(filename, index=False)
        tqdm.write(f'Arquivo salvo: {filename}')

print('\nProcesso finalizado')

# validos = validacao_cnpj[validacao_cnpj['Condição'] == 'Válido']
# validos['Razão Social'] = pd.Series(razao_social).reindex(validos.index)
# validos.to_excel('teste_validos.xlsx', index=False)

# rodar por chunks e exportar os dataframes
