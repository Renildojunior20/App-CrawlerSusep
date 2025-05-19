# Instale o Streamlit antes: pip install streamlit openpyxl requests pandas

import streamlit as st
import pandas as pd
import re
import requests
from datetime import datetime
from openpyxl import load_workbook

def limpar_cnpj(cnpj):
    return re.sub(r'\D', '', str(cnpj))

def consultar_cnpj(cnpj, url_base, cert_path):
    url = url_base.replace("{CNPJ_key}", cnpj)
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
    }
    try:
        response = requests.get(url, headers=headers, verify=cert_path)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        st.warning(f"Erro ao consultar {cnpj}: {e}")
        return None

st.title("Validador SUSEP - Upload de Excel")

uploaded_file = st.file_uploader("Carregue seu arquivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Planilha1', dtype={'CNPJ': str}, engine='openpyxl')
    df['CNPJ'] = df['CNPJ'].apply(limpar_cnpj)
    df['Validação'] = df['Validação'].astype(object)

    url_base = "https://www2.susep.gov.br/safe/corretoresapig/dadospublicos/pesquisar?tipoPessoa=PJ&cnpj={CNPJ_key}&cpfCnpj={CNPJ_key}&page=1"
    cert_path = "susep.gov.br.pem"  # Atualize para o caminho correto

    dados_externos = pd.DataFrame(columns=['CNPJ', 'Produtos', 'Situação'])

    for index, row in df.iterrows():
        cnpj = row['CNPJ']
        resultado = consultar_cnpj(cnpj, url_base, cert_path)
        if resultado:
            registros = resultado.get("retorno", {}).get("registros", [])
            if registros:
                for registro in registros:
                    produtos = registro.get('produtos', '')
                    situacao = registro.get('situacao', '')
                    dados_externos = pd.concat([dados_externos, pd.DataFrame({'CNPJ': [cnpj], 'Produtos': [produtos], 'Situação': [situacao]})], ignore_index=True)
            else:
                dados_externos = pd.concat([dados_externos, pd.DataFrame({'CNPJ': [cnpj], 'Produtos': [None], 'Situação': [None]})], ignore_index=True)
        else:
            dados_externos = pd.concat([dados_externos, pd.DataFrame({'CNPJ': [cnpj], 'Produtos': [None], 'Situação': [None]})], ignore_index=True)

    # Atualizar a coluna "Validação"
    for index, row in df.iterrows():
        cnpj = row['CNPJ']
        produtos = dados_externos.loc[dados_externos['CNPJ'] == cnpj, 'Produtos'].values[0]
        situacao = dados_externos.loc[dados_externos['CNPJ'] == cnpj, 'Situação'].values[0]
        if produtos is not None and situacao is not None:
            if 'Seguros de Danos' in produtos and situacao == 'Ativo':
                df.at[index, 'Validação'] = 'Corretor habilitado'
            else:
                df.at[index, 'Validação'] = 'Corretor inválido'
        else:
            df.at[index, 'Validação'] = 'Corretor inválido'

    st.write("Resultado da validação:")
    st.dataframe(df)

    # Download do resultado
    data_atual = datetime.now().strftime("%d%m%Y")
    nome_arquivo = f'Consulta{data_atual}.xlsx'
    df.to_excel(nome_arquivo, index=False)
    with open(nome_arquivo, "rb") as file:
        st.download_button("Baixar resultado Excel", file, file_name=nome_arquivo)