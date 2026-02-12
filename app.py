import streamlit as st
import pandas as pd
import io

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Bens Móveis", layout="wide")

st.title("📂 Processador de Planilha de Bens Móveis")
st.markdown("""
**Instruções:**
1. Faça o upload da planilha principal (que contém as abas a serem processadas).
2. Faça o upload da planilha MATRIZ (que contém os códigos e descrições).
3. O sistema irá gerar um novo arquivo Excel com as formatações, cores e totais.
""")

# --- BARRA LATERAL (UPLOADS) ---
st.sidebar.header("Carregar Arquivos")
uploaded_file = st.sidebar.file_uploader("1. Carregar Planilha Principal (.xlsx)", type=["xlsx"])
uploaded_matriz = st.sidebar.file_uploader("2. Carregar Planilha MATRIZ (.xlsx)", type=["xlsx"])

# --- PROCESSAMENTO ---
if st.sidebar.button("Processar Planilhas"):
    if uploaded_file is None or uploaded_matriz is None:
        st.error("⚠️ Por favor, faça o upload de AMBOS os arquivos (Principal e MATRIZ).")
    else:
        try:
            # 1. LEITURA E TRATAMENTO DA MATRIZ
            # Lê colunas A e B (A=Chave, B=Descrição)
            # header=None assume que a primeira linha já é dado. Se tiver cabeçalho, o código ajusta.
            df_matriz = pd.read_excel(uploaded_matriz, usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao']
            
            # --- CORREÇÃO DO ERRO DE REINDEXING ---
            # Remove duplicatas na coluna 'Chave', mantendo a primeira ocorrência.
            # Isso simula exatamente o comportamento do PROCV do Excel.
            df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
            
            # Cria o dicionário para substituição rápida (PROCV em memória)
            lookup_dict = dict(zip(df_matriz['Chave'], df_matriz
