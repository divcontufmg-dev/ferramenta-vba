import streamlit as st
import pandas as pd
import io
import xlsxwriter
import zipfile
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Bens Móveis", layout="wide")

# --- INICIALIZAÇÃO DA MEMÓRIA INTERNA ---
if 'arquivos_memoria' not in st.session_state:
    st.session_state['arquivos_memoria'] = {}

st.title("📂 Processador de Planilha de Bens Móveis")
st.markdown("""
**Funcionalidades:**
1. **Processar:** Aplica PROCV (usando MATRIZ.xlsx local), filtros e cores.
2. **Armazenamento Interno:** Salva cada aba como um arquivo Excel individual na memória do sistema (em vez de baixar o ZIP).
""")

# --- BARRA LATERAL (UPLOADS) ---
st.sidebar.header("Carregar Arquivos")
uploaded_file = st.sidebar.file_uploader("Carregar Planilha Principal (.xlsx)", type=["xlsx"])

# --- FUNÇÃO AUXILIAR DE FORMATAÇÃO ---
def formatar_aba(writer, sheet_name, data_rows, header_rows):
    # Escreve Cabeçalho (Deslocado 1 coluna para direita)
    header_rows.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False, header=False)
    
    # Escreve Dados (Começando na coluna A, linha 8)
    data_rows.to_excel(writer, sheet_name=sheet_name, startrow=7, startcol=0, index=False, header=False)

    worksheet = writer.sheets[sheet_name]
    workbook = writer.book

    # --- DEFINIÇÃO DE FORMATOS (Recriados para cada workbook) ---
    fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
    fmt_total_label = workbook.add_format({'bold': True, 'align': 'right'})
    fmt_total_value = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
    fmt_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}) 
    fmt_blue = workbook.add_format({'bg_color': '#0000FF', 'font_color': '#FFFFFF'})

    # Largura das colunas
    worksheet.set_column('A:A', 40) # Nova Descrição
    worksheet.set_column('B:C', 15) 
    worksheet.set_column('D:D', 18, fmt_currency)

    num_rows = len(data_rows)
    start_row_excel = 7 # Linha 8
    
    # Loop para pintar as linhas
    for i in range(num_rows):
        val_conta = data_rows.iloc[i, 1] # Coluna B
        val_valor = data_rows.iloc[i, 3] # Coluna D
        
        try: val_valor = float(val_valor)
        except: val_valor = 0

        row_idx = start_row_excel + i
        
        if val_conta == 123110801 and val_valor != 0:
            worksheet.write(row_idx, 1, val_conta, fmt_red)
            worksheet.write(row_idx, 2, data_rows.iloc[i, 2], fmt_red)
            worksheet.write(row_idx, 3, val_valor, fmt_red)
        
        elif val_conta == 123119905 and val_valor != 0:
            worksheet.write(row_idx, 1, val_conta, fmt_blue)
            worksheet.write(row_idx, 2, data_rows.iloc[i, 2], fmt_blue)
            worksheet.write(row_idx, 3, val_valor, fmt_blue)

    # Total
    total_row = start_row_excel + num_rows
    soma_total = pd.to_numeric(data_rows.iloc[:, 3], errors='coerce').sum()
    worksheet.write(total_row, 2, "TOTAL", fmt_total_label)
    worksheet.write(total_row, 3, soma_total, fmt_total_value)


# --- PROCESSAMENTO PRINCIPAL ---
if st.sidebar.button("Processar Planilhas"):
    # Verifica MATRIZ local
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo 'MATRIZ.xlsx' não foi encontrado no sistema.")
    elif uploaded_file is None:
        st.error("⚠️ Por favor, faça o upload da Planilha Principal.")
    else:
        try:
            # 1. PREPARAÇÃO DOS DADOS (Lê direto do arquivo local)
            df_matriz = pd.read_excel("MATRIZ.xlsx", usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao']
            df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
            lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))

            xls_file = pd.ExcelFile(uploaded_file)
            
            processed_sheets = []

            # Loop de Processamento
            for sheet_name in xls_file.sheet_names:
                if sheet_name == "MATRIZ": continue

                df_raw = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
                if len(df_raw) < 8: continue 

                header_rows = df_raw.iloc[:7]
                data_rows = df_raw.iloc[7:].copy()

                data_rows[0] = pd.to_numeric(data_rows[0], errors='coerce')
                
                # Filtro
                exclusion_list = [123110703, 123110402, 123119910]
                data_rows = data_rows[~data_rows[0].isin(exclusion_list)]

                # PROCV
                data_rows['Nova_Descricao'] = data_rows[0].map(lookup_dict)

                # Reordenar colunas
                cols = list(data_rows.columns)
                if 'Nova_Descricao' in cols:
                    cols.insert(0, cols.pop(cols.index('Nova_Descricao')))
                data_rows = data_rows[cols]

                # Ordenar linhas
                data_rows = data_rows.sort_values(by='Nova_Descricao', ascending=True)

                processed_sheets.append({
                    'name': sheet_name,
                    'header': header_rows,
                    'data': data_rows
                })

            st.success(f"✅ Processamento concluído! {len(processed_sheets)} abas foram tratadas.")
            st.markdown("---")

            # --- SALVANDO ARQUIVOS SEPARADOS NA MEMÓRIA INTERNA ---
            # Limpa o "cofre" de memória para evitar sobreposição de processamentos anteriores
            st.session_state['arquivos_memoria'] = {}
            
            for item in processed_sheets:
                single_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(single_excel_buffer, engine='xlsxwriter') as single_writer:
                    formatar_aba(single_writer, item['name'], item['data'], item['header'])
                
                single_excel_buffer.seek(0)
                
                # Guarda o buffer na memória associado ao nome do arquivo
                nome_arquivo = f"{item['name']}.xlsx"
                st.session_state['arquivos_memoria'][nome_arquivo] = single_excel_buffer

            # Aviso de confirmação para o usuário
            st.success(f"💾 Tudo certo! {len(st.session_state['arquivos_memoria'])} arquivos foram gerados e estão salvos na memória interna.")
            
            with st.expander("Ver lista de arquivos gerados na memória"):
                for nome in st.session_state['arquivos_memoria'].keys():
                    st.write(f"📄 {nome}")

        except Exception as e:
            st.error(f"❌ Ocorreu um erro: {e}")
