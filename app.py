import streamlit as st
import pandas as pd
import io
import xlsxwriter

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
            df_matriz = pd.read_excel(uploaded_matriz, usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao']
            
            # Remove duplicatas na coluna 'Chave', mantendo a primeira ocorrência.
            df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
            
            # --- CORREÇÃO DA LINHA DO ERRO ---
            # Cria o dicionário: Chave -> Descrição
            lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))

            # 2. PREPARAÇÃO DO ARQUIVO DE SAÍDA
            output = io.BytesIO()
            xls_file = pd.ExcelFile(uploaded_file)

            # Engine 'xlsxwriter' permite formatar cores e estilos
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- ESTILOS ---
                fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
                fmt_total_label = workbook.add_format({'bold': True, 'align': 'right'})
                fmt_total_value = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
                
                # Cores condicionais
                fmt_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}) # Fundo Vermelho
                fmt_blue = workbook.add_format({'bg_color': '#0000FF', 'font_color': '#FFFFFF'}) # Fundo Azul

                # Salva a aba MATRIZ no arquivo final
                df_matriz.to_excel(writer, sheet_name='MATRIZ', index=False, header=False)

                # 3. LOOP PELAS ABAS
                for sheet_name in xls_file.sheet_names:
                    if sheet_name == "MATRIZ":
                        continue

                    # Lê a aba inteira sem cabeçalho
                    df_raw = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)

                    # Verifica se tem dados suficientes (header + pelo menos 1 linha de dados)
                    if len(df_raw) < 8:
                        df_raw.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        continue

                    # Separa Cabeçalho (Linhas 0 a 6) e Dados (Linhas 7 em diante)
                    header_rows = df_raw.iloc[:7]
                    data_rows = df_raw.iloc[7:].copy()

                    # Garante que a coluna de Código (índice 0) seja numérica
                    data_rows[0] = pd.to_numeric(data_rows[0], errors='coerce')

                    # 4. FILTRO DE EXCLUSÃO
                    exclusion_list = [123110703, 123110402, 44905287]
                    data_rows = data_rows[~data_rows[0].isin(exclusion_list)]

                    # 5. APLICAÇÃO DO "PROCV"
                    data_rows['Nova_Descricao'] = data_rows[0].map(lookup_dict)

                    # Reorganiza: Coloca a 'Nova_Descricao' como primeira coluna
                    cols = list(data_rows.columns)
                    if 'Nova_Descricao' in cols:
                        cols.insert(0, cols.pop(cols.index('Nova_Descricao')))
                    data_rows = data_rows[cols]

                    # 6. ORDENAÇÃO
                    data_rows = data_rows.sort_values(by='Nova_Descricao', ascending=True)

                    # 7. GRAVAÇÃO NA PLANILHA
                    # Escreve Cabeçalho (Deslocado 1 coluna para direita)
                    header_rows.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False, header=False)
                    
                    # Escreve Dados (Começando na coluna A, linha 8)
                    data_rows.to_excel(writer, sheet_name=sheet_name, startrow=7, startcol=0, index=False, header=False)

                    # 8. FORMATAÇÃO CONDICIONAL E TOTAIS
                    worksheet = writer.sheets[sheet_name]
                    
                    # Define largura das colunas
                    worksheet.set_column('A:A', 40) # Nova Descrição
                    worksheet.set_column('B:C', 15) 
                    worksheet.set_column('D:D', 18, fmt_currency)

                    num_rows = len(data_rows)
                    start_row_excel = 7 # Linha 8
                    
                    for i in range(num_rows):
                        # Índices: 0=Desc, 1=Conta, 2=Antiga B, 3=Valor
                        val_conta = data_rows.iloc[i, 1]
                        val_valor = data_rows.iloc[i, 3]
                        
                        try: val_valor = float(val_valor)
                        except: val_valor = 0

                        row_idx = start_row_excel + i
                        
                        # Vermelho
                        if val_conta == 123110801 and val_valor != 0:
                            worksheet.write(row_idx, 1, val_conta, fmt_red)
                            worksheet.write(row_idx, 2, data_rows.iloc[i, 2], fmt_red)
                            worksheet.write(row_idx, 3, val_valor, fmt_red)
                        
                        # Azul
                        elif val_conta == 123119905 and val_valor != 0:
                            worksheet.write(row_idx, 1, val_conta, fmt_blue)
                            worksheet.write(row_idx, 2, data_rows.iloc[i, 2], fmt_blue)
                            worksheet.write(row_idx, 3, val_valor, fmt_blue)

                    # 9. TOTAL
                    total_row = start_row_excel + num_rows
                    soma_total = pd.to_numeric(data_rows.iloc[:, 3], errors='coerce').sum()
                    
                    worksheet.write(total_row, 2, "TOTAL", fmt_total_label)
                    worksheet.write(total_row, 3, soma_total, fmt_total_value)

            # Finalização
            output.seek(0)
            st.success("✅ Processamento concluído! Baixe seu arquivo abaixo.")
            
            st.download_button(
                label="📥 Baixar Planilha Pronta",
                data=output,
                file_name="Bens_Moveis_Processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Ocorreu um erro: {e}")
