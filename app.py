import streamlit as st
import pandas as pd
import io

# Configuração da página (Layout similar às suas outras ferramentas)
st.set_page_config(page_title="Processador de Bens Móveis", layout="wide")

st.title("📂 Processador de Planilha de Bens Móveis")
st.markdown("""
Esta ferramenta automatiza o tratamento da planilha de Bens Móveis, replicando as funções da macro:
* Realiza o PROCV com a base MATRIZ.
* Filtra contas específicas (exclusão).
* Aplica formatação condicional (Vermelho/Azul).
* Calcula totais.
""")

# --- BARRA LATERAL (Inputs) ---
st.sidebar.header("Carregar Arquivos")

uploaded_file = st.sidebar.file_uploader("1. Carregar Planilha Principal (.xlsx)", type=["xlsx"])
uploaded_matriz = st.sidebar.file_uploader("2. Carregar Planilha MATRIZ (.xlsx)", type=["xlsx"])

# Botão de processamento
if st.sidebar.button("Processar Planilhas"):
    if uploaded_file is None or uploaded_matriz is None:
        st.error("Por favor, faça o upload de AMBOS os arquivos (Principal e MATRIZ).")
    else:
        try:
            # --- 1. CARREGAMENTO DOS DADOS ---
            # Lê a planilha MATRIZ (Assume que os dados estão na primeira aba, colunas A e B)
            # A macro usa colunas A e B da Matriz. A=Chave, B=Valor a retornar.
            df_matriz = pd.read_excel(uploaded_matriz, usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao'] # Renomeando para facilitar

            # Lê o arquivo principal (todas as abas)
            xls_file = pd.ExcelFile(uploaded_file)
            
            # Buffer para salvar o arquivo final em memória
            output = io.BytesIO()

            # Usamos XlsxWriter como engine para poder pintar as células
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # --- DEFINIÇÃO DE FORMATOS (ESTILOS) ---
                fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
                fmt_total_label = workbook.add_format({'bold': True, 'align': 'right'})
                fmt_total_value = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
                
                # Formatos condicionais (Fundo Vermelho e Azul)
                # O VBA pinta de B até D.
                fmt_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}) # Vermelho
                fmt_blue = workbook.add_format({'bg_color': '#0000FF', 'font_color': '#FFFFFF'}) # Azul

                # Aba MATRIZ (Cria a aba Matriz no arquivo final, conforme a macro original)
                df_matriz.to_excel(writer, sheet_name='MATRIZ', index=False, header=False)

                # Loop por todas as abas do arquivo principal
                for sheet_name in xls_file.sheet_names:
                    # Ignora se a aba já se chamar MATRIZ (regra do VBA)
                    if sheet_name == "MATRIZ":
                        continue

                    # Lê a aba atual. Assume que começa na linha 1, mas os dados reais (VBA) começam na linha 8.
                    # Ajuste: Vamos ler tudo e tratar como DataFrame.
                    # O VBA insere coluna na A. A antiga A vira B. O PROCV usa a B.
                    # Portanto, vamos ler o arquivo original. Vamos assumir que a 'Chave' está na Coluna A original.
                    df = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
                    
                    # Identificar onde começam os dados. O VBA diz "A8". 
                    # Pandas é index 0, então linha 8 do Excel é index 7.
                    # Vamos separar o cabeçalho (linhas 0 a 6) do corpo (7 em diante).
                    header_rows = df.iloc[:7] # Linhas de 1 a 7 do Excel
                    data_rows = df.iloc[7:].copy() # Linhas de 8 em diante
                    
                    if data_rows.empty:
                        # Se não tiver dados, apenas copia a aba original
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        continue

                    # --- LÓGICA DE DADOS ---
                    
                    # 1. Converter coluna A original (que será a B) para Numérico
                    # A coluna A original é a coluna 0 no Pandas
                    data_rows[0] = pd.to_numeric(data_rows[0], errors='coerce')

                    # 2. Filtrar valores (Exclusão) - VBA passo 5
                    # Valores: 123110703, 123110402, 44905287
                    exclusion_list = [123110703, 123110402, 44905287]
                    data_rows = data_rows[~data_rows[0].isin(exclusion_list)]

                    # 3. PROCV (Merge) - VBA passo 3
                    # Cria a nova coluna de descrição baseada na chave (Col 0 do data_rows vs Chave da Matriz)
                    # O merge pode bagunçar a ordem, então usamos map para preservar índice se possível
                    lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))
                    data_rows['Nova_Coluna_A'] = data_rows[0].map(lookup_dict)

                    # 4. Reordenar colunas
                    # A nova coluna deve ser a primeira.
                    cols = list(data_rows.columns)
                    cols.insert(0, cols.pop(cols.index('Nova_Coluna_A')))
                    data_rows = data_rows[cols]

                    # 5. Classificar por Coluna A (Nova Descrição) - VBA passo 8
                    data_rows = data_rows.sort_values(by='Nova_Coluna_A', ascending=True)

                    # --- ESCRITA NO EXCEL ---
                    
                    # Escreve o cabeçalho original (mas precisa deslocar 1 coluna para direita pois inserimos uma nova)
                    # O VBA insere coluna A, empurrando tudo. Então o cabeçalho antigo A vira B.
                    # Escrevemos o cabeçalho começando da coluna 1 (B)
                    header_rows.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False, header=False)

                    # Escreve os dados processados começando da linha 7 (A8)
                    data_rows.to_excel(writer, sheet_name=sheet_name, startrow=7, startcol=0, index=False, header=False)
                    
                    # --- FORMATAÇÃO E TOTAL ---
                    worksheet = writer.sheets[sheet_name]
                    
                    # Pega o número de linhas de dados
                    num_rows = len(data_rows)
                    start_row = 7 # Linha 8 do Excel
                    end_row = start_row + num_rows

                    # Acessa os dados para verificar condições de cor
                    # No Excel final: Col A = Descrição, Col B = Conta (Chave), Col C = ?, Col D = Valor
                    # Pandas index: 'Nova_Coluna_A' é col 0, Coluna Original 0 (Conta) é col 1, ... Coluna Valor é col 3
                    
                    # Iterar sobre as linhas para pintar (VBA passo 9)
                    for row_idx in range(num_rows):
                        excel_row = start_row + row_idx
                        
                        # Pega valores para condicional
                        # iloc[row_idx, 1] -> Conta (Coluna B no Excel)
                        # iloc[row_idx, 3] -> Valor (Coluna D no Excel)
                        val_conta = data_rows.iloc[row_idx, 1] 
                        val_valor = data_rows.iloc[row_idx, 3]

                        # Garante que é número para comparação
                        try:
                            val_valor = float(val_valor)
                        except:
                            val_valor = 0

                        # Condicional Vermelha (Conta 123110801 e Valor != 0)
                        if val_conta == 123110801 and val_valor != 0:
                            worksheet.set_row(excel_row, cell_format=fmt_red) # Pinta a linha toda ou intervalo específico
                            # O VBA pinta range B:D. Vamos ser específicos:
                            worksheet.write(excel_row, 1, val_conta, fmt_red) # B
                            worksheet.write(excel_row, 2, data_rows.iloc[row_idx, 2], fmt_red) # C
                            worksheet.write(excel_row, 3, val_valor, fmt_red) # D

                        # Condicional Azul (Conta 123119905 e Valor != 0)
                        elif val_conta == 123119905 and val_valor != 0:
                            worksheet.write(excel_row, 1, val_conta, fmt_blue) # B
                            worksheet.write(excel_row, 2, data_rows.iloc[row_idx, 2], fmt_blue) # C
                            worksheet.write(excel_row, 3, val_valor, fmt_blue) # D

                    # --- LINHA DE TOTAL (VBA Passo 6) ---
                    # Coluna C (índice 2): Escrever "TOTAL"
                    # Coluna D (índice 3): Somatório
                    
                    total_row = end_row
                    # Soma da coluna D (índice 3 no dataframe processado)
                    soma_total = pd.to_numeric(data_rows.iloc[:, 3], errors='coerce').sum()
                    
                    worksheet.write(total_row, 2, "TOTAL", fmt_total_label)
                    worksheet.write(total_row, 3, soma_total, fmt_total_value)

                    # Ajuste de largura (AutoFit simulado)
                    worksheet.set_column('A:A', 40) # Descrição
                    worksheet.set_column('B:C', 15) 
                    worksheet.set_column('D:D', 18, fmt_currency) # Valor com formatação

            # Finaliza o arquivo
            output.seek(0)
            
            st.success("PLANILHA DE BENS MÓVEIS ATUALIZADA COM ÊXITO!")
            
            st.download_button(
                label="📥 Baixar Planilha Processada",
                data=output,
                file_name="Bens_Moveis_Processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocorreu um erro durante o processamento: {e}")
