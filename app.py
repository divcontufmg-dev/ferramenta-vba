import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

st.set_page_config(page_title="VBA Converter Pro", layout="wide")
st.title("🚀 Conversor VBA para Compatibilidade com App RMB")

st.info("Esta ferramenta prepara a planilha exatamente como o VBA, garantindo que o seu App de Conciliação consiga ler os dados (Coluna E e Códigos).")

col1, col2 = st.columns(2)
with col1:
    file_target = st.file_uploader("1. Planilha ALVO (Para processar)", type=["xlsx"])
with col2:
    file_matriz = st.file_uploader("2. Planilha MATRIZ", type=["xlsx"])

def safe_float(val):
    """Converte para float de forma segura, mantendo zero se falhar"""
    try:
        return float(val)
    except:
        return 0.0

def processar_planilha_final(target_file, matriz_file):
    # 1. Carregar a MATRIZ (Referência para o PROCV)
    # Usamos dtype=str para garantir que o código (ex: "100") seja lido como texto, igual ao PROCV
    df_matriz = pd.read_excel(matriz_file, header=None, dtype=str) 
    # Dicionário: Chave (Col A) -> Valor (Col B)
    matriz_dict = dict(zip(df_matriz.iloc[:, 0].str.strip(), df_matriz.iloc[:, 1]))

    # 2. Carregar o arquivo ALVO usando Pandas (leitura bruta para manter estrutura)
    # header=None garante que lemos da linha 1 até o fim, sem perder nada
    df_original = pd.read_excel(target_file, header=None)

    # Criamos um buffer para salvar o resultado
    output = BytesIO()
    
    # Carregamos também com OpenPyXL para formatação final (Cores)
    wb = load_workbook(target_file)
    
    # Estilos de cor do VBA
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")

    for sheet_name in wb.sheetnames:
        if sheet_name == "MATRIZ":
            continue
            
        ws = wb[sheet_name]
        
        # Se a planilha for muito pequena, ignora (lógica do VBA lastRow < 8)
        if ws.max_row < 8:
            continue

        # --- FASE 1: MANIPULAÇÃO DE DADOS (PANDAS) ---
        
        # Ler a aba atual como DataFrame
        # Forçamos a leitura de todas as colunas para não "encolher" a planilha
        data = ws.values
        cols = next(data) # Pega primeira linha para definir largura
        df = pd.DataFrame(ws.values, columns=None) # Lê tudo sem cabeçalho

        # Separar Cabeçalho (Linhas 0 a 6 - ou seja, Excel 1 a 7) e Dados (Linhas 7+ - Excel 8+)
        df_header = df.iloc[:7].copy()
        df_dados = df.iloc[7:].copy()

        # Se não tiver dados, pula
        if df_dados.empty:
            continue

        # == LÓGICA DO VBA ==
        
        # 1. Filtros de Exclusão (Coluna B atual = Índice 1 no Pandas antes da inserção da coluna A?)
        # O VBA insere a coluna A PRIMEIRO. Depois verifica B.
        # Então B (novo) é a coluna A (antiga).
        # Vamos filtrar baseados na coluna 0 do DataFrame atual (que virará B)
        
        # Normalizar para texto para comparação
        col_check = df_dados.iloc[:, 0].astype(str).str.strip().str.replace('.0', '', regex=False)
        valores_proibidos = ["123110703", "123110402", "44905287"]
        df_dados = df_dados[~col_check.isin(valores_proibidos)]

        # 2. PROCV (Novo valor da Coluna A)
        # O VBA faz VLOOKUP(B8...). O B8 é a coluna 0 do nosso DF atual.
        def aplicar_procv(valor):
            chave = str(valor).strip().replace('.0', '')
            return matriz_dict.get(chave, "#N/A")
        
        nova_coluna_a = df_dados.iloc[:, 0].apply(aplicar_procv)
        
        # 3. Ordenação
        # O VBA ordena pela Coluna A (o resultado do PROCV).
        df_dados.insert(0, 'PROCV_TEMP', nova_coluna_a)
        df_dados = df_dados.sort_values(by='PROCV_TEMP', ascending=True)
        
        # 4. Construção da Tabela Final (Inserção da Coluna)
        # Inserimos uma coluna vazia no Header para alinhar
        df_header.insert(0, 'Spacer', None) 
        
        # Juntar Header + Dados
        df_final = pd.concat([df_header, df_dados])
        
        # --- FASE 2: ESCRITA NO EXCEL (OPENPYXL) ---
        
        # Limpar a planilha atual para receber os dados novos e limpos
        ws.delete_rows(1, ws.max_row)
        
        # Escrever linha a linha
        rows = dataframe_to_rows(df_final, index=False, header=False)
        
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                
                # CORREÇÃO CRÍTICA PARA O SEU APP:
                # O seu app espera ler números na coluna de valor.
                # A Coluna D do VBA (que é somada) é o índice 4 (D).
                # A Coluna E do seu App (que é lida) é o índice 5 (E).
                # Vamos garantir que D e E sejam números se parecerem números.
                if r_idx >= 8 and c_idx >= 4: # Colunas D, E, F...
                    try:
                        if value is not None:
                            # Tenta converter para float limpo
                            cell.value = float(value)
                    except:
                        pass # Deixa como texto se não for número

        # Atualizar last_row
        new_last_row = ws.max_row

        # --- FASE 3: TOTAIS E FORMATAÇÃO (Igual ao VBA) ---
        soma_d = 0
        
        for r in range(8, new_last_row + 1):
            # Lógica VBA: Pintar se B=123... e D <> 0
            # Agora que inserimos a coluna, indices: A=1, B=2, C=3, D=4
            
            try:
                # Pegar valores com segurança
                val_b = str(ws.cell(row=r, column=2).value).replace('.0', '').strip()
                val_d = ws.cell(row=r, column=4).value
                val_d_float = safe_float(val_d)
                
                # Somatório (Coluna D)
                soma_d += val_d_float
                
                # Cores
                tem_saldo = abs(val_d_float) > 0
                
                if val_b == "123110801" and tem_saldo:
                    for c in range(2, 5): # B até D
                        ws.cell(row=r, column=c).fill = fill_red
                        
                if val_b == "123119905" and tem_saldo:
                    for c in range(2, 5):
                        ws.cell(row=r, column=c).fill = fill_blue
            except Exception:
                continue

        # Escrever Totais (Linha Final)
        ws.cell(row=new_last_row + 1, column=3).value = "TOTAL"
        cell_total = ws.cell(row=new_last_row + 1, column=4)
        cell_total.value = soma_d
        cell_total.number_format = '#,##0.00'

    wb.save(output)
    output.seek(0)
    return output

# Botão de Ação
if file_target and file_matriz:
    if st.button("🔄 Executar Processo Completo"):
        try:
            processed = processar_planilha_final(file_target, file_matriz)
            st.success("Planilha processada! Pronta para o Conciliador.")
            st.download_button("📥 Baixar Planilha Pronta", processed, "Planilha_Formatada_RMB.xlsx")
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
