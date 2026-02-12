import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

st.set_page_config(page_title="Automação VBA -> Python", layout="wide")
st.title("🛠️ Automação de Bens Móveis (Versão Corrigida)")

col1, col2 = st.columns(2)
with col1:
    file_target = st.file_uploader("1. Planilha ALVO (.xlsx)", type=["xlsx"])
with col2:
    file_matriz = st.file_uploader("2. Planilha MATRIZ (.xlsx)", type=["xlsx"])

def converter_para_numero(valor):
    """Tenta converter qualquer coisa para float, se falhar devolve o original"""
    try:
        return float(valor)
    except (ValueError, TypeError):
        return valor

def processar_planilha_robusta(target_file, matriz_file):
    # 1. Carregar a MATRIZ
    df_matriz = pd.read_excel(matriz_file, header=None)
    # Criar dicionário {Chave: Valor}
    matriz_dict = dict(zip(df_matriz.iloc[:, 0], df_matriz.iloc[:, 1]))

    # 2. Carregar o arquivo alvo
    wb = load_workbook(target_file)
    
    # Estilos
    fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")

    for sheet_name in wb.sheetnames:
        if sheet_name == "MATRIZ":
            continue
            
        ws = wb[sheet_name]
        last_row_inicial = ws.max_row
        
        if last_row_inicial < 8:
            continue

        # --- PROCESSAMENTO ---
        
        # 3. Ler dados para o Pandas
        data = []
        for row in ws.iter_rows(min_row=8, values_only=True):
            data.append(list(row))
        
        if not data:
            continue
            
        df = pd.DataFrame(data)
        
        # Garantir que a coluna de busca (agora índice 0) seja tratada corretamente
        df[0] = df[0].apply(converter_para_numero)

        # 4. Filtros de Exclusão
        valores_proibidos = [123110703, 123110402, 44905287]
        df = df[~df[0].isin(valores_proibidos)]

        # 5. Criar a Coluna do PROCV
        def aplicar_vlookup(valor_chave):
            res = matriz_dict.get(valor_chave)
            if res is None:
                # Tenta buscar convertendo para string caso a chave no dict seja string
                res = matriz_dict.get(str(valor_chave))
            return res if res is not None else "#N/A"

        nova_coluna_a = df[0].apply(aplicar_vlookup)
        df.insert(0, 'Nova_A', nova_coluna_a)

        # 6. Ordenação (A CORREÇÃO ESTÁ AQUI)
        # Usamos key=str para ordenar tratando tudo como texto temporariamente
        # Isso evita o erro de comparar int com str ("#N/A")
        try:
            df = df.sort_values(by='Nova_A', key=lambda col: col.astype(str), ascending=True)
        except:
            # Fallback se der erro extremo, ordena sem key (assumindo tipos iguais) ou ignora
            pass

        # 7. ESCREVER DE VOLTA NO EXCEL
        ws.insert_cols(1) 
        
        rows_to_write = dataframe_to_rows(df, index=False, header=False)
        
        for r_idx, row in enumerate(rows_to_write, 8):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value

        new_last_row = 8 + len(df) - 1

        if new_last_row < ws.max_row:
            ws.delete_rows(new_last_row + 1, amount=(ws.max_row - new_last_row))

        # 8. Totais e Cores
        soma_d = 0
        
        for r in range(8, new_last_row + 1):
            cell_b = ws.cell(row=r, column=2)
            if cell_b.value:
                cell_b.value = converter_para_numero(cell_b.value)
            
            cell_d = ws.cell(row=r, column=4)
            val_d = converter_para_numero(cell_d.value)
            
            if isinstance(val_d, (int, float)):
                soma_d += val_d
            
            val_b_check = cell_b.value
            
            if val_b_check == 123110801 and val_d != 0:
                for c in range(2, 5): 
                    ws.cell(row=r, column=c).fill = fill_red
            
            if val_b_check == 123119905 and val_d != 0:
                for c in range(2, 5):
                    ws.cell(row=r, column=c).fill = fill_blue

        # Totais Finais
        ws.cell(row=new_last_row + 1, column=3).value = "TOTAL"
        cell_total = ws.cell(row=new_last_row + 1, column=4)
        cell_total.value = soma_d
        cell_total.number_format = '#,##0.00'

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if file_target and file_matriz:
    if st.button("🚀 Processar"):
        try:
            processed = processar_planilha_robusta(file_target, file_matriz)
            st.success("Sucesso!")
            st.download_button("Baixar Resultado", processed, "Planilha_Final.xlsx")
        except Exception as e:
            st.error(f"Erro Detalhado: {e}")
