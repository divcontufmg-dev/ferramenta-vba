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
    # 1. Carregar a MATRIZ para memória (Pandas é mais seguro para PROCV)
    df_matriz = pd.read_excel(matriz_file, header=None) # Assume col A e B
    # Cria dicionário {Chave: Valor} para busca rápida
    # Força a chave a ser string para garantir match, ou numero. Vamos tentar ambos.
    matriz_dict = dict(zip(df_matriz.iloc[:, 0], df_matriz.iloc[:, 1]))

    # 2. Carregar o arquivo Excel mantendo formatação (OpenPyXL)
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

        # --- ESTRATÉGIA HÍBRIDA: PANDAS PARA DADOS, OPENPYXL PARA FORMATO ---
        
        # 3. Ler os dados da planilha para o Pandas (apenas da linha 8 para baixo)
        data = []
        # Iterar apenas sobre as colunas relevantes para leitura inicial (ex: A até Z)
        # Nota: O VBA insere uma coluna A nova. Então os dados atuais começam na coluna A (que virará B).
        for row in ws.iter_rows(min_row=8, values_only=True):
            data.append(list(row))
        
        # Se não tem dados, pula
        if not data:
            continue
            
        df = pd.DataFrame(data)
        
        # Ajuste de colunas: O DataFrame cria colunas 0, 1, 2...
        # A coluna 0 do DF corresponde à coluna A atual do Excel (que virará B)
        # Vamos garantir que a coluna 0 seja tratada como número para filtros
        df[0] = df[0].apply(converter_para_numero)

        # 4. Filtros de Exclusão (VBA Step 5)
        # Valores: 123110703, 123110402, 44905287
        valores_proibidos = [123110703, 123110402, 44905287]
        df = df[~df[0].isin(valores_proibidos)] # O til (~) inverte a seleção (pega os que NÃO estão na lista)

        # 5. Criar a Coluna do PROCV (VBA Step 3)
        # Vamos criar uma nova coluna e inseri-la na posição 0 do DataFrame
        def aplicar_vlookup(valor_chave):
            # Tenta buscar como número, se não der, tenta como string, se não der, retorna #N/A
            res = matriz_dict.get(valor_chave)
            if res is None:
                res = matriz_dict.get(str(valor_chave))
            return res if res is not None else "#N/A"

        nova_coluna_a = df[0].apply(aplicar_vlookup)
        df.insert(0, 'Nova_A', nova_coluna_a) # Insere na primeira posição

        # 6. Ordenação (VBA Step 8)
        # Classificar pela nova coluna A (agora chamada 'Nova_A')
        df = df.sort_values(by='Nova_A', ascending=True)

        # 7. ESCREVER DE VOLTA NO EXCEL
        # Primeiro: Inserir a coluna física no Excel para empurrar a formatação
        ws.insert_cols(1) 
        
        # Limpar os dados antigos (das linhas 8 para baixo) para não sobrar lixo
        # Como inserimos uma coluna, a largura mudou, mas vamos reescrever célula a célula
        # A maneira mais segura é sobrescrever.
        
        rows_to_write = dataframe_to_rows(df, index=False, header=False)
        
        # Escrevendo dados processados e ordenados
        for r_idx, row in enumerate(rows_to_write, 8): # Começa na linha 8
            for c_idx, value in enumerate(row, 1): # Começa na coluna 1 (A)
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = value
                
                # REPLICAR FORMATAÇÃO DA LINHA 8 ORIGINAL (Opcional, mas bom para manter fontes)
                # No OpenPyXL isso é complexo, vamos focar no valor correto.

        # Atualizar last_row baseada nos novos dados
        new_last_row = 8 + len(df) - 1

        # Limpar linhas que sobraram abaixo (caso a nova tabela seja menor que a antiga)
        if new_last_row < ws.max_row:
            ws.delete_rows(new_last_row + 1, amount=(ws.max_row - new_last_row))

        # 8. Totais e Cores (Pós-Processamento)
        soma_d = 0
        
        # Iterar sobre as linhas escritas para formatar e somar
        for r in range(8, new_last_row + 1):
            # Converter Coluna B (indice 2) para garantir numero no Excel
            cell_b = ws.cell(row=r, column=2)
            if cell_b.value:
                cell_b.value = converter_para_numero(cell_b.value)
            
            # Somar Coluna D (indice 4)
            cell_d = ws.cell(row=r, column=4)
            val_d = converter_para_numero(cell_d.value)
            
            # Acumular soma se for número
            if isinstance(val_d, (int, float)):
                soma_d += val_d
            
            # Formatação Condicional (Vermelho/Azul)
            val_b_check = cell_b.value
            
            # Regra Vermelha: 123110801
            if val_b_check == 123110801 and val_d != 0:
                for c in range(2, 5): # B, C, D
                    ws.cell(row=r, column=c).fill = fill_red
            
            # Regra Azul: 123119905
            if val_b_check == 123119905 and val_d != 0:
                for c in range(2, 5):
                    ws.cell(row=r, column=c).fill = fill_blue

        # Escrever Totais Finais
        ws.cell(row=new_last_row + 1, column=3).value = "TOTAL" # Coluna C
        cell_total = ws.cell(row=new_last_row + 1, column=4)    # Coluna D
        cell_total.value = soma_d
        cell_total.number_format = '#,##0.00'

    # Salvar
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

if file_target and file_matriz:
    if st.button("🚀 Processar Corretamente"):
        try:
            processed = processar_planilha_robusta(file_target, file_matriz)
            st.success("Concluído! Dados tipados e ordenados.")
            st.download_button("Baixar Planilha", processed, "Resultado_Final.xlsx")
        except Exception as e:
            st.error(f"Erro: {e}")
