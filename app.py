import streamlit as st
import pandas as pd
import pdfplumber
import re
from fpdf import FPDF, XPos, YPos
import io
import os
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
from pytesseract import Output
import xlsxwriter

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Inicializa o cofre na memória do Streamlit
if 'arquivos_memoria' not in st.session_state:
    st.session_state['arquivos_memoria'] = {}

# ==========================================
# FUNÇÕES E CLASSES
# ==========================================

def limpar_valor(v):
    if v is None or pd.isna(v): return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).replace('"', '').replace("'", "").strip()
    if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
    elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
    try: return float(re.sub(r'[^\d.-]', '', v))
    except: return 0.0

def limpar_codigo_bruto(v):
    try:
        s = str(v).strip()
        if s.endswith('.0'): s = s[:-2]
        return s
    except: return ""

def extrair_chave_vinculo(codigo_str):
    try: return int(codigo_str[-2:])
    except: return 0

def formatar_real(valor):
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de conferência patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15); self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

def formatar_aba(writer, sheet_name, data_rows, header_rows):
    # Escreve Cabeçalho (Deslocado 1 coluna para direita)
    header_rows.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False, header=False)
    # Escreve Dados (Começando na coluna A, linha 8)
    data_rows.to_excel(writer, sheet_name=sheet_name, startrow=7, startcol=0, index=False, header=False)

    worksheet = writer.sheets[sheet_name]
    workbook = writer.book

    fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
    fmt_total_label = workbook.add_format({'bold': True, 'align': 'right'})
    fmt_total_value = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
    fmt_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}) 
    fmt_blue = workbook.add_format({'bg_color': '#0000FF', 'font_color': '#FFFFFF'})

    worksheet.set_column('A:A', 40) # Nova Descrição
    worksheet.set_column('B:C', 15) # Conta e Descrição Original
    worksheet.set_column('D:D', 18, fmt_currency) # Valor

    num_rows = len(data_rows)
    start_row_excel = 7 
    
    for i in range(num_rows):
        val_conta = data_rows.iloc[i, 1] 
        val_valor = data_rows.iloc[i, 3] 
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

    total_row = start_row_excel + num_rows
    soma_total = pd.to_numeric(data_rows.iloc[:, 3], errors='coerce').sum()
    worksheet.write(total_row, 2, "TOTAL", fmt_total_label)
    worksheet.write(total_row, 3, soma_total, fmt_total_value)

# ==========================================
# INTERFACE
# ==========================================
st.title("📊 Ferramenta Definitiva: Conciliador RMB x SIAFI")
st.markdown("A ferramenta agora processa, separa em memória e concilia tudo num só clique!")
st.markdown("---")

col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ Planilha SIAFI")
    uploaded_siafi = st.file_uploader("Carregar Planilha Principal (.xlsx)", type=["xlsx"])

with col2:
    st.subheader("2️⃣ Relatórios RMB")
    uploaded_pdfs = st.file_uploader("Carregar PDFs (Sintético Patrimonial)", accept_multiple_files=True, type=["pdf"])

st.markdown("---")

if st.button("▶️ Iniciar Processamento e Conciliação Automática", use_container_width=True, type="primary"):
    
    if not uploaded_siafi or not uploaded_pdfs:
        st.warning("⚠️ Por favor, adicione a Planilha SIAFI e os arquivos PDF antes de processar.")
    elif not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo 'MATRIZ.xlsx' não foi encontrado no sistema.")
    else:
        status_text = st.empty()
        progresso = st.progress(0)
        logs = []

        # =========================================================================
        # FASE 1: PREPARAR E ALOCAR NA MEMÓRIA
        # =========================================================================
        status_text.text("⚙️ FASE 1: Preparando a planilha SIAFI e alocando na memória...")
        
        try:
            df_matriz = pd.read_excel("MATRIZ.xlsx", usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao']
            df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
            lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))

            xls_file = pd.ExcelFile(uploaded_siafi)
            st.session_state['arquivos_memoria'] = {} 

            for sheet_name in xls_file.sheet_names:
                if sheet_name == "MATRIZ": continue

                df_raw = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
                if len(df_raw) < 8: continue 

                header_rows = df_raw.iloc[:7]
                data_rows = df_raw.iloc[7:].copy()

                data_rows[0] = pd.to_numeric(data_rows[0], errors='coerce')
                exclusion_list = [123110703, 123110402, 123119910]
                data_rows = data_rows[~data_rows[0].isin(exclusion_list)]

                data_rows['Nova_Descricao'] = data_rows[0].map(lookup_dict)

                # Reordena (A descrição da Matriz vai para a coluna 0, a Conta vai para 1)
                cols = list(data_rows.columns)
                if 'Nova_Descricao' in cols:
                    cols.insert(0, cols.pop(cols.index('Nova_Descricao')))
                data_rows = data_rows[cols]
                data_rows = data_rows.sort_values(by='Nova_Descricao', ascending=True)

                # Gera o arquivo Excel em memória
                single_excel_buffer = io.BytesIO()
                with pd.ExcelWriter(single_excel_buffer, engine='xlsxwriter') as single_writer:
                    formatar_aba(single_writer, sheet_name, data_rows, header_rows)
                
                single_excel_buffer.seek(0)
                st.session_state['arquivos_memoria'][f"{sheet_name}.xlsx"] = single_excel_buffer
                
        except Exception as e:
            st.error(f"❌ Erro na Fase 1: {e}")
            st.stop()

        # =========================================================================
        # FASE 2: CONCILIAÇÃO
        # =========================================================================
        status_text.text("⚙️ FASE 2: Conciliando relatórios...")
        
        pdfs = {f.name: f for f in uploaded_pdfs}
        excels_memoria = st.session_state['arquivos_memoria']
        pares = []

        # Cruzar abas geradas em memória com os PDFs
        for name_ex, file_ex in excels_memoria.items():
            match = re.search(r'(\d+)', name_ex)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if ug in n), None)
                if pdf_match:
                    pares.append({'ug': ug, 'nome_aba': name_ex, 'excel': file_ex, 'pdf': pdf_match})
                else:
                    logs.append(f"⚠️ Aba '{name_ex}' (UG {ug}): Faltando PDF correspondente.")

        if not pares:
            st.error("❌ Nenhum par completo (Aba SIAFI + PDF) foi identificado.")
        else:
            pdf_out = PDF_Report()
            pdf_out.add_page()
            st.markdown("---")
            st.subheader("🔍 Resultados da Análise")

            for idx, par in enumerate(pares):
                ug = par['ug']
                status_text.text(f"Processando Unidade Gestora: {ug}...")
                
                with st.container():
                    st.info(f"🏢 **Unidade Gestora: {ug}**")
                    
                    # === LEITURA EXCEL (MEMÓRIA) ===
                    df_padrao = pd.DataFrame()
                    saldo_2042 = 0.0
                    tem_2042_com_saldo = False
                    
                    try:
                        par['excel'].seek(0)
                        df_aba = pd.read_excel(par['excel'], header=None)
                        df_data = df_aba.iloc[7:].copy()
                        
                        if len(df_data.columns) >= 4:
                            df_calc = pd.DataFrame()
                            
                            # Após a formatação da Fase 1, o Excel gerado tem:
                            # Col 0 (A): Nova Descrição
                            # Col 1 (B): Conta
                            # Col 2 (C): Descrição Original
                            # Col 3 (D): Valor Monetário
                            df_calc['Codigo_Limpo'] = df_data.iloc[:, 1].apply(limpar_codigo_bruto)
                            df_calc['Descricao_Excel'] = df_data.iloc[:, 0].astype(str).str.strip().str.upper()
                            df_calc['Valor_Limpo'] = df_data.iloc[:, 3].apply(limpar_valor)
                            
                            mask_2042 = df_calc['Codigo_Limpo'] == '2042'
                            if mask_2042.any():
                                saldo_2042 = df_calc.loc[mask_2042, 'Valor_Limpo'].sum()
                                if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                            
                            mask_padrao = df_calc['Codigo_Limpo'].str.startswith('449')
                            df_dados = df_calc[mask_padrao].copy()
                            df_dados['Chave_Vinculo'] = df_dados['Codigo_Limpo'].apply(extrair_chave_vinculo)
                            
                            df_padrao = df_dados.groupby('Chave_Vinculo').agg({
                                'Valor_Limpo': 'sum',
                                'Descricao_Excel': 'first'
                            }).reset_index()
                            df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                    except Exception as e:
                        logs.append(f"❌ Erro Excel UG {ug}: {e}")

                    # === LEITURA PDF ===
                    df_pdf_final = pd.DataFrame()
                    dados_pdf = []
                    
                    try:
                        par['pdf'].seek(0)
                        pdf_bytes = par['pdf'].read()
                        
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                            for page in p_doc.pages:
                                txt = page.extract_text()
                                is_ocr = False
                                tem_dados_validos = False
                                
                                if txt and re.search(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}', txt):
                                    tem_dados_validos = True
                                
                                if not txt or not tem_dados_validos or len(txt) < 50:
                                    is_ocr = True
                                    try:
                                        imagens = convert_from_bytes(pdf_bytes, first_page=page.page_number, last_page=page.page_number, dpi=300)
                                        if imagens:
                                            img = imagens[0]
                                            try:
                                                osd = pytesseract.image_to_osd(img, output_type=Output.DICT)
                                                if osd['rotate'] != 0: img = img.rotate(-osd['rotate'], expand=True)
                                            except: pass
                                            txt = pytesseract.image_to_string(img, lang='por', config='--psm 6')
                                    except Exception: pass

                                if not txt: continue
                                if "SINTÉTICO PATRIMONIAL" not in txt.upper(): continue
                                if "DE ENTRADAS" in txt.upper() or "DE SAÍDAS" in txt.upper(): continue

                                for line in txt.split('\n'):
                                    if re.match(r'^"?\d+"?\s+', line):
                                        vals = []
                                        if is_ocr:
                                            vals_raw = re.findall(r'([\d\.\s]+,\d{2})', line)
                                            vals = [v.replace(' ', '') for v in vals_raw]
                                        else:
                                            vals = re.findall(r'([0-9]{1,3}(?:[.,][0-9]{3})*[.,]\d{2})', line)
                                        
                                        if len(vals) >= 4:
                                            chave_match = re.match(r'^"?(\d+)', line)
                                            if chave_match:
                                                chave_raw = chave_match.group(1)
                                                dados_pdf.append({
                                                    'Chave_Vinculo': int(chave_raw),
                                                    'Saldo_PDF': limpar_valor(vals[-4])
                                                })
                        
                        if dados_pdf:
                            df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
                    except Exception as e:
                        logs.append(f"❌ Erro Leitura PDF UG {ug}: {e}")

                    # === CRUZAMENTO ===
                    if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                    if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                    final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                    final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if x['Descricao_Completa'] != 0 else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
                    final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                    divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                    # === EXIBIÇÃO ===
                    soma_pdf = final['Saldo_PDF'].sum()
                    soma_excel = final['Saldo_Excel'].sum()
                    dif_total = soma_pdf - soma_excel

                    col_m1, col_m2, col_m3 = st.columns(3)
                    col_m1.metric("Total RMB (PDF)", f"R$ {soma_pdf:,.2f}")
                    col_m2.metric("Total SIAFI (Excel)", f"R$ {soma_excel:,.2f}")
                    col_m3.metric("Diferença", f"R$ {dif_total:,.2f}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
                    
                    if not divergencias.empty:
                        st.warning(f"⚠️ Atenção: {len(divergencias)} conta(s) com divergência.")
                        with st.expander("Ver Detalhes das Divergências"):
                            st.dataframe(divergencias[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']])
                    else:
                        st.success("✅ Tudo certo! Nenhuma divergência encontrada.")

                    if tem_2042_com_saldo:
                        st.warning(f"ℹ️ Conta de Estoque Interno tem saldo: R$ {saldo_2042:,.2f}")

                    st.markdown("---")

                    # === GERAÇÃO PDF FINAL ===
                    pdf_out.set_font("helvetica", 'B', 11)
                    pdf_out.set_fill_color(240, 240, 240)
                    pdf_out.cell(0, 10, text=f"Unidade Gestora: {ug}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, fill=True)
                    
                    if not divergencias.empty:
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.set_fill_color(255, 200, 200)
                        pdf_out.cell(15, 8, "Item", 1, fill=True)
                        pdf_out.cell(85, 8, "Descrição da Conta", 1, fill=True)
                        pdf_out.cell(30, 8, "SALDO RMB", 1, fill=True)
                        pdf_out.cell(30, 8, "SALDO SIAFI", 1, fill=True)
                        pdf_out.cell(30, 8, "Diferença", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                        
                        pdf_out.set_font("helvetica", '', 8)
                        for _, row in divergencias.iterrows():
                            pdf_out.cell(15, 7, str(int(row['Chave_Vinculo'])), 1)
                            pdf_out.cell(85, 7, str(row['Descricao'])[:48], 1)
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_PDF']), 1)
                            pdf_out.cell(30, 7, formatar_real(row['Saldo_Excel']), 1)
                            pdf_out.set_text_color(200, 0, 0)
                            pdf_out.cell(30, 7, formatar_real(row['Diferenca']), 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                            pdf_out.set_text_color(0, 0, 0)
                    else:
                        pdf_out.set_font("helvetica", 'I', 9)
                        pdf_out.cell(0, 8, "Nenhuma divergência encontrada.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                    if tem_2042_com_saldo:
                        pdf_out.ln(2)
                        pdf_out.set_font("helvetica", 'B', 9)
                        pdf_out.set_fill_color(255, 255, 200)
                        pdf_out.cell(100, 8, "SALDO ESTOQUE INTERNO", 1, fill=True)
                        pdf_out.cell(90, 8, f"R$ {formatar_real(saldo_2042)}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                    pdf_out.ln(2)
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(220, 230, 241)
                    pdf_out.cell(100, 8, "TOTAIS", 1, fill=True)
                    pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True)
                    pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True)
                    if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
                    pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf_out.set_text_color(0, 0, 0)
                    pdf_out.ln(5)
                
                progresso.progress((idx + 1) / len(pares))

            status_text.text("Processamento concluído com sucesso!")
            progresso.empty()
            
            if logs:
                with st.expander("⚠️ Avisos"):
                    for log in logs: st.write(log)
            
            try:
                pdf_bytes = bytes(pdf_out.output())
                st.download_button("BAIXAR RELATÓRIO FINAL EM PDF", pdf_bytes, "RELATORIO_FINAL.pdf", "application/pdf", type="primary", use_container_width=True)
            except Exception as e:
                st.error(f"Erro no download do PDF: {e}")
