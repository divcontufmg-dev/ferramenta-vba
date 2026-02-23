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
# CONFIGURAÇÃO INICIAL E MEMÓRIA
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

if 'arquivos_memoria' not in st.session_state:
    st.session_state['arquivos_memoria'] = {}

# ==========================================
# FUNÇÕES E CLASSES
# ==========================================
def carregar_macro(nome_arquivo):
    try:
        with open(nome_arquivo, "r", encoding="utf-8") as f: return f.read()
    except:
        try:
            with open(nome_arquivo, "r", encoding="latin-1") as f: return f.read()
        except: return "Erro: Arquivo da macro não encontrado."

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
        if pd.isna(v): return ""
        s = str(v).strip()
        if s.endswith('.0'): s = s[:-2]
        s = re.sub(r'\D', '', s) # Garante que só sobram números (remove pontos e traços)
        return s
    except: return ""

def extrair_chave_vinculo(codigo_str):
    try:
        s = str(codigo_str).strip()
        if len(s) >= 2: return int(s[-2:])
        return int(s)
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
    header_rows.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False, header=False)
    data_rows.to_excel(writer, sheet_name=sheet_name, startrow=7, startcol=0, index=False, header=False)

    worksheet = writer.sheets[sheet_name]
    workbook = writer.book

    fmt_currency = workbook.add_format({'num_format': '#,##0.00'})
    fmt_total_label = workbook.add_format({'bold': True, 'align': 'right'})
    fmt_total_value = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'top': 1})
    fmt_red = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}) 
    fmt_blue = workbook.add_format({'bg_color': '#0000FF', 'font_color': '#FFFFFF'})

    worksheet.set_column('A:A', 40) 
    worksheet.set_column('B:C', 15) 
    worksheet.set_column('D:D', 18, fmt_currency)

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
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📊 Ferramenta Unificada: Conciliador RMB x SIAFI")
st.markdown("---")

with st.expander("📘 GUIA DE USO E MACROS (Clique para abrir)", expanded=False):
    st.markdown("### 🚀 Passo a Passo Completo")
    col_tut1, col_tut2 = st.columns(2)
    with col_tut1:
        st.info("💻 **Fase 1: Preparação no Excel (Opcional)**")
        macro1_content = carregar_macro("macro_preparar.txt")
        st.download_button("📥 Baixar Macro 1: Preparar (.txt)", macro1_content, "Macro_1_Preparar.txt", "text/plain")
        macro2_content = carregar_macro("macro_dividir.txt")
        st.download_button("📥 Baixar Macro 2: Dividir (.txt)", macro2_content, "Macro_2_Dividir.txt", "text/plain")

    with col_tut2:
        st.success("🤖 **Fase 2: Na ferramenta (Automático)**")
        st.markdown("1. Carregue a Planilha SIAFI completa e os PDFs abaixo.\n2. O sistema fará todo o processamento em memória e auditoria.")

st.subheader("📂 1. Carregar Arquivos")
col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    uploaded_siafi = st.file_uploader("Planilha Principal SIAFI (.xlsx)", type=["xlsx"])
with col_upload2:
    uploaded_pdfs = st.file_uploader("Relatórios RMB (.pdf)", accept_multiple_files=True, type=['pdf'])

st.markdown("---")

# ==========================================
# PASSO 1: PROCESSAR PLANILHAS (MEMÓRIA)
# ==========================================
st.subheader("⚙️ 2. Preparação dos Dados")

if st.button("▶️ 1º Passo: Processar SIAFI na Memória", type="secondary", use_container_width=True):
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo 'MATRIZ.xlsx' não foi encontrado no sistema.")
    elif uploaded_siafi is None:
        st.error("⚠️ Por favor, carregue a Planilha Principal SIAFI antes.")
    else:
        with st.spinner("Lendo planilhas, aplicando filtros e a MATRIZ..."):
            try:
                df_matriz = pd.read_excel("MATRIZ.xlsx", usecols="A:B", header=None)
                df_matriz.columns = ['Chave', 'Descricao']
                df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
                lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))

                xls_file = pd.ExcelFile(uploaded_siafi)
                processed_sheets = []

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
                    cols = list(data_rows.columns)
                    if 'Nova_Descricao' in cols: cols.insert(0, cols.pop(cols.index('Nova_Descricao')))
                    data_rows = data_rows[cols]
                    data_rows = data_rows.sort_values(by='Nova_Descricao', ascending=True)

                    processed_sheets.append({'name': sheet_name, 'header': header_rows, 'data': data_rows})

                st.session_state['arquivos_memoria'] = {}
                
                for item in processed_sheets:
                    single_excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(single_excel_buffer, engine='xlsxwriter') as single_writer:
                        formatar_aba(single_writer, item['name'], item['data'], item['header'])
                    single_excel_buffer.seek(0)
                    st.session_state['arquivos_memoria'][f"{item['name']}.xlsx"] = single_excel_buffer

                st.success(f"✅ Sucesso! {len(st.session_state['arquivos_memoria'])} abas prontas na memória.")
            except Exception as e:
                st.error(f"❌ Erro ao processar: {e}")

# ==========================================
# VISUALIZADOR DA MEMÓRIA
# ==========================================
if st.session_state.get('arquivos_memoria'):
    st.markdown("---")
    st.subheader("👀 3. Visualizador de Arquivos")
    
    nomes_arquivos = list(st.session_state['arquivos_memoria'].keys())
    arquivo_selecionado = st.selectbox("Selecione a aba para visualizar a estrutura:", nomes_arquivos)
    
    if arquivo_selecionado:
        buffer = st.session_state['arquivos_memoria'][arquivo_selecionado]
        buffer.seek(0)
        df_visualizacao = pd.read_excel(buffer, header=None)
        buffer.seek(0)
        st.dataframe(df_visualizacao, use_container_width=True)

    st.markdown("---")
    
    # ==========================================
    # PASSO 2: CONCILIAÇÃO
    # ==========================================
    st.subheader("🔍 4. Conciliação Automática")
    
    if st.button("▶️ 2º Passo: Iniciar Auditoria", type="primary", use_container_width=True):
        if not uploaded_pdfs:
            st.warning("⚠️ Faltam os relatórios RMB (.pdf) para conciliar.")
        else:
            progresso = st.progress(0)
            status_text = st.empty()
            
            pdfs = {f.name: f for f in uploaded_pdfs}
            pares = []
            logs = []

            for name_ex, file_ex in st.session_state['arquivos_memoria'].items():
                match = re.search(r'^(\d+)', name_ex)
                if match:
                    ug = match.group(1)
                    pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                    if pdf_match: pares.append({'ug': ug, 'excel': file_ex, 'pdf': pdf_match})
                    else: logs.append(f"⚠️ UG {ug}: Planilha encontrada, mas falta o PDF correspondente.")
            
            if not pares:
                st.error("❌ Nenhum par completo identificado.")
            else:
                pdf_out = PDF_Report()
                pdf_out.add_page()

                for idx, par in enumerate(pares):
                    ug = par['ug']
                    status_text.text(f"Processando Unidade Gestora: {ug}...")
                    
                    with st.container():
                        st.info(f"🏢 **Unidade Gestora: {ug}**")
                        
                        # === LEITURA EXCEL ===
                        df_padrao = pd.DataFrame()
                        saldo_2042 = 0.0
                        tem_2042_com_saldo = False
                        df_dados = pd.DataFrame() # Inicializa vazio para evitar erros no Raio-X
                        
                        try:
                            par['excel'].seek(0)
                            df_excel = pd.read_excel(par['excel'], header=None)
                            
                            if len(df_excel.columns) >= 4:
                                df_dados = df_excel.iloc[7:].copy()
                                
                                # Limpeza forçada (Extração Inteligente)
                                df_dados['Codigo_Limpo'] = df_dados.iloc[:, 1].astype(str).apply(limpar_codigo_bruto)
                                df_dados['Descricao_Excel'] = df_dados.iloc[:, 0].astype(str).str.strip().str.upper()
                                df_dados['Valor_Limpo'] = df_dados.iloc[:, 3].apply(limpar_valor)
                                
                                # Conta 2042 (Mantida por segurança)
                                mask_2042 = df_dados['Codigo_Limpo'] == '2042'
                                if mask_2042.any():
                                    saldo_2042 = df_dados.loc[mask_2042, 'Valor_Limpo'].sum()
                                    if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                                
                                # A MAGIA ACONTECE AQUI: Em vez de barrar apenas "449", aceitamos TUDO que tem um código preenchido!
                                df_filtrado = df_dados[df_dados['Codigo_Limpo'] != ''].copy()
                                
                                df_filtrado['Chave_Vinculo'] = df_filtrado['Codigo_Limpo'].apply(extrair_chave_vinculo)
                                
                                df_padrao = df_filtrado.groupby('Chave_Vinculo').agg({
                                    'Valor_Limpo': 'sum',
                                    'Descricao_Excel': 'first'
                                }).reset_index()
                                df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                        except Exception as e: logs.append(f"❌ Erro Excel UG {ug}: {e}")

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
                                    
                                    if txt and re.search(r'\d{1,3}(?:[.,]\d{3})*[.,]\d{2}', txt): tem_dados_validos = True
                                    
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
                                        line = line.strip()
                                        if re.match(r'^"?\d+', line):
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
                        except Exception as e: logs.append(f"❌ Erro Leitura PDF UG {ug}: {e}")

                        # === CRUZAMENTO ===
                        if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                        if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                        final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                        
                        final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if pd.notna(x['Descricao_Completa']) and str(x['Descricao_Completa']).strip() != '0' else "ITEM SEM DESCRIÇÃO", axis=1)
                        final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                        divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                        # === EXIBIÇÃO E RAIO-X ===
                        soma_pdf = final['Saldo_PDF'].sum()
                        soma_excel = final['Saldo_Excel'].sum()
                        dif_total = soma_pdf - soma_excel

                        # Painel de Debug Atualizado
                        with st.expander("🛠️ Raio-X da Extração (Veja o que o sistema leu)"):
                            st.write(f"**EXCEL:** Linhas totais lidas na Tabela: `{len(df_dados)}`")
                            st.write(f"**EXCEL:** Contas válidas conciliadas: `{len(df_padrao)}`")
                            st.write(f"**PDF:** Contas válidas extraídas do arquivo: `{len(df_pdf_final)}`")

                        col1, col2, col3 = st.columns(3)
                        col1.metric("Total RMB (PDF)", f"R$ {soma_pdf:,.2f}")
                        col2.metric("Total SIAFI (Excel)", f"R$ {soma_excel:,.2f}")
                        col3.metric("Diferença", f"R$ {dif_total:,.2f}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
                        
                        if not divergencias.empty:
                            st.warning(f"⚠️ Atenção: {len(divergencias)} conta(s) com divergência.")
                            with st.expander("Ver Detalhes das Divergências"):
                                st.dataframe(divergencias[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']])
                        else: st.success("✅ Tudo certo! Nenhuma divergência encontrada.")

                        if tem_2042_com_saldo: st.warning(f"ℹ️ Conta de Estoque Interno tem saldo: R$ {saldo_2042:,.2f}")
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

                status_text.text("Processamento concluído!")
                progresso.empty()
                
                if logs:
                    with st.expander("⚠️ Avisos do Sistema"):
                        for log in logs: st.write(log)
                
                try:
                    pdf_bytes = bytes(pdf_out.output())
                    st.download_button("BAIXAR RELATÓRIO PDF FINAL", pdf_bytes, "RELATORIO_FINAL_CONCILIACAO.pdf", "application/pdf", type="primary", use_container_width=True)
                except Exception as e: st.error(f"Erro no download: {e}")
