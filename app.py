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

# ==========================================
# CONFIGURAÇÃO INICIAL
# ==========================================
st.set_page_config(
    page_title="Conciliador RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES E CLASSES ORIGINAIS
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

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📊 Ferramenta Unificada: Conciliador RMB x SIAFI")
st.markdown("""
**Como funciona agora:**
Você não precisa mais baixar macros ou dividir os arquivos. Basta subir a **Planilha SIAFI completa**, os **Relatórios PDF do RMB**, garantir que a `MATRIZ.xlsx` está na mesma pasta do sistema, e clicar em processar. O sistema extrai e cruza os dados diretamente.
""")
st.markdown("---")

st.subheader("📂 Carregar Arquivos")
col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    uploaded_siafi = st.file_uploader("1. Planilha Principal SIAFI (.xlsx)", type=["xlsx"])
with col_upload2:
    uploaded_pdfs = st.file_uploader("2. Relatórios RMB (.pdf)", accept_multiple_files=True, type=['pdf'])

st.markdown("---")

# ==========================================
# MOTOR UNIFICADO DE PROCESSAMENTO
# ==========================================
if st.button("🚀 Iniciar Conciliação Completa", type="primary", use_container_width=True):
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo 'MATRIZ.xlsx' não foi encontrado no sistema.")
    elif uploaded_siafi is None:
        st.warning("⚠️ Por favor, carregue a Planilha Principal SIAFI.")
    elif not uploaded_pdfs:
        st.warning("⚠️ Faltam os relatórios RMB (.pdf) para conciliar.")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        # 1. CARREGAR A MATRIZ
        try:
            df_matriz = pd.read_excel("MATRIZ.xlsx", usecols="A:B", header=None)
            df_matriz.columns = ['Chave', 'Descricao']
            df_matriz = df_matriz.drop_duplicates(subset=['Chave'], keep='first')
            lookup_dict = dict(zip(df_matriz['Chave'], df_matriz['Descricao']))
        except Exception as e:
            st.error(f"❌ Erro ao ler a MATRIZ.xlsx: {e}")
            st.stop()

        # 2. PREPARAR DICIONÁRIO DE ARQUIVOS
        pdfs = {f.name: f for f in uploaded_pdfs}
        pares = []
        logs = []

        try:
            xls_file = pd.ExcelFile(uploaded_siafi)
        except Exception as e:
            st.error(f"❌ Erro ao ler o arquivo SIAFI: {e}")
            st.stop()

        # Mapeamento das UGs
        for sheet_name in xls_file.sheet_names:
            if sheet_name == "MATRIZ": continue
            match = re.search(r'^(\d+)', sheet_name)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match: 
                    pares.append({'ug': ug, 'sheet_name': sheet_name, 'pdf': pdf_match})
                else: 
                    logs.append(f"⚠️ UG {ug}: Aba encontrada no SIAFI, mas falta o PDF correspondente.")

        if not pares:
            st.error("❌ Nenhum par completo (Aba do SIAFI + PDF correspondente) foi identificado.")
        else:
            pdf_out = PDF_Report()
            pdf_out.add_page()
            
            st.subheader("🔍 Resultados da Análise")

            for idx, par in enumerate(pares):
                ug = par['ug']
                sheet_name = par['sheet_name']
                status_text.text(f"Processando Unidade Gestora: {ug}...")
                
                with st.container():
                    st.info(f"🏢 **Unidade Gestora: {ug}**")
                    
                    # === LEITURA DIRETA DO SIAFI ===
                    df_padrao = pd.DataFrame()
                    saldo_2042 = 0.0
                    tem_2042_com_saldo = False
                    
                    try:
                        df_raw = pd.read_excel(xls_file, sheet_name=sheet_name, header=None)
                        if len(df_raw) >= 8:
                            df_dados = df_raw.iloc[7:].copy() # A partir da linha 8
                            
                            # Coluna 0 é a Conta, Coluna 3 é o Valor (Conforme planilha original)
                            df_dados['Conta_Num'] = pd.to_numeric(df_dados[0], errors='coerce')
                            
                            # Filtro de Exclusão Original
                            exclusion_list = [123110703, 123110402, 123119910]
                            df_dados = df_dados[~df_dados['Conta_Num'].isin(exclusion_list)].copy()
                            
                            df_dados['Codigo_Limpo'] = df_dados[0].apply(limpar_codigo_bruto)
                            df_dados['Valor_Limpo'] = df_dados[3].apply(limpar_valor)
                            
                            # PROCV Direto
                            df_dados['Descricao_Completa'] = df_dados['Conta_Num'].map(lookup_dict).fillna("ITEM SEM DESCRIÇÃO NO SIAFI")
                            
                            # Filtro 2042
                            mask_2042 = df_dados['Codigo_Limpo'] == '2042'
                            if mask_2042.any():
                                saldo_2042 = df_dados.loc[mask_2042, 'Valor_Limpo'].sum()
                                if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                            
                            # Filtro 449 (Original restabelecido)
                            mask_padrao = df_dados['Codigo_Limpo'].str.startswith('449')
                            df_filtrado = df_dados[mask_padrao].copy()
                            df_filtrado['Chave_Vinculo'] = df_filtrado['Codigo_Limpo'].apply(extrair_chave_vinculo)
                            
                            df_padrao = df_filtrado.groupby('Chave_Vinculo').agg({
                                'Valor_Limpo': 'sum',
                                'Descricao_Completa': 'first'
                            }).reset_index()
                            df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                    except Exception as e:
                        logs.append(f"❌ Erro leitura SIAFI UG {ug}: {e}")

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
                    
                    final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if pd.notna(x['Descricao_Completa']) and str(x['Descricao_Completa']).strip() != '0' else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
                    final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                    divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                    # === EXIBIÇÃO NO SISTEMA ===
                    soma_pdf = final['Saldo_PDF'].sum()
                    soma_excel = final['Saldo_Excel'].sum()
                    dif_total = soma_pdf - soma_excel

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

            status_text.text("Processamento concluído com sucesso!")
            progresso.empty()
            
            if logs:
                with st.expander("⚠️ Avisos do Sistema"):
                    for log in logs: st.write(log)
            
            try:
                pdf_bytes = bytes(pdf_out.output())
                st.download_button(
                    label="📥 BAIXAR RELATÓRIO PDF FINAL", 
                    data=pdf_bytes, 
                    file_name="RELATORIO_FINAL_CONCILIACAO.pdf", 
                    mime="application/pdf", 
                    type="primary", 
                    use_container_width=True
                )
            except Exception as e: st.error(f"Erro no download: {e}")
