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
    page_title="Conciliador RMB x SIAFI (Motor Inteligente)",
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
# FUNÇÕES CORE (LÓGICA BLINDADA)
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

def extrair_chave_universal(codigo_original, dict_matriz):
    """
    Motor universal que garante a chave de 2 dígitos correta, 
    sem gerar números aleatórios.
    """
    codigo = str(codigo_original).strip()
    if not codigo or codigo == '0' or codigo == '2042':
        return None
        
    # 1. Já é código de 2 dígitos? (Ex: 07, 16)
    if len(codigo) <= 2 and codigo.isdigit():
        return int(codigo)
        
    # 2. Já começa com 449 ou 339? (Pega os 2 últimos dígitos)
    if (codigo.startswith('449') or codigo.startswith('339')) and len(codigo) >= 4:
        return int(codigo[-2:])
        
    # 3. Consulta a MATRIZ para traduzir o 123...
    if codigo in dict_matriz:
        valor_matriz = str(dict_matriz[codigo]).strip()
        match = re.search(r'((?:449|339)\d+)', valor_matriz)
        if match:
            return int(match.group(1)[-2:])
        if len(valor_matriz) <= 2 and valor_matriz.isdigit():
            return int(valor_matriz)
            
    return None

def identificar_estrutura_excel(df_raw):
    """
    Radar que descobre sozinho onde estão as colunas, 
    evitando que a descrição vire código numérico.
    """
    start_row = -1
    col_conta = -1
    col_desc = -1
    
    for i in range(min(15, len(df_raw))):
        val0 = str(df_raw.iloc[i, 0]).strip().replace('.0', '')
        val1 = str(df_raw.iloc[i, 1]).strip().replace('.0', '') if len(df_raw.columns) > 1 else ""
        
        # Procura a primeira linha onde tem número de conta válido
        if (val0.isdigit() and len(val0) >= 2) or (val1.isdigit() and len(val1) >= 2):
            start_row = i
            if val0.isdigit() and not val1.isdigit():
                col_conta = 0
                col_desc = 1
            elif val1.isdigit() and not val0.isdigit():
                col_conta = 1
                col_desc = 0
            else:
                col_conta = 0
                col_desc = 1
            break
            
    return start_row, col_conta, col_desc

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
st.title("📊 Conciliador RMB x SIAFI (Motor Inteligente)")
st.markdown("""
**Novo Radar de Colunas:** Pode enviar a planilha **Bruta** ou a que já passou pela **Fase 1**. O sistema descobre sozinho as colunas de Conta e Descrição sem embaralhar os dados!
""")
st.markdown("---")

col_upload1, col_upload2 = st.columns(2)
with col_upload1:
    uploaded_siafi = st.file_uploader("1. Planilha Principal SIAFI (.xlsx)", type=["xlsx"])
with col_upload2:
    uploaded_pdfs = st.file_uploader("2. Relatórios RMB (.pdf)", accept_multiple_files=True, type=['pdf'])

st.markdown("---")

# ==========================================
# PROCESSAMENTO PRINCIPAL
# ==========================================
if st.button("🚀 Iniciar Auditoria", type="primary", use_container_width=True):
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo 'MATRIZ.xlsx' não foi encontrado na pasta do sistema.")
    elif uploaded_siafi is None:
        st.warning("⚠️ Por favor, carregue a Planilha Principal SIAFI.")
    elif not uploaded_pdfs:
        st.warning("⚠️ Faltam os relatórios RMB (.pdf).")
    else:
        progresso = st.progress(0)
        status_text = st.empty()
        
        # 1. Carregar a Matriz
        try:
            df_matriz = pd.read_excel("MATRIZ.xlsx", header=None)
            # Garantir que a chave seja o 123... independente da ordem das colunas da matriz
            if '449' in str(df_matriz.iloc[0, 0]):
                lookup_dict = {str(k).strip(): str(v).strip() for k, v in zip(df_matriz[1], df_matriz[0])}
            else:
                lookup_dict = {str(k).strip(): str(v).strip() for k, v in zip(df_matriz[0], df_matriz[1])}
        except Exception as e:
            st.error(f"Erro ao ler a MATRIZ.xlsx: {e}")
            st.stop()

        # 2. Parear Arquivos
        pdfs = {f.name: f for f in uploaded_pdfs}
        pares = []
        logs = []

        try:
            xls_file = pd.ExcelFile(uploaded_siafi)
            for sheet_name in xls_file.sheet_names:
                if sheet_name == "MATRIZ": continue
                match = re.search(r'^(\d+)', sheet_name)
                if match:
                    ug = match.group(1)
                    pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                    if pdf_match: 
                        pares.append({'ug': ug, 'sheet_name': sheet_name, 'pdf': pdf_match})
                    else: 
                        logs.append(f"⚠️ UG {ug}: Aba encontrada no SIAFI, mas falta o PDF.")
        except Exception as e:
            st.error(f"Erro ao abrir o arquivo SIAFI: {e}")
            st.stop()

        if not pares:
            st.error("❌ Nenhum par completo foi identificado.")
        else:
            pdf_out = PDF_Report()
            pdf_out.add_page()
            st.subheader("🔍 Resultados da Análise")

            for idx, par in enumerate(pares):
                ug = par['ug']
                status_text.text(f"Processando Unidade Gestora: {ug}...")
                
                with st.container():
                    st.info(f"🏢 **Unidade Gestora: {ug}**")
                    
                    # ==========================================
                    # LEITURA DO EXCEL
                    # ==========================================
                    df_padrao = pd.DataFrame()
                    saldo_2042 = 0.0
                    tem_2042_com_saldo = False
                    
                    try:
                        df_raw = pd.read_excel(xls_file, sheet_name=par['sheet_name'], header=None)
                        
                        start_row, col_conta, col_desc = identificar_estrutura_excel(df_raw)
                        
                        if start_row != -1:
                            df_dados = df_raw.iloc[start_row:].copy()
                            
                            df_dados['Codigo_Siafi'] = df_dados.iloc[:, col_conta].apply(limpar_codigo_bruto)
                            df_dados['Descricao_Planilha'] = df_dados.iloc[:, col_desc].astype(str).str.strip().str.upper()
                            
                            # O Valor fica na C(2) da Bruta ou D(3) da Fase 1
                            def extract_value(row):
                                col_idx = 3 if col_conta == 1 else 2
                                try:
                                    if col_idx < len(row):
                                        val = limpar_valor(row.iloc[col_idx])
                                        if abs(val) > 0.0: return val
                                except: pass
                                return limpar_valor(row.iloc[-1]) # Tenta a última coluna
                                
                            df_dados['Valor_Siafi'] = df_dados.apply(extract_value, axis=1)
                            
                            # Exclusões base (Garante não apagar os 07, 16 etc se for consumo)
                            exclusion_list = ['123110703', '123110402', '123119910']
                            df_dados = df_dados[~df_dados['Codigo_Siafi'].isin(exclusion_list)].copy()
                            
                            # Estoque Interno 2042
                            mask_2042 = df_dados['Codigo_Siafi'] == '2042'
                            if mask_2042.any():
                                saldo_2042 = df_dados.loc[mask_2042, 'Valor_Siafi'].sum()
                                if abs(saldo_2042) > 0.00: tem_2042_com_saldo = True
                            
                            # APLICA A TRADUÇÃO MATRIZ -> CHAVE
                            df_dados['Chave_Vinculo'] = df_dados['Codigo_Siafi'].apply(lambda c: extrair_chave_universal(c, lookup_dict))
                            
                            df_dados_filtrados = df_dados.dropna(subset=['Chave_Vinculo']).copy()
                            
                            if not df_dados_filtrados.empty:
                                df_dados_filtrados['Chave_Vinculo'] = df_dados_filtrados['Chave_Vinculo'].astype(int)
                                
                                df_padrao = df_dados_filtrados.groupby('Chave_Vinculo').agg({
                                    'Valor_Siafi': 'sum',
                                    'Descricao_Planilha': 'first'
                                }).reset_index()
                                
                                df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']

                    except Exception as e:
                        logs.append(f"❌ Erro no SIAFI UG {ug}: {e}")

                    # ==========================================
                    # LEITURA DO PDF (Com Trava contra números de linha)
                    # ==========================================
                    df_pdf_final = pd.DataFrame()
                    dados_pdf = []
                    
                    try:
                        par['pdf'].seek(0)
                        pdf_bytes = par['pdf'].read()
                        
                        with pdfplumber.open(io.BytesIO(pdf_bytes)) as p_doc:
                            for page in p_doc.pages:
                                txt = page.extract_text()
                                is_ocr = False
                                
                                if not txt or len(txt) < 50:
                                    is_ocr = True
                                    try:
                                        imagens = convert_from_bytes(pdf_bytes, first_page=page.page_number, last_page=page.page_number, dpi=300)
                                        if imagens:
                                            txt = pytesseract.image_to_string(imagens[0], lang='por', config='--psm 6')
                                    except: pass

                                if not txt: continue
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
                                            chave_final = None
                                            # Busca o código completo da conta no meio da frase para não ser enganado pelo OCR
                                            match_longo = re.search(r'\b((?:449|339)\d{5,})\b', line)
                                            
                                            if match_longo:
                                                chave_final = int(match_longo.group(1)[-2:])
                                            else:
                                                chave_match = re.match(r'^"?(\d+)', line)
                                                if chave_match:
                                                    c_raw = chave_match.group(1)
                                                    chave_final = int(c_raw[-2:]) if len(c_raw) >= 4 else int(c_raw)
                                                    
                                            if chave_final is not None:
                                                dados_pdf.append({
                                                    'Chave_Vinculo': chave_final,
                                                    'Saldo_PDF': limpar_valor(vals[-4])
                                                })
                        if dados_pdf:
                            df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
                    except Exception as e: logs.append(f"❌ Erro Leitura PDF UG {ug}: {e}")

                    # ==========================================
                    # CRUZAMENTO
                    # ==========================================
                    if df_padrao.empty: df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                    if df_pdf_final.empty: df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])

                    final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                    final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if pd.notna(x['Descricao_Completa']) and str(x['Descricao_Completa']).strip() != '0' else "ITEM SEM DESCRIÇÃO", axis=1)
                    final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                    divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                    soma_pdf = final['Saldo_PDF'].sum()
                    soma_excel = final['Saldo_Excel'].sum()
                    dif_total = soma_pdf - soma_excel

                    # --- EXIBIÇÃO ---
                    with st.expander("🛠️ Raio-X da Extração (Log de Auditoria)"):
                        st.write(f"**EXCEL:** Padrão detectado - Código na Coluna `{col_conta}`, Descrição na `{col_desc}`.")
                        st.write(f"**EXCEL:** Contas válidas extraídas: `{len(df_padrao)}`")
                        st.write(f"**PDF:** Contas válidas extraídas: `{len(df_pdf_final)}`")

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

                    # ==========================================
                    # RELATÓRIO PDF
                    # ==========================================
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
