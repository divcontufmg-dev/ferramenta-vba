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
    page_title="Conciliação Patrimonial: RMB x SIAFI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Ocultar marcas do Streamlit para um visual mais limpo
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE PROCESSAMENTO (BASTIDORES)
# ==========================================
def limpar_valor(v):
    if v is None or pd.isna(v) or str(v).strip() == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    v = str(v).replace('"', '').replace("'", "").strip()
    if re.search(r',\d{1,2}$', v): v = v.replace('.', '').replace(',', '.')
    elif re.search(r'\.\d{1,2}$', v): v = v.replace(',', '')
    try: return float(re.sub(r'[^\d.-]', '', v))
    except: return 0.0

def extract_excel_data(df_raw):
    extracted_data = []
    for idx, row in df_raw.iterrows():
        if row.isna().all(): continue
        val_0 = str(row.iloc[0]).strip().replace('.0', '')
        
        if val_0.startswith('123'):
            codigo = val_0
            desc = "SEM DESCRIÇÃO"
            val = 0.0
            cols = [c for c in row.iloc[1:] if pd.notna(c) and str(c).strip() != '']
            
            if len(cols) >= 2:
                desc = str(cols[0]).strip().upper()
                val = limpar_valor(cols[1])
            elif len(cols) == 1:
                parsed_val = limpar_valor(cols[0])
                if parsed_val != 0.0 or str(cols[0]).strip() in ['0', '0.0']:
                    val = parsed_val
                else:
                    desc = str(cols[0]).strip().upper()
                    
            extracted_data.append({'Conta': codigo, 'Descricao': desc, 'Valor': val})
    return pd.DataFrame(extracted_data)

def get_chave_vinculo(conta, dict_matriz):
    conta = str(conta).strip()
    if conta in dict_matriz:
        val_matriz = str(dict_matriz[conta])
        match = re.search(r'(\d+)$', val_matriz)
        if match:
            digits = match.group(1)
            return int(digits[-2:]) if len(digits) >= 2 else int(digits)
    return None

def formatar_real(valor):
    return f"{valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

class PDF_Report(FPDF):
    def header(self):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, 'Relatório de Conferência Patrimonial', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(5)
    def footer(self):
        self.set_y(-15); self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', align='C')

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================
st.title("📊 Sistema de Conciliação: RMB x SIAFI")
st.markdown("""
Bem-vindo! Esta ferramenta automatiza a conferência entre os saldos do SIAFI e os relatórios do RMB.
**Instruções:**
1. Arraste para a área abaixo a sua **Planilha SIAFI** e todos os **Relatórios PDF (RMB)** correspondentes.
2. Clique em "Gerar Relatório de Conciliação" e aguarde a análise.
""")
st.markdown("---")

# Área de Upload Unificada
arquivos_enviados = st.file_uploader(
    "📂 Arraste aqui a Planilha SIAFI (.xlsx) e os PDFs do RMB juntos", 
    accept_multiple_files=True, 
    type=['xlsx', 'pdf']
)

st.markdown("---")

# ==========================================
# EXECUÇÃO DO SISTEMA
# ==========================================
if st.button("🚀 Gerar Relatório de Conciliação", type="primary", use_container_width=True):
    
    # Validação inicial dos arquivos necessários
    if not os.path.exists("MATRIZ.xlsx"):
        st.error("❌ O arquivo de configuração interno ('MATRIZ.xlsx') não foi encontrado. Contate o suporte técnico.")
        st.stop()
        
    if not arquivos_enviados:
        st.warning("⚠️ Por favor, insira os arquivos para iniciar a conciliação.")
        st.stop()

    # Separa os arquivos enviados entre a planilha principal e os PDFs
    uploaded_siafi = None
    uploaded_pdfs = []
    
    for arquivo in arquivos_enviados:
        if arquivo.name.lower().endswith('.xlsx'):
            if uploaded_siafi is None:
                uploaded_siafi = arquivo
            else:
                st.info(f"ℹ️ O sistema utilizará a planilha '{uploaded_siafi.name}' como base principal.")
        elif arquivo.name.lower().endswith('.pdf'):
            uploaded_pdfs.append(arquivo)

    # Verificações de envio
    if not uploaded_siafi:
        st.warning("⚠️ A Planilha SIAFI (.xlsx) não foi encontrada entre os arquivos enviados.")
        st.stop()
    if not uploaded_pdfs:
        st.warning("⚠️ Nenhum relatório RMB (.pdf) foi encontrado entre os arquivos enviados.")
        st.stop()

    progresso = st.progress(0)
    status_text = st.empty()
    status_text.text("Preparando ambiente de conciliação...")

    # 1. Carregar a Matriz de Relacionamento (Transparente para o usuário)
    try:
        df_matriz = pd.read_excel("MATRIZ.xlsx", header=None)
        dict_matriz = {}
        for i in range(len(df_matriz)):
            c0 = str(df_matriz.iloc[i, 0]).strip().replace('.0', '')
            c1 = str(df_matriz.iloc[i, 1]).strip().replace('.0', '')
            if c0.startswith('123'): dict_matriz[c0] = c1
            elif c1.startswith('123'): dict_matriz[c1] = c0
    except Exception as e:
        st.error("❌ Ocorreu um erro ao ler a Matriz de configuração. Verifique o arquivo MATRIZ.xlsx.")
        st.stop()

    # 2. Parear as abas da Planilha com os PDFs correspondentes
    pdfs = {f.name: f for f in uploaded_pdfs}
    pares = []
    avisos_usuario = []

    try:
        xls_file = pd.ExcelFile(uploaded_siafi)
        for sheet_name in xls_file.sheet_names:
            if sheet_name.upper() == "MATRIZ": continue
            
            # Identifica o número da Unidade Gestora pelo nome da aba
            match = re.search(r'^(\d+)', sheet_name)
            if match:
                ug = match.group(1)
                pdf_match = next((f for n, f in pdfs.items() if n.startswith(ug)), None)
                if pdf_match: 
                    pares.append({'ug': ug, 'sheet_name': sheet_name, 'pdf': pdf_match})
                else: 
                    avisos_usuario.append(f"Falta PDF: A Unidade Gestora {ug} está na planilha, mas o PDF correspondente não foi enviado.")
    except Exception as e:
        st.error("❌ Não foi possível ler a Planilha SIAFI. Certifique-se de que o arquivo não está corrompido.")
        st.stop()

    if not pares:
        st.error("❌ Não foi possível encontrar pares correspondentes (Aba do Excel + PDF com o mesmo número de UG). Verifique o nome dos arquivos.")
    else:
        pdf_out = PDF_Report()
        pdf_out.add_page()
        st.subheader("🔍 Resultados da Conciliação")

        for idx, par in enumerate(pares):
            ug = par['ug']
            status_text.text(f"Analisando dados da Unidade Gestora: {ug}...")
            
            with st.container():
                st.info(f"🏢 **Unidade Gestora: {ug}**")
                
                # --- LEITURA DO EXCEL ---
                df_padrao = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa'])
                saldo_estoque = 0.0
                tem_estoque_com_saldo = False
                
                try:
                    df_raw = pd.read_excel(xls_file, sheet_name=par['sheet_name'], header=None)
                    df_dados = extract_excel_data(df_raw)
                    
                    if not df_dados.empty:
                        # Extrai saldo de Estoque Interno para informação adicional
                        if '123110801' in df_dados['Conta'].values:
                            saldo_estoque = df_dados[df_dados['Conta'] == '123110801']['Valor'].sum()
                            if abs(saldo_estoque) > 0.0: tem_estoque_com_saldo = True
                        
                        # Remove contas que não participam do cruzamento
                        contas_ignoradas = ['123110703', '123110402', '123119910', '123110801']
                        df_dados = df_dados[~df_dados['Conta'].isin(contas_ignoradas)].copy()
                        
                        df_dados['Chave_Vinculo'] = df_dados['Conta'].apply(lambda c: get_chave_vinculo(c, dict_matriz))
                        df_valid = df_dados.dropna(subset=['Chave_Vinculo']).copy()
                        
                        if not df_valid.empty:
                            df_valid['Chave_Vinculo'] = df_valid['Chave_Vinculo'].astype(int)
                            df_padrao = df_valid.groupby('Chave_Vinculo').agg({
                                'Valor': 'sum',
                                'Descricao': 'first'
                            }).reset_index()
                            df_padrao.columns = ['Chave_Vinculo', 'Saldo_Excel', 'Descricao_Completa']
                except Exception as e:
                    avisos_usuario.append(f"Erro ao processar os dados da planilha para a UG {ug}.")

                # --- LEITURA DO PDF ---
                df_pdf_final = pd.DataFrame(columns=['Chave_Vinculo', 'Saldo_PDF'])
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
                                        chave_match = re.match(r'^"?(\d+)', line)
                                        if chave_match:
                                            chave_raw = chave_match.group(1)
                                            chave_final = int(chave_raw[-2:]) if len(chave_raw) >= 4 else int(chave_raw)
                                            
                                            dados_pdf.append({
                                                'Chave_Vinculo': chave_final,
                                                'Saldo_PDF': limpar_valor(vals[-4])
                                            })
                    if dados_pdf:
                        df_pdf_final = pd.DataFrame(dados_pdf).groupby('Chave_Vinculo')['Saldo_PDF'].sum().reset_index()
                except Exception as e: 
                    avisos_usuario.append(f"Erro ao ler o documento PDF da UG {ug}.")

                # --- CRUZAMENTO DOS DADOS ---
                final = pd.merge(df_pdf_final, df_padrao, on='Chave_Vinculo', how='outer').fillna(0)
                final['Descricao'] = final.apply(lambda x: x['Descricao_Completa'] if pd.notna(x['Descricao_Completa']) and str(x['Descricao_Completa']).strip() != '0' else "ITEM SEM DESCRIÇÃO NO SIAFI", axis=1)
                final['Diferenca'] = (final['Saldo_PDF'] - final['Saldo_Excel']).round(2)
                divergencias = final[abs(final['Diferenca']) > 0.05].copy()

                soma_pdf = final['Saldo_PDF'].sum()
                soma_excel = final['Saldo_Excel'].sum()
                dif_total = soma_pdf - soma_excel

                # --- EXIBIÇÃO EM TELA ---
                col1, col2, col3 = st.columns(3)
                col1.metric("Total RMB (PDF)", f"R$ {soma_pdf:,.2f}")
                col2.metric("Total SIAFI (Excel)", f"R$ {soma_excel:,.2f}")
                col3.metric("Diferença Encontrada", f"R$ {dif_total:,.2f}", delta_color="inverse" if abs(dif_total) > 0.05 else "normal")
                
                if not divergencias.empty:
                    st.warning(f"Atenção: Foram encontradas {len(divergencias)} conta(s) com divergência de valores.")
                    with st.expander("Visualizar Contas com Divergência"):
                        st.dataframe(divergencias[['Chave_Vinculo', 'Descricao', 'Saldo_PDF', 'Saldo_Excel', 'Diferenca']])
                else: 
                    st.success("✅ Conciliado com sucesso! Nenhuma divergência de valores foi encontrada.")

                if tem_estoque_com_saldo: 
                    st.info(f"Aviso Contábil: A Conta de Estoque Interno (123110801) possui saldo de R$ {saldo_estoque:,.2f}.")
                st.markdown("---")

                # --- GERAÇÃO DO PDF FINAL ---
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
                    pdf_out.cell(0, 8, "Nenhuma divergência encontrada entre SIAFI e RMB.", 1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                if tem_estoque_com_saldo:
                    pdf_out.ln(2)
                    pdf_out.set_font("helvetica", 'B', 9)
                    pdf_out.set_fill_color(255, 255, 200)
                    pdf_out.cell(100, 8, "SALDO ESTOQUE INTERNO (123110801)", 1, fill=True)
                    pdf_out.cell(90, 8, f"R$ {formatar_real(saldo_estoque)}", 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                pdf_out.ln(2)
                pdf_out.set_font("helvetica", 'B', 9)
                pdf_out.set_fill_color(220, 230, 241)
                pdf_out.cell(100, 8, "RESUMO DOS TOTAIS", 1, fill=True)
                pdf_out.cell(30, 8, formatar_real(soma_pdf), 1, fill=True)
                pdf_out.cell(30, 8, formatar_real(soma_excel), 1, fill=True)
                if abs(dif_total) > 0.05: pdf_out.set_text_color(200, 0, 0)
                pdf_out.cell(30, 8, formatar_real(dif_total), 1, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf_out.set_text_color(0, 0, 0)
                pdf_out.ln(5)
            
            progresso.progress((idx + 1) / len(pares))

        status_text.text("Concluído! O relatório final está pronto para download.")
        progresso.empty()
        
        # Exibe os avisos apenas se houver algum
        if avisos_usuario:
            st.warning("⚠️ **Avisos do Sistema:**")
            for aviso in avisos_usuario:
                st.write(f"- {aviso}")
        
        # Botão final de Download
        try:
            pdf_bytes = bytes(pdf_out.output())
            st.download_button(
                label="📥 BAIXAR RELATÓRIO CONSOLIDADO (.PDF)", 
                data=pdf_bytes, 
                file_name="Relatorio_Conciliacao_Patrimonial.pdf", 
                mime="application/pdf", 
                type="primary", 
                use_container_width=True
            )
        except Exception as e: 
            st.error("Ocorreu um erro ao gerar o arquivo PDF para download.")
