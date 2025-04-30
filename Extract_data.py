import os
import re
import tempfile
import shutil
from io import BytesIO
from collections import defaultdict
import pandas as pd
import streamlit as st
import pdfplumber
from fpdf import FPDF
from datetime import datetime
from PIL import Image
from tempfile import NamedTemporaryFile
from PyPDF2 import PdfReader

# =================== CONFIGURAÇÃO ===================
st.set_page_config(page_title="CREA-RJ", layout="wide", page_icon="")

# =================== FUNÇÕES AUXILIARES ===================
def criar_temp_dir():
    """Cria diretório temporário"""
    return tempfile.mkdtemp()

def limpar_temp_dir(temp_dir):
    """Remove diretório temporário"""
    shutil.rmtree(temp_dir, ignore_errors=True)

def ajustar_texto_para_pdf(texto, largura_max, fonte=8):
    """Ajusta texto para caber na coluna do PDF"""
    if not texto:
        return texto
    tamanho = len(texto) * fonte * 0.5
    return texto if tamanho <= largura_max else texto[:int(largura_max/(fonte*0.5))-3] + '...'

# =================== MÓDULO DE EXTRAÇÃO ===================
def extrair_dados_vinculos_si(texto, filename):
    """Extrai dados para Vínculos e S.I"""
    dados = {'Arquivo': filename}
    
    # Extrai coordenadas
    coord = re.search(r"Latitude\s*:\s*([-\d,.]+).*?Longitude\s*:\s*([-\d,.]+)", texto)
    if coord:
        try:
            dados.update({
                'Latitude': float(coord.group(1).replace(',', '.')),
                'Longitude': float(coord.group(2).replace(',', '.'))
            })
        except ValueError:
            dados.update({'Latitude': None, 'Longitude': None})
    
    # Extrai endereço
    endereco = re.search(r'01\s*-\s*Endereço\s*Empreendimento\s*(.*?)(?=\s*(?:02\s*-|Descritivo:|$))', texto, re.DOTALL)
    dados['Endereço'] = re.sub(r'\s+', ' ', endereco.group(1).strip().replace('\n', ' ')) if endereco else None
    
    # Conta vínculos
    secao = re.search(r'04\s*-\s*Identificação.*?(?=05\s*-|$)', texto, re.DOTALL)
    if secao:
        secao_texto = secao.group()
        dados.update({
            'CONTRATADO': len(re.findall(r'CONTRATADO\s*:', secao_texto)),
            'RESPONSAVEL TECNICO': len(re.findall(r'RESPONSAVEL\s*TECNICO\s*:', secao_texto))
        })
        dados['Vínculos'] = dados['CONTRATADO'] + dados['RESPONSAVEL TECNICO']
    
    # Extrai ofícios GFIS
    oficios = re.findall(r'(?:OF[IÍ]CIO\s*[Nnº°]*\s*[.:-]*\s*)?(\d+.*?GFIS.*?)(?:\n|$)', texto, re.IGNORECASE)
    dados.update({
        'Ofícios GFIS': '; '.join(o.strip() for o in oficios) if oficios else '',
        'S.I': len(oficios) if oficios else 0
    })
    
    return dados

def extrair_dados_ramo_atividade(texto, filename):
    """Extrai dados para Ramo de Atividade"""
    dados = {
        'Arquivo': filename,
        'Ramo': '',
        'Qtd. Ramo': ''
    }
    
    secao = re.search(r'04\s*-\s*Identificação.*?(?=05\s*-|$)', texto, re.DOTALL|re.IGNORECASE)
    if secao:
        ramos = re.findall(r'Ramo\s*Atividade\s*:\s*(.*?)(?=\n|$)', secao.group(), re.IGNORECASE)
        if ramos:
            contagem = defaultdict(int)
            for ramo in [r.strip() for r in ramos if r.strip()]:
                contagem[ramo] += 1
            
            dados['Ramo'] = ", ".join(contagem.keys())
            dados['Qtd. Ramo'] = ", ".join(map(str, contagem.values()))
    
    return dados

def extrair_dados_processo_protocolo(texto, filename):
    """Extrai dados para Processo/Protocolo com foco em Legalização (anotações após OUTROS)"""
    dados = {
        'Arquivo': filename,
        'Fiscal': '',
        'Protocolo': '',
        'Legalização': '',
        'Qtd. Protocolo': 0,
        'Qtd. Legalização': 0
    }
    
    # Extrai fiscal
    fiscal = re.search(r'Agente\s+de\s+Fiscalização\s*:\s*([^\n]+)', texto)
    if fiscal:
        nome = re.sub(r'^\d+\s*-\s*', '', fiscal.group(1).strip())
        dados['Fiscal'] = nome.split()[0].capitalize() if nome.split() else nome
    
    # Extrai protocolo
    protocolo = re.search(r'Fato\s+Gerador\s*:\s*.*?(PROCESSO/PROTOCOLO[\s\-]*\d+[/\-]?\d*)', texto, re.IGNORECASE)
    if protocolo:
        nums = re.findall(r'\d+', protocolo.group(1))
        dados['Protocolo'] = ''.join(nums) if nums else ''
        dados['Qtd. Protocolo'] = 1 if nums else 0
    
    # Extrai conteúdo após "OUTROS" na seção de documentos recebidos
    secao_docs = re.search(r'(?i)06\s*-\s*Documentos\s*Recebidos(.*?)(?=\n\s*(?:07|$))', texto, re.DOTALL)
    if secao_docs:
        outros_match = re.search(r'(?i)OUTROS\s*[:\-]\s*(.*?)(?=\n|$)', secao_docs.group(1))
        if outros_match:
            outros_texto = outros_match.group(1).strip()
            # Remove múltiplos espaços e limpa o texto
            outros_texto = ' '.join(outros_texto.split())
            if outros_texto:
                dados['Legalização'] = outros_texto
                dados['Qtd. Legalização'] = 1
    
    return dados

# =================== GERADORES DE RELATÓRIO PDF ===================
def gerar_relatorio_vinculos_si(df):
    """Gera PDF para Vínculos e S.I"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, 'Relatório Vínculos e S.I', 0, 1, 'C')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}', 0, 1)
    pdf.ln(10)
    
    col_widths = [80, 30, 30]
    headers = ['Arquivo', 'Vínculos', 'S.I']
    
    pdf.set_font('Arial', 'B', 12)
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, 1, 0, 'C')
    pdf.ln()
    
    pdf.set_font('Arial', '', 10)
    for _, row in df[df['Arquivo'] != 'TOTAL'].iterrows():
        pdf.cell(col_widths[0], 10, str(row['Arquivo']), 1)
        pdf.cell(col_widths[1], 10, str(row['Vínculos']), 1, 0, 'C')
        pdf.cell(col_widths[2], 10, str(row['S.I']), 1, 0, 'C')
        pdf.ln()
    
    total = df[df['Arquivo'] == 'TOTAL'].iloc[0]
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(col_widths[0], 10, 'TOTAL GERAL', 1)
    pdf.cell(col_widths[1], 10, str(total['Vínculos']), 1, 0, 'C')
    pdf.cell(col_widths[2], 10, str(total['S.I']), 1, 0, 'C')
    
    return pdf.output(dest='S').encode('latin1')

def gerar_relatorio_ramo_atividade(df):
    """Gera PDF para Ramo de Atividade"""
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "RELATÓRIO DE RAMOS DE ATIVIDADE", 0, 1, 'C')
    pdf.ln(10)
    
    # Processa dados
    contagem = defaultdict(int)
    for _, row in df[df['Arquivo'] != 'TOTAL GERAL'].iterrows():
        if row['Ramo'] and row['Qtd. Ramo']:
            for ramo, qtd in zip(row['Ramo'].split(','), row['Qtd. Ramo'].split(',')):
                if ramo.strip() and qtd.strip().isdigit():
                    contagem[ramo.strip()] += int(qtd.strip())
    
    # Ordena por quantidade
    ramos_ordenados = sorted(contagem.items(), key=lambda x: x[1], reverse=True)
    total = sum(contagem.values())
    
    # Cabeçalho
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "RAMO DE ATIVIDADE", 1, 0, 'C')
    pdf.cell(30, 8, "QUANTIDADE", 1, 0, 'C')
    pdf.cell(30, 8, "PORCENTAGEM", 1, 1, 'C')
    
    # Dados
    pdf.set_font("Arial", size=9)
    for ramo, qtd in ramos_ordenados:
        pdf.cell(120, 8, ramo[:60] + ('...' if len(ramo) > 60 else ''), 1)
        pdf.cell(30, 8, str(qtd), 1, 0, 'C')
        pdf.cell(30, 8, f"{(qtd/total)*100:.1f}%" if total > 0 else "0%", 1, 1, 'C')
    
    # Rodapé
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "TOTAL GERAL", 1)
    pdf.cell(30, 8, str(total), 1, 0, 'C')
    pdf.cell(30, 8, "100%", 1, 1, 'C')
    
    pdf.ln(10)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 0, 'C')
    
    return pdf.output(dest='S').encode('latin1')

def gerar_relatorio_processo_protocolo(df):
    """Gera PDF para Processo/Protocolo com foco em Legalização"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "RELATÓRIO DE PROCESSOS/PROTOCOLOS", 0, 1, 'C')
    pdf.ln(8)
    
    # Configuração de colunas
    widths = [35, 30, 40, 40, 20, 20]
    headers = ['Arquivo', 'Fiscal', 'Protocolo', 'Legalização', 'Protocolo', 'Legalização']
    
    # Cabeçalho
    pdf.set_font("Arial", 'B', 8)
    for w, h in zip(widths, headers):
        pdf.cell(w, 8, h, 1, 0, 'C')
    pdf.ln()
    
    # Dados
    pdf.set_font("Arial", size=8)
    for _, row in df[df['Arquivo'] != 'TOTAL GERAL'].iterrows():
        # Calcula número de linhas necessárias
        lines_arquivo = len(str(row['Arquivo'])) // 30 + 1
        lines_legalizacao = len(str(row['Legalização'])) // 25 + 1 if row['Legalização'] else 1
        lines = max(lines_arquivo, lines_legalizacao, 1)
        line_height = 5
        
        # Arquivo
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.multi_cell(widths[0], line_height, str(row['Arquivo']), border=1, align='L')
        pdf.set_xy(x + widths[0], y)
        
        # Fiscal
        pdf.cell(widths[1], line_height * lines, str(row['Fiscal']), border=1, align='C')
        
        # Protocolo
        pdf.cell(widths[2], line_height * lines, str(row['Protocolo']), border=1, align='C')
        
        # Legalização
        x_legal = pdf.get_x()
        y_legal = pdf.get_y()
        legalizacao_text = str(row['Legalização']) if row['Legalização'] else ''
        pdf.multi_cell(widths[3], line_height, legalizacao_text, border=1, align='L')
        pdf.set_xy(x_legal + widths[3], y_legal)
        
        # Contagens
        pdf.cell(widths[4], line_height * lines, str(row['Qtd. Protocolo']), border=1, align='C')
        pdf.cell(widths[5], line_height * lines, str(row['Qtd. Legalização']), border=1, align='C')
        
        pdf.ln(line_height * lines)
        
        if pdf.get_y() > 270 - (line_height * lines):
            pdf.add_page()
            pdf.set_font("Arial", 'B', 8)
            for w, h in zip(widths, headers):
                pdf.cell(w, 8, h, 1, 0, 'C')
            pdf.ln()
            pdf.set_font("Arial", size=8)
    
    # Total
    if 'TOTAL GERAL' in df['Arquivo'].values:
        total = df[df['Arquivo'] == 'TOTAL GERAL'].iloc[0]
        pdf.set_font("Arial", 'B', 8)
        pdf.cell(sum(widths[:4]), line_height, 'TOTAL GERAL', 1, 0, 'R')
        pdf.cell(widths[4], line_height, str(total['Qtd. Protocolo']), 1, 0, 'C')
        pdf.cell(widths[5], line_height, str(total['Qtd. Legalização']), 1, 0, 'C')
    
    pdf.ln(12)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 0, 'C')
    
    return pdf.output(dest='S').encode('latin1')

# =================== MÓDULO PRINCIPAL ===================
def extrator_pdf_consolidado():
    st.title("📊 Extrator PDF Consolidado")
    st.markdown("""
    **Extrai automaticamente dados de:**
    - Vínculos e S.I
    - Ramos de Atividade  
    - Processos/Protocolos
    """)

    uploaded_files = st.file_uploader("Selecione os PDFs", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        temp_dir = criar_temp_dir()
        try:
            with st.spinner("Processando arquivos..."):
                dados_vs, dados_ra, dados_pp = [], [], []
                
                for file in uploaded_files:
                    temp_path = os.path.join(temp_dir, file.name)
                    with open(temp_path, "wb") as f:
                        f.write(file.getbuffer())
                    
                    with pdfplumber.open(temp_path) as pdf:
                        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
                    
                    dados_vs.append(extrair_dados_vinculos_si(texto, file.name))
                    dados_ra.append(extrair_dados_ramo_atividade(texto, file.name))
                    dados_pp.append(extrair_dados_processo_protocolo(texto, file.name))
                    
                    os.unlink(temp_path)
                
                # Cria DataFrames
                df_vs = pd.DataFrame(dados_vs)
                df_ra = pd.DataFrame(dados_ra)
                df_pp = pd.DataFrame(dados_pp)
                
                # Adiciona totais
                df_vs.loc['TOTAL'] = {
                    'Arquivo': 'TOTAL',
                    'Vínculos': df_vs['Vínculos'].sum(),
                    'S.I': df_vs['S.I'].sum()
                }
                
                total_ra = sum(int(q) for r in dados_ra for q in r['Qtd. Ramo'].split(',') if r['Qtd. Ramo'] and q.strip().isdigit())
                df_ra = pd.concat([df_ra, pd.DataFrame({
                    'Arquivo': ['TOTAL GERAL'],
                    'Ramo': [''],
                    'Qtd. Ramo': [str(total_ra)]
                })], ignore_index=True)
                
                # Total para Processo/Protocolo
                total_pp = pd.DataFrame({
                    'Arquivo': ['TOTAL GERAL'],
                    'Qtd. Protocolo': [df_pp['Qtd. Protocolo'].sum()],
                    'Qtd. Legalização': [df_pp['Qtd. Legalização'].sum()]
                })
                df_pp = pd.concat([df_pp, total_pp], ignore_index=True)
                
                # Exibição
                tab1, tab2, tab3 = st.tabs(["Vínculos e S.I", "Ramo Atividade", "Processo/Protocolo"])
                
                with tab1:
                    st.dataframe(df_vs)
                
                with tab2:
                    st.dataframe(df_ra)
                
                with tab3:
                    st.dataframe(df_pp)
                
                # Geração de relatórios
                pdf_vs = gerar_relatorio_vinculos_si(df_vs)
                pdf_ra = gerar_relatorio_ramo_atividade(df_ra)
                pdf_pp = gerar_relatorio_processo_protocolo(df_pp)
                
                # Excel consolidado
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_vs.to_excel(writer, sheet_name='Vínculos e S.I', index=False)
                    df_ra.to_excel(writer, sheet_name='Ramo Atividade', index=False)
                    df_pp.to_excel(writer, sheet_name='Processo Protocolo', index=False)
                
                # Download
                st.success("Processamento concluído!")
                st.download_button(
                    "⬇️ Baixar Excel Completo",
                    excel_buffer.getvalue(),
                    "dados_consolidados.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                with st.expander("⬇️ Baixar Relatórios Individuais (PDF)"):
                    cols = st.columns(3)
                    cols[0].download_button(
                        "Vínculos e S.I",
                        pdf_vs,
                        "relatorio_vinculos_si.pdf"
                    )
                    cols[1].download_button(
                        "Ramo Atividade",
                        pdf_ra,
                        "relatorio_ramos.pdf"
                    )
                    cols[2].download_button(
                        "Processo/Protocolo",
                        pdf_pp,
                        "relatorio_processos.pdf"
                    )
        
        finally:
            limpar_temp_dir(temp_dir)

# =================== INTERFACE PRINCIPAL ===================
def main():
    # Configuração visual
    try:
        logo = Image.open("10.png")
    except:
        logo = None
    
    # Layout do cabeçalho
    col1, col2 = st.columns([1, 2])
    with col1:
        if logo: st.image(logo, width=400)
    with col2:
        st.title("CREA-RJ - Conselho Regional de Engenharia e Agronomia do Rio de Janeiro")
    
    st.markdown("---")
    
    # Exibe apenas o módulo principal
    extrator_pdf_consolidado()

if __name__ == "__main__":
    main()