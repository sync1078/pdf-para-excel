import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="Conversor PDF para Excel", layout="wide")

def is_numeric(s):
    """Verifica se a string é composta por números/valores financeiros."""
    return re.match(r'^-?[\d.,]+$', s.strip()) is not None

def parse_line_coords(words):
    """
    Classifica cada palavra lida do PDF com base na sua posição horizontal (Eixo X).
    Evita que textos longos de serviços se misturem com as colunas numéricas.
    """
    words_sorted = sorted(words, key=lambda w: w['x0'])
    
    adt = chd = inf = qtd = voucher = tarifa = venda = fee = ""
    categoria_words = []
    texto_principal_words = []
    
    for w in words_sorted:
        text = w['text']
        x0 = w['x0']
        
        if 270 <= x0 < 305 and is_numeric(text):
            adt = text
        elif 305 <= x0 < 335 and is_numeric(text):
            chd = text
        elif 335 <= x0 < 365 and is_numeric(text):
            inf = text
        elif 365 <= x0 < 390 and is_numeric(text):
            qtd = text
        elif 390 <= x0 < 470 and is_numeric(text):
            voucher = text
        elif 470 <= x0 < 530 and is_numeric(text):
            tarifa = text
        elif 530 <= x0 < 590 and is_numeric(text):
            venda = text
        elif 590 <= x0 < 720:
            categoria_words.append(text)
        elif 720 <= x0 < 780:
            pass  # Ignora porcentagem do FEE (ex: 1,00 %)
        elif 780 <= x0 and is_numeric(text):
            fee = text
        else:
            texto_principal_words.append(text)
            
    texto_principal = " ".join(texto_principal_words)
    categoria = " ".join(categoria_words)
    return texto_principal, categoria.strip(), adt, chd, inf, qtd, voucher, tarifa, venda, fee

def extract_client_info(texto_principal):
    """Extrai ID, Nome e Site da string do cliente."""
    file_id, site = "", ""
    
    m_file = re.search(r'^(\d{5,})', texto_principal)
    if m_file:
        file_id = m_file.group(1)
        texto_principal = texto_principal.replace(file_id, "", 1)
        
    sites = re.findall(r'\((.*?)\)', texto_principal)
    if sites:
        site = sites[-1]
        texto_principal = texto_principal.rsplit(f"({site})", 1)[0]
        
    nome = re.sub(r'^[\s-]*', '', texto_principal).strip()
    return file_id, nome, site

def extract_service_info(texto_principal):
    """Extrai Data e Nome do Serviço."""
    data_servico = ""
    m_data = re.search(r'^(\d{2}/\d{2}/\d{2})', texto_principal)
    if m_data:
        data_servico = m_data.group(1)
        texto_principal = texto_principal.replace(data_servico, "", 1)
        
    servico = re.sub(r'^[\s-]*', '', texto_principal).strip()
    return data_servico, servico

def process_pdf(pdf_file):
    data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            
            # Agrupa palavras na mesma linha vertical (mesmo eixo Y)
            words_by_top = sorted(words, key=lambda w: w['top'])
            lines = []
            current_line = []
            current_top = None
            
            for w in words_by_top:
                if current_top is None or abs(w['top'] - current_top) <= 3:
                    current_line.append(w)
                    if current_top is None: 
                        current_top = w['top']
                else:
                    lines.append(current_line)
                    current_line = [w]
                    current_top = w['top']
            if current_line:
                lines.append(current_line)
                
            current_file, current_nome, current_site = "", "", ""
            
            for line in lines:
                line_sorted = sorted(line, key=lambda w: w['x0'])
                if not line_sorted:
                    continue
                
                primeira_palavra = line_sorted[0]['text']
                inicio_x0 = line_sorted[0]['x0']
                
                # Descarta cabeçalhos e rodapés
                if re.search(r'(DATA|FEE|ORIGEM|SERVIÇO|TOTAL|Página|EMISSÃO|INÍCIO|FIM)', primeira_palavra, re.IGNORECASE):
                    continue
                
                # Linha de Cliente (Ex: 743133 - RODRIGO...)
                if re.match(r'^\d{5,}$', primeira_palavra) and inicio_x0 < 40:
                    texto_principal, _, _, _, _, _, _, _, _, _ = parse_line_coords(line_sorted)
                    current_file, current_nome, current_site = extract_client_info(texto_principal)
                    continue
                
                # Linha de Serviço (Ex: 04/09/26 Ticket Bustour...)
                elif re.match(r'^\d{2}/\d{2}/\d{2}$', primeira_palavra) and inicio_x0 < 40:
                    texto_principal, categoria, adt, chd, inf, qtd, voucher, tarifa, venda, fee = parse_line_coords(line_sorted)
                    data_servico, servico = extract_service_info(texto_principal)
                    
                    data.append({
                        'FILE': current_file,
                        'NOME_CLIENTE': current_nome,
                        'SITE_ORIGEM': current_site,
                        'DATA_SERVICO': data_servico,
                        'SERVICO': servico,
                        'CATEGORIA_SERVICO': categoria,
                        'ADT': adt,
                        'CHD': chd,
                        'INF': inf,
                        'QTD': qtd,
                        'VOUCHER_RECIBO': voucher,
                        'TARIFA': tarifa,
                        'VALOR_VENDA': venda,
                        'VALOR_FEE': fee
                    })
                    
                # Linha de continuação do nome do serviço
                elif inicio_x0 >= 70 and inicio_x0 < 270 and len(data) > 0:
                    texto_principal, categoria_extra, _, _, _, _, _, _, _, _ = parse_line_coords(line_sorted)
                    if texto_principal:
                        data[-1]['SERVICO'] += " " + texto_principal.strip()
                    if categoria_extra:
                        data[-1]['CATEGORIA_SERVICO'] += " " + categoria_extra.strip()

    df = pd.DataFrame(data, columns=[
        'FILE', 'NOME_CLIENTE', 'SITE_ORIGEM', 'DATA_SERVICO', 'SERVICO', 
        'CATEGORIA_SERVICO', 'ADT', 'CHD', 'INF', 'QTD', 'VOUCHER_RECIBO', 
        'TARIFA', 'VALOR_VENDA', 'VALOR_FEE'
    ])
    
    # Tratamento das colunas numéricas
    cols_numericas = ['ADT', 'CHD', 'INF', 'QTD', 'TARIFA', 'VALOR_VENDA', 'VALOR_FEE']
    for col in cols_numericas:
        if col in df.columns:
            df[col] = df[col].replace("", "0")
            df[col] = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.-]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    return df

# -- INTERFACE STREAMLIT --
st.title("📄 Conversor de PDF para Excel")
st.markdown("Transformação do relatório de **Manutenção Comissionada**.")

uploaded_file = st.file_uploader("Selecione o arquivo PDF", type="pdf")

if uploaded_file is not None:
    with st.spinner("Mapeando coordenadas e processando tabela..."):
        try:
            df_final = process_pdf(uploaded_file)
            st.success("Arquivo convertido com sucesso!")
            
            st.write("### Prévia dos Dados:")
            st.dataframe(df_final.head(15))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Dados')
            processed_data = output.getvalue()
            
            st.download_button(
                label="⬇️ Baixar Planilha Excel (.xlsx)",
                data=processed_data,
                file_name=uploaded_file.name.replace('.pdf', '_convertido.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar: {e}")
