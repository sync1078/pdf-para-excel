import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="Conversor PDF para Excel", layout="wide")

def is_numeric(s):
    """Verifica se a string é composta por números/valores financeiros."""
    return bool(re.match(r'^-?[\d.,]+$', s.strip()))

def parse_service_line(line_words):
    """
    Separa rigorosamente as palavras de uma linha de serviço entre a descrição do Serviço,
    a Categoria e as colunas numéricas (ADT, CHD, INF, QTD, TARIFA, VALOR_VENDA, VALOR_FEE).
    """
    adt = chd = inf = qtd = voucher = tarifa = venda = fee = ""
    servico_words = []
    categoria_words = []
    
    for w in line_words:
        text = w['text']
        x0 = w['x0']
        
        # Identificação de colunas por coordenadas X exatas do relatório
        if 275 <= x0 <= 298 and is_numeric(text):
            adt = text
        elif 305 <= x0 <= 328 and is_numeric(text):
            chd = text
        elif 335 <= x0 <= 358 and is_numeric(text):
            inf = text
        elif 365 <= x0 <= 388 and is_numeric(text):
            qtd = text
        elif 390 <= x0 < 470 and is_numeric(text):
            voucher = text
        elif 470 <= x0 < 528 and is_numeric(text):
            tarifa = text
        elif 528 <= x0 < 588 and is_numeric(text):
            venda = text
        elif 588 <= x0 < 720:
            categoria_words.append(text)
        elif 720 <= x0 < 770:
            pass  # Percentual do Fee (ex: 1,00 %) - ignorado
        elif 770 <= x0 <= 820 and is_numeric(text):
            fee = text
        elif 70 <= x0 < 588:
            # Texto da descrição do serviço
            servico_words.append(text)
            
    servico = " ".join(servico_words).strip()
    categoria = " ".join(categoria_words).strip()
    
    return servico, categoria, adt, chd, inf, qtd, voucher, tarifa, venda, fee

def extract_client_info(line_words):
    """Extrai ID (FILE), Nome do Cliente e Site da linha do cliente."""
    full_text = " ".join([w['text'] for w in line_words]).strip()
    file_id, site = "", ""
    
    m_file = re.search(r'^(\d{5,})', full_text)
    if m_file:
        file_id = m_file.group(1)
        full_text = full_text.replace(file_id, "", 1)
        
    sites = re.findall(r'\((.*?)\)', full_text)
    if sites:
        site = sites[-1]
        full_text = full_text.rsplit(f"({site})", 1)[0]
        
    nome = re.sub(r'^[\s-]*', '', full_text).strip()
    return file_id, nome, site

def process_pdf(pdf_file):
    data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # x_tolerance=1.5 evita a fusão de textos sobrepostos com números de colunas
            words = page.extract_words(x_tolerance=1.5, y_tolerance=3)
            
            # Agrupa palavras por altura (eixo Y)
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
                
                # Ignora cabeçalhos e rodapés da página
                if re.search(r'(DATA|FEE|ORIGEM|SERVIÇO|TOTAL|Página|EMISSÃO|INÍCIO|FIM)', primeira_palavra, re.IGNORECASE):
                    continue
                
                # 1. Linha do Cliente (começa com número do FILE ex: 743133)
                if re.match(r'^\d{5,}$', primeira_palavra) and inicio_x0 < 40:
                    current_file, current_nome, current_site = extract_client_info(line_sorted)
                    continue
                
                # 2. Linha do Serviço (começa com Data ex: 04/09/26)
                elif re.match(r'^\d{2}/\d{2}/\d{2}$', primeira_palavra) and inicio_x0 < 40:
                    data_servico = primeira_palavra
                    resto_palavras = [w for w in line_sorted if w['x0'] >= 70]
                    
                    servico, categoria, adt, chd, inf, qtd, voucher, tarifa, venda, fee = parse_service_line(resto_palavras)
                    
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
                    
                # 3. Linha de continuação do nome do serviço (quando quebra em 2 linhas)
                elif inicio_x0 >= 70 and inicio_x0 < 270 and len(data) > 0:
                    servico_extra, cat_extra, _, _, _, _, _, _, _, _ = parse_service_line(line_sorted)
                    if servico_extra:
                        data[-1]['SERVICO'] += " " + servico_extra
                    if cat_extra:
                        data[-1]['CATEGORIA_SERVICO'] += " " + cat_extra

    df = pd.DataFrame(data, columns=[
        'FILE', 'NOME_CLIENTE', 'SITE_ORIGEM', 'DATA_SERVICO', 'SERVICO', 
        'CATEGORIA_SERVICO', 'ADT', 'CHD', 'INF', 'QTD', 'VOUCHER_RECIBO', 
        'TARIFA', 'VALOR_VENDA', 'VALOR_FEE'
    ])
    
    # Tratamento e conversão de tipos numéricos
    cols_numericas = ['ADT', 'CHD', 'INF', 'QTD', 'TARIFA', 'VALOR_VENDA', 'VALOR_FEE']
    for col in cols_numericas:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({'': '0', 'nan': '0', 'None': '0'})
            df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.-]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
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
