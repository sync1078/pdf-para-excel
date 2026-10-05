import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="Conversor PDF para Excel", layout="wide")

def extract_client_info(text):
    """Extrai FILE, NOME DO CLIENTE e SITE DE ORIGEM a partir da linha de cabeçalho do cliente."""
    text_str = str(text)
    file_id, nome, site = "", "", ""
    
    # Extrai ID (ex: 743133)
    file_match = re.search(r'\b(\d{5,})\b', text_str)
    if file_match:
        file_id = file_match.group(1)
        
    # Extrai Site Origem dentro dos parênteses
    site_match = re.search(r'\((.*?)\)', text_str)
    if site_match:
        site = site_match.group(1)
        
    # Limpa o texto restante para obter apenas o nome
    nome = text_str
    if file_id: nome = nome.replace(file_id, "")
    if site_match: nome = nome.replace(f"({site})", "")
    nome = re.sub(r'[-\n|]', ' ', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    
    return file_id, nome, site

def extract_service_info(text):
    """Extrai DATA DO SERVIÇO e NOME DO SERVIÇO a partir da célula principal."""
    text_str = str(text)
    date_match = re.search(r'(\d{2}/\d{2}/\d{2})', text_str)
    data_servico = date_match.group(1) if date_match else ""
    
    servico = text_str
    if data_servico:
        servico = servico.replace(data_servico, "")
    servico = re.sub(r'[-\n|]', ' ', servico)
    servico = re.sub(r'\s+', ' ', servico).strip()
    
    return data_servico, servico

def process_pdf(pdf_file):
    data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Estratégia focada em alinhamento de texto (ótimo para relatórios de sistema)
            table_settings = {
                "vertical_strategy": "text", 
                "horizontal_strategy": "text"
            }
            table = page.extract_table(table_settings)
            
            if not table:
                continue
                
            current_file = ""
            current_nome = ""
            current_site = ""
            
            for row in table:
                # Limpa nulos e junta a linha para avaliação
                row_clean = [str(cell).strip() if cell else "" for cell in row]
                full_row = " ".join(row_clean)
                
                # Pular cabeçalhos do PDF
                if "SERVIÇO" in full_row and "DATA" in full_row:
                    continue
                if "FEE" in full_row and "Lista Serviços" in full_row:
                    continue
                    
                # Identifica Linha do Cliente
                if re.search(r'\b\d{5,}\b', full_row) and '(' in full_row and ')' in full_row:
                    current_file, current_nome, current_site = extract_client_info(full_row)
                    continue
                    
                # Identifica Linha do Serviço (onde a data está presente)
                if re.search(r'\d{2}/\d{2}/\d{2}', full_row):
                    data_servico, servico = extract_service_info(row_clean[0])
                    
                    # Garantindo leitura mesmo se algumas colunas do PDF vierem mescladas
                    # Lendo da esquerda para a direita (quantidades)
                    adt = row_clean[1] if len(row_clean) > 1 else ""
                    chd = row_clean[2] if len(row_clean) > 2 else ""
                    inf = row_clean[3] if len(row_clean) > 3 else ""
                    qtd = row_clean[4] if len(row_clean) > 4 else ""
                    
                    voucher = row_clean[5] if len(row_clean) > 5 and not re.match(r'^[\d,.]+$', row_clean[5]) else ""
                    
                    # Lendo da direita para a esquerda (valores são mais consistentes no final da tabela)
                    fee = row_clean[-1] if len(row_clean) > 0 else ""
                    categoria = row_clean[-3] if len(row_clean) > 2 else ""
                    venda = row_clean[-4] if len(row_clean) > 3 else ""
                    tarifa = row_clean[-5] if len(row_clean) > 4 else ""
                    
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
                    
    # Cria o DataFrame com as colunas na exata ordem exigida pelo seu Excel
    df = pd.DataFrame(data, columns=[
        'FILE', 'NOME_CLIENTE', 'SITE_ORIGEM', 'DATA_SERVICO', 'SERVICO', 
        'CATEGORIA_SERVICO', 'ADT', 'CHD', 'INF', 'QTD', 'VOUCHER_RECIBO', 
        'TARIFA', 'VALOR_VENDA', 'VALOR_FEE'
    ])
    
    # Conversão e limpeza de colunas numéricas (substituindo vírgula por ponto)
    cols_numericas = ['ADT', 'CHD', 'INF', 'QTD', 'TARIFA', 'VALOR_VENDA', 'VALOR_FEE']
    for col in cols_numericas:
        if col in df.columns:
            # Remove qualquer caractere que não seja dígito, vírgula, ponto ou sinal de menos
            df[col] = df[col].astype(str).str.replace(r'[^\d,.-]', '', regex=True)
            df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    return df

# -- INTERFACE STREAMLIT --
st.title("📄 Conversor de PDF para Excel")
st.markdown("Transforme o relatório PDF de **Manutenção Comissionada** no formato Excel padronizado instantaneamente.")

uploaded_file = st.file_uploader("Selecione o arquivo PDF", type="pdf")

if uploaded_file is not None:
    with st.spinner("Processando o arquivo, isso pode levar alguns segundos..."):
        try:
            df_final = process_pdf(uploaded_file)
            
            st.success("Arquivo processado com sucesso!")
            
            # Mostra uma prévia na tela
            st.write("### Prévia dos Dados:")
            st.dataframe(df_final.head(10))
            
            # Botão de download (Convertendo para BytesIO)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Dados')
            processed_data = output.getvalue()
            
            st.download_button(
                label="⬇️ Baixar Planilha Excel",
                data=processed_data,
                file_name=uploaded_file.name.replace('.pdf', '.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
