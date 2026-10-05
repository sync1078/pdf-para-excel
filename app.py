import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

def parse_pdf_text(text):
    data = []
    
    current_file = None
    current_client = None
    current_site = None
    buffer = []
    
    # Expressões Regulares para encontrar os blocos
    file_re = re.compile(r'^(\d{6})-?$')
    client_re = re.compile(r'^(.*?)\s*\(((?:SITE|CONSUMIDOR)[^)]*)\)$', re.IGNORECASE)
    combined_re = re.compile(r'^(\d{6})\s*-?\s*(.*?)\s*\(((?:SITE|CONSUMIDOR)[^)]*)\)$', re.IGNORECASE)
    date_re = re.compile(r'^\d{2}/\d{2}/\d{2}$')
    
    # Regex para capturar linhas de serviço que não tenham o delimitador '|'
    num_fmt = r'[\d.]*,\d{2}'
    pct_fmt = r'[\d.]*,\d{2}%'
    pattern = rf'^(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*(.*?)\s+({num_fmt})\s+({num_fmt})\s+(.*?)\s+({pct_fmt})\s+({num_fmt})$'

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    for line in lines:
        detail = None
        
        # O parser tenta encontrar linhas separadas por pipe '|' (padrão comum em tabelas) ou por espaços
        if '|' in line:
            parts = [p.strip() for p in line.split('|')]
            if len(parts) >= 11 and ('%' in parts[9] or parts[8] != ''):
                if re.match(num_fmt, parts[6]) and re.match(num_fmt, parts[7]):
                    detail = parts[1:11]
        else:
            m = re.match(pattern, line)
            if m:
                detail = list(m.groups())
                
        # Se encontrou uma linha de serviço (seja em texto corrido ou tabela)
        if detail:
            adt, chd, inf, qtd, voucher, tarifa, venda, categoria, pct, fee = detail
            
            data_servico = None
            servico = None
            
            # Buscar no histórico (buffer) a data e o nome do serviço imediatamente acima
            for b in reversed(buffer):
                if date_re.match(b):
                    data_servico = b
                elif not file_re.match(b) and not client_re.match(b) and not combined_re.match(b) and '|' not in b:
                    # Verifica se não é uma linha de totais 
                    if not re.match(r'^\d+\s+\d+\s+\d+\s+\d+', b):
                        if servico is None:
                            servico = b
            
            # Limpeza de números no padrão brasileiro (1.000,00 -> 1000.00)
            def clean_number(val):
                if not val: return 0.0
                return float(val.replace('.', '').replace(',', '.'))
            
            data.append({
                'FILE': current_file,
                'NOME_CLIENTE': current_client,
                'SITE_ORIGEM': current_site,
                'DATA_SERVICO': data_servico,
                'SERVICO': servico,
                'CATEGORIA_SERVICO': categoria,
                'ADT': float(adt) if adt else 0.0,
                'CHD': float(chd) if chd else 0.0,
                'INF': float(inf) if inf else 0.0,
                'QTD': float(qtd) if qtd else 0.0,
                'VOUCHER_RECIBO': voucher,
                'TARIFA': clean_number(tarifa),
                'VALOR_VENDA': clean_number(venda),
                'VALOR_FEE': clean_number(fee)
            })
            buffer = [] # Limpa o histórico após processar a linha
            continue
            
        # Atualizar quem é o Cliente Atual sendo lido
        m_comb = combined_re.match(line)
        if m_comb:
            current_file = m_comb.group(1)
            current_client = m_comb.group(2)
            current_site = m_comb.group(3)
        else:
            m_file = file_re.match(line)
            if m_file: current_file = m_file.group(1)
            
            m_client = client_re.match(line)
            if m_client:
                current_client = m_client.group(1)
                current_site = m_client.group(2)
        
        buffer.append(line)
        
    return pd.DataFrame(data)

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Comissionada')
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Conversor PDF para Excel", page_icon="📄")

st.title("📄 Conversor de PDF para Excel")
st.subheader("Extrator de Manutenção Comissionada")

uploaded_file = st.file_uploader("Selecione o relatório em formato PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Analisando o PDF e extraindo os dados..."):
        text = ""
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text(layout=False)
                    if page_text:
                        text += page_text + "\n"
                        
            df = parse_pdf_text(text)
            
            if not df.empty:
                st.success(f"Extração concluída! {len(df)} serviços encontrados.")
                st.dataframe(df.head(15)) # Mostra uma prévia das primeiras 15 linhas
                
                excel_data = to_excel(df)
                
                st.download_button(
                    label="📥 Baixar arquivo Excel (.xlsx)",
                    data=excel_data,
                    file_name="MANUTENCAO_COMISSIONADA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Não foi possível encontrar dados de serviços no formato esperado.")
        
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
