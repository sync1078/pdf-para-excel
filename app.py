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
    
    # Expressões Regulares
    file_re = re.compile(r'^(\d{6})-?$')
    client_re = re.compile(r'^(.*?)\s*\(((?:SITE|CONSUMIDOR)[^)]*)\)$', re.IGNORECASE)
    combined_re = re.compile(r'^(\d{6})\s*-?\s*(.*?)\s*\(((?:SITE|CONSUMIDOR)[^)]*)\)$', re.IGNORECASE)
    date_re = re.compile(r'^\d{2}/\d{2}/\d{2,4}$')
    
    # TRUQUE MÁGICO: O PDF quebra as colunas em várias linhas antes do pipe "|".
    # Aqui nós "colamos" essas quebras de volta para reconstruir a linha da tabela.
    text = text.replace('\n |', ' |').replace('\n|', ' |')
    
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    for line in lines:
        detail = None
        
        # Tenta fatiar a linha pelos pipes '|'
        if '|' in line:
            parts = [p.strip() for p in line.split('|')]
            # Se a linha tem várias colunas e a coluna ADT (índice 1) é um número, é um serviço!
            if len(parts) >= 9 and parts[1].isdigit():
                try:
                    adt = parts[1]
                    chd = parts[2]
                    inf = parts[3]
                    qtd = parts[4]
                    voucher = parts[5]
                    tarifa = parts[6]
                    venda = parts[7]
                    categoria = parts[8]
                    fee = parts[-1] # O Fee é sempre o último item da linha
                    
                    detail = (adt, chd, inf, qtd, voucher, tarifa, venda, categoria, fee)
                except IndexError:
                    pass
                    
        if detail:
            adt, chd, inf, qtd, voucher, tarifa, venda, categoria, fee = detail
            
            data_servico = None
            servico = None
            
            # Buscar no histórico as linhas acima para achar a Data e o Nome do Serviço
            for b in reversed(buffer):
                if date_re.match(b):
                    data_servico = b
                elif not file_re.match(b) and not client_re.match(b) and not combined_re.match(b) and '|' not in b:
                    # Ignorar linhas de totais ou números avulsos
                    if not re.match(r'^\d+\s+\d+\s+\d+\s+\d+', b):
                        if servico is None:
                            servico = b
            
            # Limpeza financeira (ajusta R$ 1.000,00 para 1000.00 pro Excel entender)
            def clean_number(val):
                if not val: return 0.0
                val = re.sub(r'[^\d.,]', '', val)
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
            buffer = [] # Limpa a memória após gravar a linha
            continue
            
        # Atualizar quem é o Cliente atual
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
st.set_page_config(page_title="Conversor PDF para Excel", page_icon="📄", layout="wide")

st.title("📄 Conversor de PDF para Excel")
st.subheader("Extrator de Manutenção Comissionada")

uploaded_file = st.file_uploader("Selecione o relatório em formato PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Analisando o PDF e juntando as tabelas..."):
        text = ""
        try:
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    # Extração do texto base
                    page_text = page.extract_text(layout=False)
                    if page_text:
                        text += page_text + "\n"
                        
            df = parse_pdf_text(text)
            
            if not df.empty:
                st.success(f"Extração concluída com sucesso! {len(df)} serviços encontrados.")
                st.dataframe(df.head(15)) 
                
                excel_data = to_excel(df)
                
                st.download_button(
                    label="📥 Baixar arquivo Excel (.xlsx)",
                    data=excel_data,
                    file_name="MANUTENCAO_COMISSIONADA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Não foi possível extrair a tabela corretamente.")
                
                # MODO DEBUG: Expander para ver como o texto do PDF está saindo "cru"
                with st.expander("🛠️ Modo Debug - Ver texto bruto extraído do PDF"):
                    st.info("O texto abaixo é exatamente como o computador está lendo o seu PDF. Se estiver muito desconfigurado, precisamos ajustar as Regras (Regex) de leitura.")
                    st.text(text[:4000]) # Mostra um bom pedaço do texto extraído para diagnóstico
        
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
