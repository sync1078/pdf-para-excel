import streamlit as st
import pandas as pd
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
import re
import io

st.set_page_config(page_title="Conversor PDF para Excel", layout="wide")

def process_pdf(pdf_file):
    all_data = []
    
    # Processa página a página utilizando extração de containers nativos do PDF
    for page_layout in extract_pages(pdf_file):
        lines = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    txt = text_line.get_text().strip()
                    if txt:
                        lines.append((text_line.bbox[0], text_line.bbox[1], txt))
                        
        # Agrupa elementos pertencentes à mesma linha horizontal
        lines_grouped = {}
        for x0, y0, txt in lines:
            y_round = round(y0, 1)
            matched_y = None
            for y_key in lines_grouped:
                if abs(y_key - y_round) <= 2.5:
                    matched_y = y_key
                    break
            if matched_y is None:
                matched_y = y_round
                lines_grouped[matched_y] = []
            lines_grouped[matched_y].append((x0, txt))
            
        sorted_ys = sorted(lines_grouped.keys(), reverse=True)
        
        current_file = ""
        current_nome = ""
        current_site = ""
        
        for y in sorted_ys:
            row_items = sorted(lines_grouped[y], key=lambda item: item[0])
            first_x, first_txt = row_items[0]
            
            # Pula cabeçalhos e rodapés
            if re.search(r'(DATA SERVIÇO|FEE -|ORIGEM:|INÍCIO|FIM SERVIÇO|EMISSÃO:|SERVIÇO:|Página|TOTAL)', first_txt):
                continue
            
            # Identifica Linha do Cliente (ex: "743133 - RODRIGO ALMEIDA (SITE BROCKER)")
            m_client = re.match(r'^(\d{5,})\s*-\s*(.*?)\s*\((.*?)\)$', first_txt)
            if m_client and first_x < 50:
                current_file = m_client.group(1)
                current_nome = m_client.group(2).strip()
                current_site = m_client.group(3).strip()
                continue
            
            # Identifica Linha do Serviço (começa com data ex: "04/09/26")
            m_date = re.match(r'^\d{2}/\d{2}/\d{2}$', first_txt)
            if m_date and first_x < 50:
                data_servico = first_txt
                
                servico = ""
                adt = chd = inf = qtd = voucher = tarifa = venda = categoria = fee = ""
                
                for x0, txt in row_items[1:]:
                    if 70 <= x0 < 270:
                        servico = (servico + " " + txt).strip()
                    elif 270 <= x0 < 305:
                        adt = txt
                    elif 305 <= x0 < 335:
                        chd = txt
                    elif 335 <= x0 < 365:
                        inf = txt
                    elif 365 <= x0 < 390:
                        qtd = txt
                    elif 390 <= x0 < 470:
                        voucher = txt
                    elif 470 <= x0 < 528:
                        tarifa = txt
                    elif 528 <= x0 < 588:
                        venda = txt
                    elif 588 <= x0 < 720:
                        categoria = (categoria + " " + txt).strip()
                    elif 720 <= x0 < 780:
                        pass  # Ignora % do Fee
                    elif 780 <= x0:
                        fee = txt
                        
                all_data.append({
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
                
            # Trata linhas com nomes de serviços e categorias que quebram em duas linhas
            elif 70 <= first_x < 270 and len(all_data) > 0 and current_file != "":
                for x0, txt in row_items:
                    if 70 <= x0 < 270:
                        all_data[-1]['SERVICO'] = (all_data[-1]['SERVICO'] + " " + txt).strip()
                    elif 588 <= x0 < 720:
                        all_data[-1]['CATEGORIA_SERVICO'] = (all_data[-1]['CATEGORIA_SERVICO'] + " " + txt).strip()

    df = pd.DataFrame(all_data, columns=[
        'FILE', 'NOME_CLIENTE', 'SITE_ORIGEM', 'DATA_SERVICO', 'SERVICO', 
        'CATEGORIA_SERVICO', 'ADT', 'CHD', 'INF', 'QTD', 'VOUCHER_RECIBO', 
        'TARIFA', 'VALOR_VENDA', 'VALOR_FEE'
    ])
    
    # Tratamento de colunas numéricas
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
    with st.spinner("Lendo estrutura PDF de alta precisão..."):
        try:
            df_final = process_pdf(uploaded_file)
            st.success(f"Arquivo convertido com sucesso! Total de {len(df_final)} registros processados.")
            
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
