import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Conversor PDF para Excel - Brocker", layout="centered")

def clean_brazilian_currency(value):
    """Converte strings de moeda no formato brasileiro (1.000,00) para float (1000.00)."""
    if pd.isna(value) or value is None or str(value).strip() == "":
        return 0.0
    val_str = str(value).strip()
    
    # Se não houver números, retorna 0
    if not re.search(r'\d', val_str):
        return 0.0
        
    # Remove pontos de milhar e troca vírgula por ponto
    val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def process_pdf(pdf_file):
    all_rows = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            # Extrai a tabela da página
            table = page.extract_table()
            
            if table:
                for row in table:
                    # Remove quebras de linha indesejadas dentro das células
                    cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                    
                    # Transforma a linha em uma string para verificar lixos de paginação
                    row_text = " ".join(cleaned_row).lower()
                    
                    # Ignorar linhas de cabeçalho repetido, datas, URLs e paginação do PDF
                    if ("sumário" in row_text or 
                        "https://" in row_text or 
                        "receita operacional" in row_text or
                        "total geral" in row_text or
                        not any(char.isdigit() for char in row_text)): # Pula se não tiver nenhum número
                        continue
                    
                    # Filtra colunas totalmente vazias criadas por artefatos do PDF
                    cleaned_row = [cell for cell in cleaned_row if cell != ""]
                    
                    # Se a linha tiver os 6 dados principais (Id, File, Total Geral, Receita, Custo, Total Net)
                    if len(cleaned_row) >= 6:
                        # Pegamos os 6 primeiros itens caso o PDF crie colunas extras fantasmas
                        all_rows.append(cleaned_row[:6])

    # Criar o DataFrame com os cabeçalhos baseados no seu arquivo
    columns = [
        "Índice", 
        "Id File", 
        "Total Geral", 
        "Receita Operacional", 
        "Custo Operacao Rateio", 
        "Total NET - Previsto"
    ]
    
    df = pd.DataFrame(all_rows, columns=columns)
    
    # Limpar a coluna 'Índice' e 'Id File' (manter como string ou converter para int)
    df['Índice'] = pd.to_numeric(df['Índice'], errors='coerce')
    df['Id File'] = df['Id File'].astype(str)
    
    # Remover linhas onde o Índice falhou em ser número (linhas sujas que passaram pelo filtro)
    df = df.dropna(subset=['Índice'])
    
    # Converter as colunas financeiras
    financial_cols = ["Total Geral", "Receita Operacional", "Custo Operacao Rateio", "Total NET - Previsto"]
    for col in financial_cols:
        df[col] = df[col].apply(clean_brazilian_currency)
        
    return df

# --- Interface do Streamlit ---
st.title("📄 Conversor de PDF para Excel")
st.write("Faça o upload do arquivo de comissionamento (ex: *Brocker - Manutenção Comissionada ManageTour - Agosto-2026 - File.pdf*).")

uploaded_file = st.file_uploader("Escolha um arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    st.info("Processando o PDF... Isso pode levar alguns segundos dependendo do tamanho (ex: 227 páginas).")
    
    try:
        # Extrai os dados
        df = process_pdf(uploaded_file)
        
        st.success(f"Extração concluída! Encontrados {len(df)} registros.")
        
        # Mostra uma prévia dos dados
        st.write("Prévia dos dados:")
        st.dataframe(df.head(10))
        
        # Converte para Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Comissionamento')
        
        excel_data = output.getvalue()
        
        # Botão de download
        st.download_button(
            label="📥 Baixar arquivo Excel (.xlsx)",
            data=excel_data,
            file_name=uploaded_file.name.replace('.pdf', '.xlsx'),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
