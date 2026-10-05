import streamlit as st
import pandas as pd
import pdfplumber
import io

st.set_page_config(page_title="Conversor de Comissões PDF para Excel", layout="centered")

def extract_table_from_pdf(file_obj):
    """
    Extrai as tabelas de um arquivo PDF usando o pdfplumber, página por página.
    """
    all_rows = []
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Limpa quebras de linha (enters) indesejados nas células de texto
                        cleaned_row = [str(cell).replace('\n', ' ').strip() if cell is not None else "" for cell in row]
                        all_rows.append(cleaned_row)
    except Exception as e:
        st.error(f"Erro ao ler o PDF: {e}")
    return all_rows

def process_dataframe(data, report_type):
    """
    Estrutura os dados extraídos em um DataFrame limpo, 
    identificando o cabeçalho correto e removendo repetições.
    """
    if not data:
        return pd.DataFrame()
    
    df = pd.DataFrame(data)
    
    # Identificar a palavra-chave do cabeçalho de acordo com o tipo de relatório
    if report_type == "MANUTENÇÃO COMISSIONADA MANAGETOUR":
        header_keyword = "Id"
    else:
        # Serve tanto para APP quanto para SITE
        header_keyword = "SERVIÇO"
        
    # Encontrar a linha exata onde o cabeçalho original começa
    header_idx = -1
    for idx, row in df.iterrows():
        if any(header_keyword in str(val).strip() for val in row.values):
            header_idx = idx
            break
            
    if header_idx != -1:
        # Define a linha encontrada como cabeçalho
        df.columns = df.iloc[header_idx]
        
        # Remove as linhas anteriores (títulos do PDF) e a própria linha de cabeçalho
        df = df.iloc[header_idx + 1:]
        
        # Remove as linhas que repetem o cabeçalho ao longo das quebras de página
        # Ex: Quando a primeira coluna for igual ao nome da própria coluna
        first_col_name = df.columns[0]
        df = df[df.iloc[:, 0] != first_col_name]
        
        # Remove linhas totalmente vazias
        df = df.dropna(how='all')
        
    return df

# --- INTERFACE STREAMLIT ---
st.title("📊 Conversor de Relatórios: PDF para Excel")
st.markdown("Transforme os relatórios de comissionamento em arquivos `.xlsx` limpos e corridos, sem perder nenhum dado.")

report_type = st.selectbox(
    "Selecione o tipo de relatório que deseja converter:",
    ("MANUTENÇÃO COMISSIONADA SITE", "MANUTENÇÃO COMISSIONADA APP", "MANUTENÇÃO COMISSIONADA MANAGETOUR")
)

# Lógica condicional para o Uploader
if report_type == "MANUTENÇÃO COMISSIONADA APP":
    st.info("Para o relatório APP, faça o upload dos **2 arquivos quinzenais** simultaneamente.")
    uploaded_files = st.file_uploader("Arraste ou selecione os arquivos PDF", type=['pdf'], accept_multiple_files=True)
else:
    uploaded_files = st.file_uploader("Arraste ou selecione o arquivo PDF mensal", type=['pdf'], accept_multiple_files=False)
    if uploaded_files is not None:
        uploaded_files = [uploaded_files] # Transforma em lista para o laço funcionar igualmente

if uploaded_files:
    if st.button("Converter para Excel"):
        with st.spinner("Extraindo dados com precisão cirúrgica. Por favor, aguarde..."):
            all_dataframes = []
            
            # Processa todos os arquivos upados (1 ou 2)
            for file in uploaded_files:
                raw_data = extract_table_from_pdf(file)
                if raw_data:
                    df_cleaned = process_dataframe(raw_data, report_type)
                    all_dataframes.append(df_cleaned)
            
            if all_dataframes:
                # Concatena os arquivos (no caso do APP, mescla os 2 mantendo a ordem contínua)
                final_df = pd.concat(all_dataframes, ignore_index=True)
                
                # Salva no buffer em memória para não precisar criar arquivo físico no servidor
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Comissoes')
                
                st.success("✅ Relatório convertido com sucesso! Nenhum dado foi perdido.")
                
                st.download_button(
                    label="📥 Baixar Arquivo Excel (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"{report_type.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Não foi possível encontrar tabelas válidas nos PDFs enviados.")
