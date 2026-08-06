import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor PDF para Excel", page_icon="📊", layout="wide"
)

st.title("📊 Conversor de Relatórios (PDF → Excel)")
st.write(
    "Faça upload do PDF do relatório para transformar os dados em uma planilha pronta para Banco de Dados."
)

uploaded_file = st.file_uploader(
    "Arraste e solte ou selecione o arquivo PDF aqui", type=["pdf"]
)


def extract_data_from_pdf(pdf_file):
    records = []
    current_voucher = None
    current_client = None
    current_channel = None

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            lines = text.split("\n")
            for line in lines:
                line_str = line.strip()

                # Ignora linhas de cabeçalho e rodapé repetidas no relatório
                if (
                    "FEE" in line_str
                    or "DATA SERVIÇO" in line_str
                    or "Pagina" in line_str
                    or "TOTAL" in line_str
                ):
                    continue

                # Captura a linha do Cliente/Voucher (Ex: 721074-ELIAS FRANCISCO DE AGUIAR JUNIOR (SITE BROCKER))
                voucher_match = re.match(
                    r"^(\d+)\s*[-–]?\s*(.*?)\s*\((.*?)\)", line_str
                )
                if voucher_match:
                    current_voucher = voucher_match.group(1)
                    current_client = voucher_match.group(2).strip()
                    current_channel = voucher_match.group(3).strip()
                    continue

                # Captura linhas que começam com Data (Ex: 03/07/26 ou 03/07/2026)
                date_match = re.match(
                    r"^(\d{2}/\d{2}/\d{2,4})\s+(.*)", line_str
                )
                if date_match and current_voucher:
                    data_servico = date_match.group(1)
                    resto_linha = date_match.group(2)

                    # Separa os campos por 'pipes' (|)
                    parts = [
                        p.strip() for p in resto_linha.split("|") if p.strip()
                    ]

                    # Estruturação básica das informações extraídas
                    servico_nome = parts[0] if len(parts) > 0 else ""
                    adt = parts[1] if len(parts) > 1 else "0"
                    chd = parts[2] if len(parts) > 2 else "0"
                    inf = parts[3] if len(parts) > 3 else "0"
                    qtd = parts[4] if len(parts) > 4 else "0"

                    records.append(
                        {
                            "ID_Voucher": current_voucher,
                            "Cliente": current_client,
                            "Canal_Origem": current_channel,
                            "Data_Servico": data_servico,
                            "Servico": servico_nome,
                            "ADT": adt,
                            "CHD": chd,
                            "INF": inf,
                            "QTD": qtd,
                            "Linha_Completa_Original": line_str,
                            "Pagina_PDF": page_num,
                        }
                    )

    return pd.DataFrame(records)


if uploaded_file is not None:
    st.info("Processando o arquivo PDF... Por favor, aguarde.")
    try:
        df = extract_data_from_pdf(uploaded_file)

        if not df.empty:
            st.success(
                f"Sucesso! {len(df)} linhas de serviços foram encontradas e estruturadas."
            )

            # Exibe prévia dos dados
            st.subheader("Prévia dos dados estruturados:")
            st.dataframe(df.head(20), use_container_width=True)

            # Cria a planilha Excel na memória
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Dados_Estruturados")

            # Botão para baixar
            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=output.getvalue(),
                file_name="relatorio_estruturado_bd.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning(
                "Nenhum registro foi encontrado. Verifique se o formato do PDF está correto."
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o PDF: {e}")
