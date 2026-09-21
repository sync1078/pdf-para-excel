import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Conversor ManageTour File → Excel",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Conversor ManageTour (Sumário por File)")
st.write(
    "Upload de relatórios PDF em formato de Sumário por File para extrair exclusivamente as colunas: "
    "**ID_FILE**, **TOTAL_GERAL**, **RECEITA_OPERACIONAL**, **CUSTO_OPERACAO_RATEIO** e **TOTAL_NET_PREVISTO**."
)

uploaded_files = st.file_uploader(
    "Arraste ou selecione um ou mais arquivos PDF do relatório por File",
    type=["pdf"],
    accept_multiple_files=True,
)


def parse_managetour_file_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    records = []

    def to_float(val_str):
        return float(val_str.replace(".", "").replace(",", "."))

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()

            # Filtrar linhas de cabeçalho, rodapé, links e datas
            if (
                not line_str
                or "SUMÁRIO DE" in line_str
                or "Sumário de" in line_str
                or "Id File" in line_str
                or "Total Geral" in line_str
                or "https://" in line_str
                or re.match(r"^\d{2}/\d{2}/\d{4}", line_str)
            ):
                continue

            # Capturar padrão: ID_FILE + 4 valores monetários
            file_match = re.search(
                r"(\d{3}\.\d{3}\.?|\b\d{6}\b)\s+([\d\.\,]+)\s+([\d\.\,]+)\s+([\d\.\,]+)\s+([\d\.\,]+)",
                line_str,
            )
            if file_match:
                file_id = file_match.group(1).replace(".", "")
                tot_geral = to_float(file_match.group(2))
                rec_oper = to_float(file_match.group(3))
                custo_op = to_float(file_match.group(4))
                tot_net = to_float(file_match.group(5))

                records.append({
                    "ID_FILE": int(file_id) if file_id.isdigit() else file_id,
                    "TOTAL_GERAL": tot_geral,
                    "RECEITA_OPERACIONAL": rec_oper,
                    "CUSTO_OPERACAO_RATEIO": custo_op,
                    "TOTAL_NET_PREVISTO": tot_net,
                })

    return pd.DataFrame(records)


if uploaded_files:
    all_dfs = []

    with st.spinner("Processando relatórios PDF..."):
        for pdf_file in uploaded_files:
            df_part = parse_managetour_file_pdf(pdf_file)
            if not df_part.empty:
                all_dfs.append(df_part)

    if all_dfs:
        df_full = pd.concat(all_dfs, ignore_index=True)

        st.markdown("---")
        st.subheader("📌 Indicadores e Resumo dos Dados")

        tot_g = round(df_full["TOTAL_GERAL"].sum(), 2)
        rec_o = round(df_full["RECEITA_OPERACIONAL"].sum(), 2)
        cus_o = round(df_full["CUSTO_OPERACAO_RATEIO"].sum(), 2)
        tot_n = round(df_full["TOTAL_NET_PREVISTO"].sum(), 2)
        total_records = len(df_full)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Qtd. Registros", f"{total_records:,}")
        c2.metric("Total Geral", f"R$ {tot_g:,.2f}")
        c3.metric("Receita Operacional", f"R$ {rec_o:,.2f}")
        c4.metric("Custo Op. Rateio", f"R$ {cus_o:,.2f}")
        c5.metric("Total NET Previsto", f"R$ {tot_n:,.2f}")

        # Linha de soma ao final
        row_total = {
            "ID_FILE": "TOTAL GERAL",
            "TOTAL_GERAL": tot_g,
            "RECEITA_OPERACIONAL": rec_o,
            "CUSTO_OPERACAO_RATEIO": cus_o,
            "TOTAL_NET_PREVISTO": tot_n,
        }
        df_export = pd.concat(
            [df_full, pd.DataFrame([row_total])], ignore_index=True
        )

        st.markdown("### 📋 Tabela Convertida")
        st.dataframe(df_export, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_export.to_excel(
                writer, index=False, sheet_name="ManageTour_File"
            )

        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name="ManageTour_Relatorio_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("Nenhum registro válido foi encontrado no PDF selecionado.")
