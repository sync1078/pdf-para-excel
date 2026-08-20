import io
import re
import pdfplumber
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Conversor de Relatórios PDF para Excel",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Conversor de Relatórios (PDF → Excel)")

tipo_relatorio = st.sidebar.selectbox(
    "Selecione o modelo do relatório:",
    [
        "Sumário / Manutenção ComISSIONADA (Tabela por Id File)",
        "Brocker / Bustour Detalhado (PDF por Cliente)",
    ],
)

uploaded_file = st.file_uploader(
    "Arraste ou selecione o arquivo PDF", type=["pdf"]
)


def to_float(val_str):
    """Converte valores no padrão brasileiro (1.218,00) para float."""
    if not val_str or str(val_str).strip() == "":
        return 0.0
    clean_str = (
        str(val_str)
        .replace(".", "")
        .replace(",", ".")
        .replace(":", ".")
        .strip()
    )
    try:
        return float(clean_str)
    except ValueError:
        return 0.0


# --- PARSER USANDO PDFPLUMBER (PARA O SUMÁRIO) ---
def parse_sumario_pdfplumber(pdf_file):
    records = []

    # Regex ajustada para capturar o Id File mesmo sem o ponto (ex: 597.762 ou 597762)
    pattern_id = re.compile(r"\b(\d{3}[\.:]?\d{3})\b")

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    row_clean = [
                        str(cell).strip() if cell is not None else ""
                        for cell in row
                    ]
                    cells = [c for c in row_clean if c != ""]

                    # Descartar cabeçalhos, rodapés e a linha final do próprio PDF
                    row_text = " ".join(cells).upper()
                    if (
                        "SUMÁRIO" in row_text
                        "TOTAL GERAL" in row_text
                        or "RECEITA OPERACIONAL" in row_text
                        or "ID FILE" in row_text
                    ):
                        continue

                    for idx, cell in enumerate(cells):
                        match_id = pattern_id.search(cell)

                        if match_id:
                            id_file_raw = (
                                match_id.group(1)
                                .replace(":", ".")
                                .replace(" ", ".")
                            )
                            # Se o ID não tiver ponto, formata para manter padrão visual
                            if "." not in id_file_raw and len(id_file_raw) == 6:
                                id_file_str = (
                                    f"{id_file_raw[:3]}.{id_file_raw[3:]}"
                                )
                            else:
                                id_file_str = id_file_raw

                            remaining_cells = cells[idx + 1 :]

                            # Garante a extração dos 4 valores monetários
                            if len(remaining_cells) >= 4:
                                total_geral = to_float(remaining_cells[0])
                                receita_op = to_float(remaining_cells[1])
                                custo_rateio = to_float(remaining_cells[2])
                                total_net = to_float(remaining_cells[3])

                                records.append({
                                    "ID_FILE": id_file_str,
                                    "TOTAL_GERAL": total_geral,
                                    "RECEITA_OPERACIONAL": receita_op,
                                    "CUSTO_OPERACAO_RATEIO": custo_rateio,
                                    "TOTAL_NET_PREVISTO": total_net,
                                })
                            break

    df = pd.DataFrame(records)

    if not df.empty:
        # ATENÇÃO: drop_duplicates foi removido para não apagar vendas do mesmo File
        row_total = {
            "ID_FILE": "TOTAL GERAL",
            "TOTAL_GERAL": round(df["TOTAL_GERAL"].sum(), 2),
            "RECEITA_OPERACIONAL": round(df["RECEITA_OPERACIONAL"].sum(), 2),
            "CUSTO_OPERACAO_RATEIO": round(
                df["CUSTO_OPERACAO_RATEIO"].sum(), 2
            ),
            "TOTAL_NET_PREVISTO": round(df["TOTAL_NET_PREVISTO"].sum(), 2),
        }
        df = pd.concat([df, pd.DataFrame([row_total])], ignore_index=True)

    return df


# --- EXECUÇÃO STREAMLIT ---
if uploaded_file is not None:
    st.info(f"Processando arquivo no modelo **{tipo_relatorio}**...")
    try:
        if (
            tipo_relatorio
            == "Sumário / Manutenção ComISSIONADA (Tabela por Id File)"
        ):
            df_result = parse_sumario_pdfplumber(uploaded_file)
        else:
            # Mantém a função anterior se for o modelo detalhado antigo
            st.warning("Selecione o modelo Sumário para este arquivo.")
            df_result = pd.DataFrame()

        if not df_result.empty:
            st.success(
                f"Pronto! {len(df_result)-1} registros foram convertidos com sucesso."
            )
            st.dataframe(df_result, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_result.to_excel(
                    writer, index=False, sheet_name="Sumario_Convertido"
                )

            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="Sumario_Brocker_Convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("Nenhum registro foi encontrado no PDF enviado.")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
