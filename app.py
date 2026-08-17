import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Conversor PDF para Excel PRO", page_icon="📊", layout="wide"
)

st.title("📊 Conversor de Relatórios Brocker / Bustour / ManageTour")
st.write(
    "Upload de relatórios em PDF (Detalhados ou Sumários) para gerar planilhas formatadas para Banco de Dados."
)

uploaded_files = st.file_uploader(
    "Arraste ou selecione um ou mais arquivos PDF dos relatórios",
    type=["pdf"],
    accept_multiple_files=True,
)


def adjust_fees_to_match_target(df_input, target_fee):
    df = df_input.copy()
    current_sum = round(df["VALOR_FEE"].sum(), 2)
    diff = round(current_sum - target_fee, 2)

    if diff == 0:
        return df

    cents_to_adjust = int(round(abs(diff) * 100))
    step = -0.01 if diff > 0 else 0.01

    df["raw_fee"] = df["VALOR_VENDA"] * 0.01
    df["round_diff"] = df["VALOR_FEE"] - df["raw_fee"]

    if diff > 0:
        candidates = (
            df[df["round_diff"] > 0]
            .sort_values(by="round_diff", ascending=False)
            .index
        )
    else:
        candidates = (
            df[df["round_diff"] < 0]
            .sort_values(by="round_diff", ascending=True)
            .index
        )

    adjusted_count = 0
    for idx in candidates:
        if adjusted_count >= cents_to_adjust:
            break
        df.loc[idx, "VALOR_FEE"] = round(df.loc[idx, "VALOR_FEE"] + step, 2)
        adjusted_count += 1

    df.drop(columns=["raw_fee", "round_diff"], inplace=True)
    return df


def parse_managetour_summary_pdf(reader):
    records = []
    current_category = "INDEFINIDO"

    def to_float(val_str):
        return float(val_str.replace(".", "").replace(",", "."))

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()

            # Ignorar cabeçalhos e números de página
            if (
                not line_str
                or "SUMÁRIO DE" in line_str
                or "Nome Categoria" in line_str
                or "Total Geral" in line_str
                or line_str.isdigit()
            ):
                continue

            # Detectar linha de mudança de Categoria (Ex: AÉREO, DIVERSOS, HOSPEDAGEM, RECEPTIVO)
            if re.match(r"^[A-ZÁÉÍÓÚÂÊÔÃÕÇ\s]{3,}$", line_str) and not re.search(
                r"\d", line_str
            ):
                current_category = line_str.strip()
                continue

            # Detectar linha de Id File e Total (Ex: 676.538 4.866,03 ou 720.584 268,00)
            file_match = re.search(
                r"(\d{3}\.\d{3}|\d{6})\s+([\d\.\,]+)", line_str
            )
            if file_match:
                file_id = file_match.group(1).replace(".", "")
                valor_str = file_match.group(2)
                try:
                    total_val = to_float(valor_str)
                    records.append({
                        "CATEGORIA": current_category,
                        "FILE": int(file_id) if file_id.isdigit() else file_id,
                        "TOTAL_GERAL": total_val,
                    })
                except:
                    pass

    df = pd.DataFrame(records)
    if not df.empty:
        total_soma = round(df["TOTAL_GERAL"].sum(), 2)
        row_total = {
            "CATEGORIA": "TOTAL GERAL",
            "FILE": None,
            "TOTAL_GERAL": total_soma,
        }
        df = pd.concat([df, pd.DataFrame([row_total])], ignore_index=True)

    return df


def parse_brocker_pdf(reader):
    records = []

    current_file = ""
    current_client = ""
    current_site = ""

    def to_float(val_str):
        return float(val_str.replace(".", "").replace(",", "."))

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()

            if (
                not line_str
                or "FEE -" in line_str
                or "DATA SERVIÇO" in line_str
                or "EMISSÃO:" in line_str
                or "ORIGEM:" in line_str
                or "SERVIÇO:" in line_str
                or "Pagina" in line_str
                or "Página" in line_str
                or "TOTAL" in line_str
            ):
                continue

            client_match = re.search(
                r"^(\d+)\s*[-–]?\s*(.*?)\s*\(([^\)]+)\)", line_str
            )
            if client_match:
                current_file = client_match.group(1).strip()
                current_client = client_match.group(2).strip()
                current_site = client_match.group(3).strip()
                continue

            date_match = re.search(r"(\d{2}/\d{2}/\d{2,4})", line_str)
            if date_match and current_file:
                data_servico = date_match.group(1)
                date_start = date_match.start()
                date_end = date_match.end()

                servico_nome = line_str[:date_start].strip()
                resto = line_str[date_end:].strip()

                pct_match = re.search(r"(\d+[\.,]\d+\s*%)", resto)
                if pct_match:
                    categoria = resto[pct_match.end() :].strip()
                    nums_part = resto[: pct_match.start()].strip()
                else:
                    categoria = ""
                    nums_part = resto

                monetaries = re.findall(r"\d+(?:\.\d{3})*,\d{2}", nums_part)

                cleaned_nums = nums_part
                for m in monetaries:
                    cleaned_nums = cleaned_nums.replace(m, " ")

                integers = re.findall(r"\b\d+\b", cleaned_nums)

                adt = float(integers[0]) if len(integers) > 0 else 0.0
                chd = float(integers[1]) if len(integers) > 1 else 0.0
                inf = float(integers[2]) if len(integers) > 2 else 0.0
                qtd = (
                    float(integers[3])
                    if len(integers) > 3
                    else (adt + chd + inf)
                )

                valor_fee = to_float(monetaries[0]) if len(monetaries) > 0 else 0.0
                valor_venda = (
                    to_float(monetaries[1]) if len(monetaries) > 1 else 0.0
                )
                voucher_recibo = (
                    to_float(monetaries[2]) if len(monetaries) > 2 else 0.0
                )

                records.append({
                    "FILE": int(current_file)
                    if str(current_file).isdigit()
                    else current_file,
                    "NOME_CLIENTE": current_client,
                    "SITE_ORIGEM": current_site,
                    "DATA_SERVICO": data_servico,
                    "SERVICO": servico_nome,
                    "CATEGORIA_SERVICO": categoria,
                    "ADT": adt,
                    "CHD": chd,
                    "INF": inf,
                    "QTD": qtd,
                    "VOUCHER_RECIBO": voucher_recibo,
                    "TARIFA": None,
                    "VALOR_VENDA": valor_venda,
                    "VALOR_FEE": valor_fee,
                })

    df = pd.DataFrame(records)

    if not df.empty:
        total_venda = round(df["VALOR_VENDA"].sum(), 2)
        total_fee = round(total_venda * 0.01, 2)
        df = adjust_fees_to_match_target(df, total_fee)

    return df


def process_pdf_file(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    first_page_text = reader.pages[0].extract_text() if reader.pages else ""

    # Identificar o tipo de relatório
    if "SUMÁRIO DE" in first_page_text or "Total Geral (Soma)" in first_page_text:
        df = parse_managetour_summary_pdf(reader)
        pdf_type = "sumario"
    else:
        df = parse_brocker_pdf(reader)
        pdf_type = "detalhado"

    return df, pdf_type


if uploaded_files:
    all_dfs_detalhado = []
    all_dfs_sumario = []

    with st.spinner("Processando arquivos PDF..."):
        for file in uploaded_files:
            df_result, p_type = process_pdf_file(file)
            if not df_result.empty:
                if p_type == "detalhado":
                    # Remove a linha total individual temporariamente para concatenar no final
                    df_clean = df_result[df_result["FILE"] != "TOTAL"]
                    all_dfs_detalhado.append(df_clean)
                else:
                    df_clean = df_result[df_result["CATEGORIA"] != "TOTAL GERAL"]
                    all_dfs_sumario.append(df_clean)

    # Exibição de Relatórios Detalhados
    if all_dfs_detalhado:
        df_full_det = pd.concat(all_dfs_detalhado, ignore_index=True)

        st.markdown("---")
        st.subheader("📈 Painel Geral de Vendas (Relatórios Detalhados)")

        col1, col2, col3, col4 = st.columns(4)
        total_venda_geral = round(df_full_det["VALOR_VENDA"].sum(), 2)
        total_fee_geral = round(total_venda_geral * 0.01, 2)
        total_passag = df_full_det["QTD"].sum()
        total_files = df_full_det["FILE"].nunique()

        col1.metric("Total de Vendas", f"R$ {total_venda_geral:,.2f}")
        col2.metric("Total Fee (1%)", f"R$ {total_fee_geral:,.2f}")
        col3.metric("Total de Passageiros", f"{int(total_passag):,}")
        col4.metric("Qtd. de Vouchers/Files", f"{total_files:,}")

        # Visualização de Gráficos
        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("**Vendas por Canal de Origem**")
            st.bar_chart(
                df_full_det.groupby("SITE_ORIGEM")["VALOR_VENDA"]
                .sum()
                .sort_values(ascending=False)
            )
        with col_r:
            st.markdown("**Top 10 Serviços Mais Vendidos (R$)**")
            st.bar_chart(
                df_full_det.groupby("SERVICO")["VALOR_VENDA"]
                .sum()
                .sort_values(ascending=False)
                .head(10)
            )

        # Reajuste fino acumulado para a exportação
        df_full_det = adjust_fees_to_match_target(
            df_full_det, total_fee_geral
        )

        row_total = {
            "FILE": "TOTAL",
            "NOME_CLIENTE": None,
            "SITE_ORIGEM": None,
            "DATA_SERVICO": None,
            "SERVICO": None,
            "CATEGORIA_SERVICO": None,
            "ADT": None,
            "CHD": None,
            "INF": None,
            "QTD": None,
            "VOUCHER_RECIBO": None,
            "TARIFA": None,
            "VALOR_VENDA": total_venda_geral,
            "VALOR_FEE": total_fee_geral,
        }
        df_export_det = pd.concat(
            [df_full_det, pd.DataFrame([row_total])], ignore_index=True
        )

        buffer_det = io.BytesIO()
        with pd.ExcelWriter(buffer_det, engine="openpyxl") as writer:
            df_export_det.to_excel(
                writer, index=False, sheet_name="Dados_Detalhados"
            )

        st.download_button(
            label="📥 Baixar Planilha Excel Detalhada (.xlsx)",
            data=buffer_det.getvalue(),
            file_name="Relatorio_Detalhado_Convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Exibição de Relatórios de Sumário
    if all_dfs_sumario:
        df_full_sum = pd.concat(all_dfs_sumario, ignore_index=True)

        st.markdown("---")
        st.subheader("📋 Relatório Sumário (ManageTour)")

        total_sumario = round(df_full_sum["TOTAL_GERAL"].sum(), 2)
        st.metric("Total Geral Sumário", f"R$ {total_sumario:,.2f}")

        row_total_sum = {
            "CATEGORIA": "TOTAL GERAL",
            "FILE": None,
            "TOTAL_GERAL": total_sumario,
        }
        df_export_sum = pd.concat(
            [df_full_sum, pd.DataFrame([row_total_sum])], ignore_index=True
        )

        st.dataframe(df_export_sum, use_container_width=True)

        buffer_sum = io.BytesIO()
        with pd.ExcelWriter(buffer_sum, engine="openpyxl") as writer:
            df_export_sum.to_excel(
                writer, index=False, sheet_name="Sumario_ManageTour"
            )

        st.download_button(
            label="📥 Baixar Planilha Excel de Sumário (.xlsx)",
            data=buffer_sum.getvalue(),
            file_name="Relatorio_Sumario_ManageTour.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
