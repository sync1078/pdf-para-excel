import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Conversor de Relatórios PDF para Excel",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Conversor de Relatórios (PDF → Excel)")
st.write(
    "Faça o upload do relatório em PDF e selecione o tipo para gerar a planilha formatada."
)

tipo_relatorio = st.sidebar.selectbox(
    "Selecione o modelo do relatório:",
    [
        "Sumário / Manutenção Comissionada (Tabela por Id File)",
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


def adjust_fees_to_match_target(df_input, target_fee):
    """Ajuste fino de centavos para conciliação no Excel."""
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


# --- PARSER 1: SUMÁRIO (OTIMIZADO PARA LEITURA RÁPIDA SEM ESTOURO DE MEMÓRIA) ---
def parse_sumario_pypdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    records = []

    # Captura valores com tolerância para quebras de linha e pipes do PDF
    pattern = re.compile(
        r"(\d{3}[\.:]?\d{3})\s*\|?\s*([\d\.,]+)\s*\|?\s*([\d\.,]+)\s*\|?\s*([\d\.,]+)\s*\|?\s*([\d\.,]+)"
    )

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()

            if (
                not line_str
                or "SUMÁRIO" in line_str.upper()
                or "TOTAL GERAL" in line_str.upper()
                or "RECEITA OPERACIONAL" in line_str.upper()
                or "ID FILE" in line_str.upper()
                or "HTTPS://" in line_str.lower()
            ):
                continue

            match = pattern.search(line_str)
            if match:
                id_file_raw = (
                    match.group(1).replace(":", ".").replace(" ", ".")
                )

                if "." not in id_file_raw and len(id_file_raw) == 6:
                    id_file_str = f"{id_file_raw[:3]}.{id_file_raw[3:]}"
                else:
                    id_file_str = id_file_raw

                total_geral = to_float(match.group(2))
                receita_op = to_float(match.group(3))
                custo_rateio = to_float(match.group(4))
                total_net = to_float(match.group(5))

                records.append({
                    "ID_FILE": id_file_str,
                    "TOTAL_GERAL": total_geral,
                    "RECEITA_OPERACIONAL": receita_op,
                    "CUSTO_OPERACAO_RATEIO": custo_rateio,
                    "TOTAL_NET_PREVISTO": total_net,
                })

    df = pd.DataFrame(records)

    if not df.empty:
        # Recalcula o Total Geral com precisão decimal
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


# --- PARSER 2: MODELO BROCKER / BUSTOUR DETALHADO ---
def parse_brocker_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    records = []

    current_file = ""
    current_client = ""
    current_site = ""

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

                monetaries = re.findall(
                    r"\d+(?:\.\d{3})*,\d{2}", nums_part
                )
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

                valor_fee = (
                    to_float(monetaries[0]) if len(monetaries) > 0 else 0.0
                )
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
            "VALOR_VENDA": total_venda,
            "VALOR_FEE": total_fee,
        }
        df = pd.concat([df, pd.DataFrame([row_total])], ignore_index=True)

    return df


# --- PROCESSAMENTO STREAMLIT ---
if uploaded_file is not None:
    st.info(f"Processando arquivo no modelo **{tipo_relatorio}**...")
    try:
        if (
            tipo_relatorio
            == "Sumário / Manutenção Comissionada (Tabela por Id File)"
        ):
            df_result = parse_sumario_pypdf(uploaded_file)
        else:
            df_result = parse_brocker_pdf(uploaded_file)

        if not df_result.empty:
            st.success(
                f"Sucesso! {len(df_result)-1} registros foram convertidos."
            )
            st.dataframe(df_result, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_result.to_excel(
                    writer, index=False, sheet_name="Dados_Convertidos"
                )

            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="Relatorio_Convertido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("Nenhum registro correspondente foi encontrado no PDF.")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
