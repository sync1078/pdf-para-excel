import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Conversor PDF para Excel", page_icon="📊", layout="wide"
)

st.title("📊 Conversor de Relatórios Brocker / Bustour (PDF → Excel)")
st.write(
    "Upload do relatório em PDF para gerar a planilha formatada no padrão exato do seu Banco de Dados."
)

uploaded_file = st.file_uploader(
    "Arraste ou selecione o arquivo PDF do relatório", type=["pdf"]
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


def parse_brocker_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
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

            # Filtrar rodapés e cabeçalhos repetidos
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

            # Captura o cabeçalho do Cliente / File
            client_match = re.search(
                r"^(\d+)\s*[-–]?\s*(.*?)\s*\(([^\)]+)\)", line_str
            )
            if client_match:
                current_file = client_match.group(1).strip()
                current_client = client_match.group(2).strip()
                current_site = client_match.group(3).strip()
                continue

            # Captura a linha do Serviço (que possui a Data DD/MM/YY)
            date_match = re.search(r"(\d{2}/\d{2}/\d{2,4})", line_str)
            if date_match and current_file:
                data_servico = date_match.group(1)
                date_start = date_match.start()
                date_end = date_match.end()

                servico_nome = line_str[:date_start].strip()
                resto = line_str[date_end:].strip()

                # Extrai a Categoria do Serviço (após o %)
                pct_match = re.search(r"(\d+[\.,]\d+\s*%)", resto)
                if pct_match:
                    categoria = resto[pct_match.end() :].strip()
                    nums_part = resto[: pct_match.start()].strip()
                else:
                    categoria = ""
                    nums_part = resto

                # Extrai todos os valores monetários com vírgula
                monetaries = re.findall(r"\d+(?:\.\d{3})*,\d{2}", nums_part)

                # Limpa os números monetários para isolar inteiros (ADT, CHD, INF, QTD)
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

        # Reajuste Fino: ajusta as linhas para que a soma da coluna no Excel dê exatamente o total do rodapé
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


if uploaded_file is not None:
    st.info("Processando o arquivo PDF...")
    try:
        df_result = parse_brocker_pdf(uploaded_file)

        if not df_result.empty:
            st.success(
                f"Pronto! {len(df_result)-1} registros foram convertidos com sucesso."
            )
            st.dataframe(df_result.head(20), use_container_width=True)

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
            st.warning("Nenhum registro foi encontrado no PDF enviado.")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
