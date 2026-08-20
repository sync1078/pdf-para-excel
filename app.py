import io
import re
import pandas as pd
import pdfplumber
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
    """Converte valores no padrão brasileiro (1.218,00 ou 88,00) para float."""
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


# --- PARSER 1: SUMÁRIO (EXTRAÇÃO VIA PALAVRAS E POSICIONAMENTO) ---
def parse_sumario_pdfplumber(pdf_file):
    records = []

    # Regex para identificar ID FILE com 6 dígitos (ex: 597.762, 609.176, 679:040)
    pattern_id = re.compile(r"^(\d{3}[\.:]?\d{3})$")
    # Regex para identificar números monetários (ex: 1.218,00 ou 88,00 ou 0,00)
    pattern_money = re.compile(r"^\d+(?:\.\d{3})*,\d{2}$")

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                page.flush_cache()
                continue

            # Agrupa as palavras da página pelo eixo Y (mesma linha visual)
            lines_dict = {}
            for w in words:
                top_key = round(w["top"], 1)
                # Tolera pequenas variações de alinhamento vertical
                matched_key = None
                for k in lines_dict.keys():
                    if abs(k - top_key) <= 3:
                        matched_key = k
                        break

                if matched_key is None:
                    lines_dict[top_key] = [w]
                else:
                    lines_dict[matched_key].append(w)

            # Processa linha por linha na ordem vertical da página
            for top_key in sorted(lines_dict.keys()):
                line_words = lines_dict[top_key]
                # Ordena as palavras da esquerda para a direita (eixo X)
                line_words.sort(key=lambda x: x["x0"])

                tokens = [
                    w["text"].replace("|", "").strip()
                    for w in line_words
                    if w["text"].replace("|", "").strip() != ""
                ]

                line_str = " ".join(tokens).upper()

                # Ignora cabeçalhos, rodapés e linha de Total do PDF
                if (
                    "SUMÁRIO" in line_str
                    or "RECEITA OPERACIONAL" in line_str
                    or "TOTAL GERAL" in line_str
                    or "ID FILE" in line_str
                    or "HTTPS://" in line_str.lower()
                ):
                    continue

                # Localiza o Id File dentro dos tokens da linha
                for idx, token in enumerate(tokens):
                    match_id = pattern_id.match(token)
                    if match_id:
                        clean_digits = (
                            token.replace(".", "").replace(":", "").strip()
                        )

                        if len(clean_digits) == 6 and clean_digits.isdigit():
                            id_file_str = (
                                f"{clean_digits[:3]}.{clean_digits[3:]}"
                            )
                            remaining_tokens = tokens[idx + 1 :]

                            # Filtra apenas os tokens que são valores monetários válidos
                            money_vals = [
                                t
                                for t in remaining_tokens
                                if pattern_money.match(t)
                            ]

                            if len(money_vals) >= 4:
                                records.append({
                                    "ID_FILE": id_file_str,
                                    "TOTAL_GERAL": to_float(money_vals[0]),
                                    "RECEITA_OPERACIONAL": to_float(
                                        money_vals[1]
                                    ),
                                    "CUSTO_OPERACAO_RATEIO": to_float(
                                        money_vals[2]
                                    ),
                                    "TOTAL_NET_PREVISTO": to_float(
                                        money_vals[3]
                                    ),
                                })
                                break

            page.flush_cache()

    df = pd.DataFrame(records)

    if not df.empty:
        # Gera o TOTAL GERAL recalculado via Python
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


# --- PARSER 2: BROCKER DETALHADO ---
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
            df_result = parse_sumario_pdfplumber(uploaded_file)
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
