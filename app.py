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

            # Filtrar rodapés e cabeçalhos repetidos
            if (
                not line_str
                or "FEE -" in line_str
                or "DATA SERVIÇO" in line_str
                or "EMISSÃO:" in line_str
                or "ORIGEM:" in line_str
                if "SERVIÇO:" in line_str
                or "Pagina" in line_str
                or "TOTAL" in line_str:
                    continue

            # Captura o cabeçalho do Cliente / File
            # Ex: "721074 - ELIAS FRANCISCO DE AGUIAR JUNIOR (SITE BROCKER)"
            client_match = re.search(
                r"^(\d+)\s*[-–]?\s*(.*?)\s*\((SITE\s+[^\)]+)\)", line_str
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
                    numeros_str = resto[: pct_match.start()].strip()
                else:
                    categoria = ""
                    numeros_str = resto

                # Extrai os números e valores
                tokens = numeros_str.split()

                adt = 0.0
                chd = 0.0
                inf = 0.0
                qtd = 0.0
                valor_fee = 0.0
                valor_venda = 0.0
                voucher_recibo = 0.0

                if len(tokens) >= 1:
                    try:
                        adt = float(tokens[0].replace(",", "."))
                    except:
                        pass

                if len(tokens) >= 2:
                    try:
                        valor_fee = float(tokens[1].replace(",", "."))
                    except:
                        pass

                if len(tokens) >= 3:
                    try:
                        chd = float(tokens[2].replace(",", "."))
                    except:
                        pass

                if len(tokens) >= 4:
                    try:
                        inf = float(tokens[3].replace(",", "."))
                    except:
                        pass

                if len(tokens) >= 5:
                    try:
                        qtd = float(tokens[4].replace(",", "."))
                    except:
                        pass

                # Tratar valores de venda e voucher/tarifa
                if len(tokens) >= 6:
                    raw_val = tokens[5].replace(".", "").replace(",", ".")
                    # Se vierem dois valores grudados (ex: 496,00248,00)
                    vals = re.findall(r"\d+\.?\d*", raw_val)
                    if len(vals) >= 2:
                        valor_venda = float(vals[0])
                        voucher_recibo = float(vals[1])
                    elif len(vals) == 1:
                        valor_venda = float(vals[0])

                records.append({
                    "FILE": int(current_file)
                    if current_file.isdigit()
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
                    "TARIFA": None,  # Mantido nulo igual a sua planilha
                    "VALOR_VENDA": valor_venda,
                    "VALOR_FEE": valor_fee,
                })

    df = pd.DataFrame(records)

    # Adicionar a linha do TOTAL no final igual à sua planilha
    if not df.empty:
        total_venda = df["VALOR_VENDA"].sum()
        total_fee = df["VALOR_FEE"].sum()

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
