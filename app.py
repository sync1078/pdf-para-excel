"""
Conversor de relatórios de vendas (PDF -> Excel) – saída ACHATADA (flatten)

Arquitetura (sem nenhum método de tabela do pdfplumber):
  1. `page.extract_text(layout=True)` preserva os espaços originais da página.
  2. O texto é dividido com `split('\\n')` e varrido linha a linha, com ESTADO em memória:
        current_file / current_cliente / current_origem
     que é copiado para cada linha de serviço seguinte (cliente -> serviços).
  3. Linha de cliente  : Regex  ^(\\d{5,})\\s*-\\s*(nome)\\s*\\((origem)\\)
     Linha de serviço  : contém data dd/mm/aa. As colunas são separadas por 2+ espaços
                         (re.split(r'\\s{2,}')); o nome do serviço fica antes da data e o
                         bloco da direita (ADT CHD INF QTD [VOUCHER] TARIFA VENDA CATEGORIA FEE)
                         é lido por gramática fixa, então VOUCHER vazio nunca desloca valores.
  4. Os registros viram uma lista de dicionários -> um único DataFrame com colunas FIXAS.
"""

import datetime as dt
import io
import re
import unicodedata

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------------------
# Colunas de saída (ordem exata)
# --------------------------------------------------------------------------------------

COLS_SERVICOS = [
    "FILE",
    "NOME_CLIENTE",
    "SITE_ORIGEM",
    "DATA_SERVICO",
    "SERVICO",
    "CATEGORIA_SERVICO",
    "ADT",
    "CHD",
    "INF",
    "QTD",
    "VOUCHER_RECIBO",
    "TARIFA",
    "VALOR_VENDA",
    "VALOR_FEE",
]
MONEY_SERVICOS = ["TARIFA", "VALOR_VENDA", "VALOR_FEE"]

COLS_MANAGETOUR = [
    "Id",
    "File",
    "Total Geral (Soma)",
    "Receita Operacional (Soma)",
    "Custo Operacao Rateio (Soma)",
    "Total NET - Previsto (Soma)",
]
MONEY_MANAGETOUR = COLS_MANAGETOUR[2:]

REPORTS = {
    "MANUTENÇÃO COMISSIONADA SITE": {"kind": "servicos", "multi": False},
    "MANUTENÇÃO COMISSIONADA APP": {"kind": "servicos", "multi": True},
    "MANUTENÇÃO COMISSIONADA MANAGETOUR": {"kind": "managetour", "multi": False},
}


class ParseError(ValueError):
    """Erro explícito e esperado de leitura/validação (mostrado ao usuário)."""

    raw = None  # texto bruto da página 1, para diagnóstico


# --------------------------------------------------------------------------------------
# Utilitários
# --------------------------------------------------------------------------------------


def norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().upper()


def collapse(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip()


MONEY = r"-?\d[\d.]*,\d{2}"
NUM_TOKEN = r"-?\d[\d.]*(?:,\d+)?"

_BR_THOUSANDS = re.compile(r"^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$")
_BR_PLAIN = re.compile(r"^-?\d+(?:,\d+)?$")


def br_to_number(token):
    """'1.234,56' -> 1234.56 | '2' -> 2. Texto não numérico é devolvido sem alteração."""
    if not isinstance(token, str):
        return token
    s = token.strip()
    if _BR_THOUSANDS.match(s) or _BR_PLAIN.match(s):
        value = float(s.replace(".", "").replace(",", "."))
        return value if "," in s else int(value)
    return token


def to_float(token):
    value = br_to_number(token)
    return float(value) if isinstance(value, (int, float)) else value


def to_date(text: str):
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return text  # mantém o texto se não for uma data válida


# --------------------------------------------------------------------------------------
# Leitura do PDF: texto bruto com layout preservado
# --------------------------------------------------------------------------------------


def extract_pages_text(file_bytes: bytes, filename: str) -> list[str]:
    try:
        pdf = pdfplumber.open(io.BytesIO(file_bytes))
    except Exception as exc:  # noqa: BLE001 - convertido em erro explícito para o usuário
        raise ParseError(f"{filename}: não foi possível abrir o arquivo como PDF ({type(exc).__name__}).") from exc

    with pdf:
        if not pdf.pages:
            raise ParseError(f"{filename}: o PDF não possui páginas.")
        pages = [page.extract_text(layout=True) or "" for page in pdf.pages]

    if not any(p.strip() for p in pages):
        raise ParseError(f"{filename}: o PDF não tem texto selecionável (pode ser imagem escaneada).")
    return pages


# --------------------------------------------------------------------------------------
# APP / SITE
# --------------------------------------------------------------------------------------

JUNK_SERVICOS = [
    r"EMISSAO\s*:",
    r"FEE\s*-\s*LISTA",
    r"LISTA SERVICOS TRANSACIONADOS",
    r"DATA SERVICO",
    r"PAGINA\s*\d+\s*(DE|/)\s*\d+",
    r"^INICIO SERVICO",
    r"^FIM SERVICO",
    r"^ORIGEM\s*:",
    r"^SERVICO\s*:",
    r"BROCKER TURISMO",
    r"^(\d{2}/\d{2}/\d{2,4}\s*)+$",
]
HEADER_WORDS = {
    "SERVICO", "DATA", "ADT", "CHD", "CHD.", "INF", "QTD", "VOUCHER/RECIBO", "VOUCHER",
    "RECIBO", "TARIFA", "VALOR", "VENDA", "CATEGORIA", "FEE", "%",
}

# Cliente: "743133 - RODRIGO ALMEIDA (SITE BROCKER) 2 1 0 3 127,00 1,27" -> file, nome, origem
# (depois do ')' só podem existir números = totais do cliente, que não entram no flatten)
CLIENT_START_RE = re.compile(r"^\d{5,}\s*-\s*\S")
CLIENT_RE = re.compile(
    rf"^(?P<file>\d{{5,}})\s*-\s*(?P<nome>.*?)\s*\((?P<origem>[^()]*)\)(?:\s+{NUM_TOKEN})*\s*$"
)
CLIENT_PARTIAL_RE = re.compile(r"^(?P<file>\d{5,})\s*-\s*(?P<nome>.*)$")

DATE_RE = re.compile(r"(?<!\d)\d{2}/\d{2}/(?:\d{4}|\d{2})(?!\d)")

# Bloco à direita da data: ADT CHD INF QTD [VOUCHER] TARIFA VENDA CATEGORIA [x,xx %] FEE
SERVICE_RIGHT_RE = re.compile(
    rf"""^
    (?P<adt>\d+)\s+(?P<chd>\d+)\s+(?P<inf>\d+)\s+(?P<qtd>\d+)\s+
    (?:(?P<voucher>(?!{MONEY}(?:\s|$))\S+)\s+)?
    (?P<tarifa>{MONEY})\s+
    (?P<venda>{MONEY})\s+
    (?P<cat>.*?)\s*
    (?:\d+(?:,\d+)?\s*%\s+)?
    (?P<fee>{MONEY})
    \s*$""",
    re.VERBOSE,
)
NUMERIC_ONLY_RE = re.compile(rf"^{NUM_TOKEN}(?:\s+{NUM_TOKEN})*$")


def parse_service_line(raw: str):
    """
    Colunas separadas por 2+ espaços. Retorna (nome_servico, data, match_direita) ou None.
    """
    cols = re.split(r"\s{2,}", raw.strip())
    date_idx = next((i for i, c in enumerate(cols) if re.fullmatch(DATE_RE, c)), None)
    if date_idx is not None:
        nome = " ".join(cols[:date_idx]).strip()
        right = " ".join(" ".join(cols[date_idx + 1:]).split())
        m = SERVICE_RIGHT_RE.match(right)
        if m:
            return nome, cols[date_idx], m

    # Data colada ao texto por um único espaço: procura a data dentro da linha
    line = collapse(raw)
    for dm in DATE_RE.finditer(line):
        m = SERVICE_RIGHT_RE.match(line[dm.end():].strip())
        if m:
            return line[: dm.start()].strip(), dm.group(), m
    return None


def parse_servicos(pages_text: list[str], label: str, extra_terms: list[str]):
    """
    Retorna (records, avisos, conferencia).
    conferencia = (ok: bool | None, texto) comparando com o TOTAL RELATÓRIO do PDF.
    """
    records: list[dict] = []
    avisos: list[str] = []
    total_pdf = None

    # Estado em memória
    current_file = current_cliente = current_origem = None
    pending_client = None  # cliente cujo nome quebrou em mais de uma linha
    boundary = 0  # índice do 1º serviço do cliente atual
    buffer: list[str] = []  # linhas de texto solto (nome de serviço quebrado)

    def new_record(nome: str) -> dict:
        rec = {c: None for c in COLS_SERVICOS}
        rec["FILE"] = current_file
        rec["NOME_CLIENTE"] = current_cliente
        rec["SITE_ORIGEM"] = current_origem
        rec["SERVICO"] = nome
        rec["_empty_prefix"] = not nome
        return rec

    def flush(next_rec):
        """Atribui as linhas soltas ao serviço correto (acima, abaixo ou divididas)."""
        if not buffer:
            return
        lines = list(buffer)
        buffer.clear()
        text = " ".join(lines)
        prev = records[-1] if len(records) > boundary else None

        if next_rec is not None and not next_rec["SERVICO"]:
            if prev is not None and prev["_empty_prefix"]:
                k = len(lines) // 2  # nome centralizado: metade acima, metade abaixo
                prev["SERVICO"] = f"{prev['SERVICO']} {' '.join(lines[:k])}".strip()
                next_rec["SERVICO"] = " ".join(lines[k:])
            else:
                next_rec["SERVICO"] = text
        elif prev is not None:
            prev["SERVICO"] = f"{prev['SERVICO']} {text}".strip()
        elif next_rec is not None:
            next_rec["SERVICO"] = f"{text} {next_rec['SERVICO']}".strip()
        else:
            avisos.append(f"{label}: texto solto sem serviço associado: '{text}'")

    def set_client(m):
        nonlocal current_file, current_cliente, current_origem
        current_file = int(m.group("file"))
        current_cliente = m.group("nome").strip()
        current_origem = m.group("origem").strip()

    def close_pending_without_origin():
        nonlocal current_file, current_cliente, current_origem, pending_client
        pm = CLIENT_PARTIAL_RE.match(pending_client)
        current_file, current_cliente, current_origem = int(pm.group("file")), pm.group("nome").strip(), ""
        avisos.append(f"{label}: cliente sem '(origem)' reconhecível: '{pending_client}'")
        pending_client = None

    for page_no, page_text in enumerate(pages_text, start=1):
        for raw in page_text.split("\n"):
            line = collapse(raw)
            if not line:
                continue
            n = norm(line)
            client_start = bool(CLIENT_START_RE.match(line))

            # 1) Lixo de quebra de página (linhas de cliente são protegidas)
            if not client_start:
                if any(re.search(p, n) for p in JUNK_SERVICOS) or any(t in n for t in extra_terms):
                    continue
                if all(t in HEADER_WORDS for t in n.split()):
                    continue

            # 2) Novo cliente: atualiza o estado
            if client_start:
                flush(None)
                if pending_client:
                    close_pending_without_origin()
                boundary = len(records)
                m = CLIENT_RE.match(line)
                if m:
                    set_client(m)
                else:
                    pending_client = line  # nome quebrado: completa nas próximas linhas
                continue

            # 3) Linha de serviço
            svc = parse_service_line(raw)
            if svc:
                if pending_client:
                    close_pending_without_origin()
                nome, data, m = svc
                rec = new_record(nome)
                rec["DATA_SERVICO"] = data
                rec["ADT"], rec["CHD"], rec["INF"], rec["QTD"] = (
                    int(m.group("adt")), int(m.group("chd")), int(m.group("inf")), int(m.group("qtd"))
                )
                rec["VOUCHER_RECIBO"] = m.group("voucher")  # costuma vir vazio
                rec["TARIFA"] = to_float(m.group("tarifa"))
                rec["VALOR_VENDA"] = to_float(m.group("venda"))
                rec["CATEGORIA_SERVICO"] = m.group("cat").strip() or None
                rec["VALOR_FEE"] = to_float(m.group("fee"))
                if current_file is None:
                    avisos.append(f"{label} (pág. {page_no}): serviço antes de qualquer cliente: '{line}'")
                flush(rec)
                records.append(rec)
                continue

            # 4) TOTAL RELATÓRIO (usado só para conferência) e demais totais (ignorados)
            if n.startswith("TOTAL"):
                flush(None)
                if n.startswith("TOTAL RELATORIO"):
                    tokens = line.split()
                    nums = [t for t in tokens if re.fullmatch(NUM_TOKEN, t)]
                    if len(nums) >= 6:
                        total_pdf = {
                            "ADT": br_to_number(nums[-6]), "CHD": br_to_number(nums[-5]),
                            "INF": br_to_number(nums[-4]), "QTD": br_to_number(nums[-3]),
                            "VALOR_VENDA": to_float(nums[-2]), "VALOR_FEE": to_float(nums[-1]),
                        }
                continue

            # 5) Continuação do nome de um cliente quebrado em várias linhas
            if pending_client:
                combined = f"{pending_client} {line}"
                m = CLIENT_RE.match(combined)
                if m:
                    set_client(m)
                    pending_client = None
                else:
                    pending_client = combined
                continue

            # 6) Números soltos sem identificação (nada a anexar com segurança)
            if NUMERIC_ONLY_RE.match(line):
                avisos.append(f"{label} (pág. {page_no}): linha numérica sem identificação: '{line}'")
                continue

            # 7) Texto solto: nome de serviço que quebrou de linha
            buffer.append(line)

    flush(None)
    if pending_client:
        close_pending_without_origin()

    if not records:
        raise ParseError(
            f"{label}: nenhuma linha de serviço (com data dd/mm/aa) foi reconhecida. "
            "Veja o texto bruto em 'Diagnóstico'."
        )

    return records, avisos, reconcile(records, total_pdf, label)


def reconcile(records: list[dict], total_pdf, label: str):
    """Compara a soma das linhas extraídas com o TOTAL RELATÓRIO impresso no PDF."""
    if total_pdf is None:
        return None, f"{label}: linha 'TOTAL RELATÓRIO' não encontrada; conferência de totais não realizada."
    diffs = []
    for col in ("ADT", "CHD", "INF", "QTD"):
        got = sum(r[col] or 0 for r in records)
        if got != total_pdf[col]:
            diffs.append(f"{col}: extraído {got} × PDF {total_pdf[col]}")
    for col in ("VALOR_VENDA", "VALOR_FEE"):
        got = round(sum(r[col] or 0 for r in records), 2)
        if abs(got - total_pdf[col]) > 0.05:
            diffs.append(f"{col}: extraído {got:,.2f} × PDF {total_pdf[col]:,.2f}")
    if diffs:
        return False, f"{label}: conferência com TOTAL RELATÓRIO divergente → " + "; ".join(diffs)
    return True, f"{label}: conferência com TOTAL RELATÓRIO OK (ADT/CHD/INF/QTD/VENDA/FEE batem)."


# --------------------------------------------------------------------------------------
# MANAGETOUR
# --------------------------------------------------------------------------------------

JUNK_MANAGETOUR = [
    r"SUMARIO DE",
    r"^\d{1,2}/\d{1,2}/\d{4},?\s+\d{1,2}:\d{2}",
    r"PAGINA\s*\d+\s*(DE|/)\s*\d+",
    r"HTTPS?://",
    r"WWW\.",
]
HEADER_MT = [norm(c) for c in COLS_MANAGETOUR]

# Ex.: "1 615.128 398,90 398,90 0,00 285,23"
MT_RE = re.compile(
    rf"^(?P<id>\d+)\s+(?P<file>\S+)\s+(?P<a>{MONEY})\s+(?P<b>{MONEY})\s+(?P<c>{MONEY})\s+(?P<d>{MONEY})\s*$"
)


def parse_managetour(pages_text: list[str], label: str, extra_terms: list[str]):
    records: list[dict] = []
    avisos: list[str] = []

    for page_no, page_text in enumerate(pages_text, start=1):
        for raw in page_text.split("\n"):
            line = collapse(raw)
            if not line:
                continue
            m = MT_RE.match(line)
            if m:
                records.append({
                    "Id": int(m.group("id")),
                    "File": m.group("file"),  # texto: preserva '615.128' exatamente
                    "Total Geral (Soma)": to_float(m.group("a")),
                    "Receita Operacional (Soma)": to_float(m.group("b")),
                    "Custo Operacao Rateio (Soma)": to_float(m.group("c")),
                    "Total NET - Previsto (Soma)": to_float(m.group("d")),
                })
                continue

            n = norm(line)
            if any(re.search(p, n) for p in JUNK_MANAGETOUR) or any(t in n for t in extra_terms):
                continue
            if sum(1 for h in HEADER_MT if h in n) >= 3:
                continue
            if len(re.findall(MONEY, line)) >= 2:  # parece dado, mas não bateu com o padrão
                avisos.append(f"{label} (pág. {page_no}): linha com valores não reconhecida: '{line}'")

    if not records:
        raise ParseError(
            f"{label}: nenhuma linha transacional (Id File + 4 valores) foi reconhecida. "
            "Veja o texto bruto em 'Diagnóstico'."
        )
    return records, avisos


# --------------------------------------------------------------------------------------
# DataFrame e Excel
# --------------------------------------------------------------------------------------


def records_to_df(records: list[dict], columns: list[str]) -> pd.DataFrame:
    rows = [{c: r.get(c) for c in columns} for r in records]  # chaves internas ficam de fora
    return pd.DataFrame(rows, columns=columns, dtype=object)


def to_excel_bytes(df: pd.DataFrame, money_cols: list[str]) -> bytes:
    out = df.copy()
    if "DATA_SERVICO" in out.columns:
        out["DATA_SERVICO"] = out["DATA_SERVICO"].map(lambda v: to_date(v) if isinstance(v, str) else v)

    buffer = io.BytesIO()
    sheet = "Relatorio"
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"

        for i, col in enumerate(out.columns, start=1):
            max_len = len(str(col))
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                cell = row[0]
                if cell.value is None:
                    continue
                if col in money_cols and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
                elif col == "DATA_SERVICO" and isinstance(cell.value, (dt.date, dt.datetime)):
                    cell.number_format = "DD/MM/YYYY"
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(i)].width = min(max_len + 2, 70)
    return buffer.getvalue()


def make_filename(tipo: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", norm(tipo)).strip("_") + ".xlsx"


# --------------------------------------------------------------------------------------
# Interface Streamlit
# --------------------------------------------------------------------------------------


def run_conversion(tipo: str, files: list, extra_terms: list[str]) -> dict:
    """Varre todos os PDFs numa única lista de registros e monta o DataFrame final."""
    spec = REPORTS[tipo]
    all_records: list[dict] = []
    avisos: list[str] = []
    conferencias: list[tuple] = []
    raw_samples: dict[str, str] = {}

    for f in sorted(files, key=lambda x: x.name):  # APP: 1ª quinzena → 2ª pelo nome do arquivo
        pages = extract_pages_text(f.getvalue(), f.name)
        raw_samples[f.name] = pages[0]
        try:
            if spec["kind"] == "servicos":
                records, warn, conf = parse_servicos(pages, f.name, extra_terms)
                conferencias.append(conf)
            else:
                records, warn = parse_managetour(pages, f.name, extra_terms)
        except ParseError as exc:
            exc.raw = pages[0]
            raise
        all_records.extend(records)
        avisos.extend(warn)

    if spec["kind"] == "servicos":
        cols, money = COLS_SERVICOS, MONEY_SERVICOS
    else:
        cols, money = COLS_MANAGETOUR, MONEY_MANAGETOUR

    df = records_to_df(all_records, cols)  # um único DataFrame, cabeçalho uma só vez
    return {
        "tipo": tipo,
        "bytes": to_excel_bytes(df, money),
        "filename": make_filename(tipo),
        "preview": df.fillna("").astype(str),
        "rows": len(df),
        "files": [f.name for f in sorted(files, key=lambda x: x.name)],
        "avisos": avisos,
        "conferencias": conferencias,
        "raw": raw_samples,
    }


def main():
    st.set_page_config(page_title="Conversor PDF → Excel", page_icon="📊", layout="wide")
    st.title("📊 Conversor de Relatórios PDF → Excel")

    tipo = st.selectbox("Tipo de relatório", list(REPORTS.keys()))
    spec = REPORTS[tipo]

    if spec["multi"]:
        st.caption("Envie os **2 arquivos** (1ª e 2ª quinzena). Eles serão unidos em uma única planilha.")
    else:
        st.caption("Envie **1 arquivo** PDF.")

    uploaded = st.file_uploader(
        "Selecione o(s) arquivo(s) PDF",
        type=["pdf"],
        accept_multiple_files=spec["multi"],
        key=f"uploader_{tipo}",
    )

    with st.expander("Opções avançadas"):
        extra_text = st.text_area(
            "Termos extras a ignorar (um por linha). Linhas que contenham o termo são descartadas.",
            value="",
            height=90,
        )
    extra_terms = [norm(t) for t in extra_text.splitlines() if t.strip()]

    if st.button("Converter para Excel", type="primary"):
        st.session_state.pop("result", None)
        files = uploaded if isinstance(uploaded, list) else ([uploaded] if uploaded else [])

        if not files:
            st.error("Envie ao menos um arquivo PDF antes de converter.")
        elif spec["multi"] and len(files) != 2:
            st.error(f"O relatório APP exige exatamente 2 arquivos (quinzenas); foram enviados {len(files)}.")
        else:
            try:
                with st.spinner("Lendo PDF(s)..."):
                    st.session_state["result"] = run_conversion(tipo, files, extra_terms)
            except ParseError as exc:
                st.error(str(exc))
                if exc.raw:
                    with st.expander("Diagnóstico – texto bruto da página 1"):
                        st.code(exc.raw, language=None)

    result = st.session_state.get("result")
    if result and result["tipo"] == tipo:
        st.success(f"Conversão concluída: {result['rows']} linhas (arquivos: {', '.join(result['files'])}).")

        for ok, text in result["conferencias"]:
            if ok is True:
                st.info("✅ " + text)
            else:
                st.warning("⚠️ " + text)
        if result["avisos"]:
            with st.expander(f"⚠️ {len(result['avisos'])} aviso(s) de linhas não encaixadas", expanded=True):
                for w in result["avisos"][:200]:
                    st.write("• " + w)

        st.download_button(
            label="⬇️ Baixar Excel (.xlsx)",
            data=result["bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Pré-visualização")
        st.dataframe(result["preview"], use_container_width=True, height=450)

        with st.expander("Diagnóstico – texto bruto da página 1 (como o pdfplumber enxergou)"):
            for name, txt in result["raw"].items():
                st.caption(name)
                st.code(txt or "(vazio)", language=None)


if __name__ == "__main__":
    main()
