"""
Conversor de relatórios de vendas (PDF -> Excel)

Tipos suportados:
  - MANUTENÇÃO COMISSIONADA SITE        (1 PDF)
  - MANUTENÇÃO COMISSIONADA APP         (2 PDFs - quinzenas, concatenados)
  - MANUTENÇÃO COMISSIONADA MANAGETOUR  (1 PDF)

Estratégia de extração (pdfplumber):
  1. Localiza as tabelas da página e extrai cada linha com a posição vertical.
  2. Captura também o texto FORA das tabelas (ex.: "601059-LETICIA GOMES (BRT OPERADORA)")
     e o intercala na ordem original da página, para que nenhum nome seja perdido.
  3. Remove cabeçalhos/rodapés repetidos de quebra de página e o cabeçalho da tabela
     repetido em cada página.
  4. Se a página não tiver linhas de grade (tabela "sem bordas"), tenta a estratégia
     baseada em alinhamento de texto.
"""

import io
import re
import unicodedata

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------------------
# Configuração dos relatórios
# --------------------------------------------------------------------------------------

COLS_SERVICOS = [
    "SERVIÇO",
    "DATA SERVIÇO",
    "ADT",
    "CHD.",
    "INF",
    "QTD",
    "VOUCHER/RECIBO",
    "TARIFA",
    "VALOR VENDA",
    "CATEGORIA SERVIÇO",
    "VALOR FEE",
]

COLS_MANAGETOUR = [
    "Id",
    "File",
    "Total Geral (Soma)",
    "Receita Operacional (Soma)",
    "Custo Operacao Rateio (Soma)",
    "Total NET - Previsto (Soma)",
]

# Padrões (aplicados sobre o texto normalizado: sem acento, MAIÚSCULO) de linhas "sujas"
JUNK_COMUM = [
    r"PAGINA\s*\d+\s*(DE|/)\s*\d+",
]

JUNK_SERVICOS = JUNK_COMUM + [
    r"FEE\s*-\s*LISTA SERVICOS TRANSACIONADOS",
    r"BROCKER TURISMO",
    r"^ORIGEM\s*:",
]

JUNK_MANAGETOUR = JUNK_COMUM + [
    r"SUMARIO DE",
    r"^\d{1,2}/\d{1,2}/\d{4},?\s+\d{1,2}:\d{2}",  # data/hora de emissão
    r"HTTPS?://",
    r"WWW\.",
]

# Linhas que começam com código de cliente (ex.: "601059-NOME") nunca são descartadas
# como lixo, mesmo que contenham um termo de lixo no texto.
PROTECT_SERVICOS = r"^\d{3,}\s*-"

CONFIG = {
    "MANUTENÇÃO COMISSIONADA SITE": {
        "columns": COLS_SERVICOS,
        "multi": False,
        "numeric": ["ADT", "CHD.", "INF", "QTD", "TARIFA", "VALOR VENDA", "VALOR FEE"],
        "money": ["TARIFA", "VALOR VENDA", "VALOR FEE"],
        "junk": JUNK_SERVICOS,
        "protect": PROTECT_SERVICOS,
    },
    "MANUTENÇÃO COMISSIONADA APP": {
        "columns": COLS_SERVICOS,
        "multi": True,
        "numeric": ["ADT", "CHD.", "INF", "QTD", "TARIFA", "VALOR VENDA", "VALOR FEE"],
        "money": ["TARIFA", "VALOR VENDA", "VALOR FEE"],
        "junk": JUNK_SERVICOS,
        "protect": PROTECT_SERVICOS,
    },
    "MANUTENÇÃO COMISSIONADA MANAGETOUR": {
        "columns": COLS_MANAGETOUR,
        "multi": False,
        "numeric": [
            "Total Geral (Soma)",
            "Receita Operacional (Soma)",
            "Custo Operacao Rateio (Soma)",
            "Total NET - Previsto (Soma)",
        ],
        "money": [
            "Total Geral (Soma)",
            "Receita Operacional (Soma)",
            "Custo Operacao Rateio (Soma)",
            "Total NET - Previsto (Soma)",
        ],
        "junk": JUNK_MANAGETOUR,
        "protect": None,
    },
}

# --------------------------------------------------------------------------------------
# Utilitários de texto
# --------------------------------------------------------------------------------------


def norm(text: str) -> str:
    """Remove acentos, coloca em maiúsculo e colapsa espaços."""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().upper()


def clean_cell(value) -> str:
    """None -> '', quebras de linha -> espaço, espaços colapsados."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


_INT_BR = re.compile(r"^-?\d+(?:,\d+)?$")
_THOUSANDS_BR = re.compile(r"^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$")


def parse_br_number(value):
    """
    Converte texto em formato brasileiro ('1.234,56') para número.
    Se não for um número reconhecível, devolve o valor original (nada se perde).
    """
    if not isinstance(value, str):
        return value
    s = value.strip().replace("R$", "").replace(" ", "")
    if not s:
        return value
    if _THOUSANDS_BR.match(s) or _INT_BR.match(s):
        has_decimal = "," in s
        num = float(s.replace(".", "").replace(",", "."))
        return num if has_decimal else int(num)
    return value


def looks_numeric(text: str) -> bool:
    return not isinstance(parse_br_number(text), str)


def token_count(text_norm: str, names_norm: list[str]) -> int:
    """Quantos nomes de coluna aparecem como 'palavras inteiras' no texto."""
    count = 0
    for n in names_norm:
        if re.search(rf"(?<![A-Z0-9]){re.escape(n)}(?![A-Z0-9])", text_norm):
            count += 1
    return count


# --------------------------------------------------------------------------------------
# Extração
# --------------------------------------------------------------------------------------


def is_junk(text: str, cfg: dict, extra_terms: list[str]) -> bool:
    t = norm(text)
    if not t:
        return True
    for term in extra_terms:
        if term and term in t:
            return True
    if cfg["protect"] and re.match(cfg["protect"], t):
        return False
    return any(re.search(p, t) for p in cfg["junk"])


def is_header(cells: list[str], text: str, cfg: dict) -> bool:
    """Detecta a linha de cabeçalho da tabela (repetida a cada página)."""
    names_norm = [norm(c) for c in cfg["columns"]]
    cells_norm = {norm(c) for c in cells if c}
    hits_cells = sum(1 for n in names_norm if n in cells_norm)
    if hits_cells >= min(3, len(names_norm)):
        return True
    return token_count(norm(text), names_norm) >= min(4, len(names_norm))


def extract_page_items(page, strategy: str):
    """
    Retorna lista de (posição_vertical, tipo, conteúdo):
      tipo 'table' -> lista de células de uma linha de tabela
      tipo 'text'  -> linha de texto fora de qualquer tabela
    """
    if strategy == "lines":
        settings = {}
    else:
        settings = {"vertical_strategy": "text", "horizontal_strategy": "text"}

    tables = page.find_tables(table_settings=settings)
    bboxes = [t.bbox for t in tables]
    items = []

    for table in tables:
        data = table.extract()
        for row_obj, row in zip(table.rows, data):
            items.append((row_obj.bbox[1], "table", row))

    def outside_tables(obj):
        if obj.get("object_type") != "char":
            return True
        cx = (obj["x0"] + obj["x1"]) / 2
        cy = (obj["top"] + obj["bottom"]) / 2
        return not any(b[0] <= cx <= b[2] and b[1] <= cy <= b[3] for b in bboxes)

    rest = page.filter(outside_tables) if bboxes else page
    for line in rest.extract_text_lines():
        items.append((line["top"], "text", line["text"]))

    return items


def has_valid_table(items, ncols: int) -> bool:
    return any(kind == "table" and len(payload) >= ncols for _, kind, payload in items)


def extract_pdf_rows(file_bytes: bytes, cfg: dict, extra_terms: list[str], log: list[str], label: str):
    """Extrai todas as linhas úteis de um PDF, já limpas, na ordem original."""
    ncols = len(cfg["columns"])
    rows: list[list[str]] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError("O PDF não possui páginas.")

        for page_no, page in enumerate(pdf.pages, start=1):
            chosen = None
            last = None
            for strategy in ("lines", "text"):
                try:
                    cand = extract_page_items(page, strategy)
                except Exception:
                    continue
                last = cand
                if has_valid_table(cand, ncols):
                    chosen = cand
                    break

            if chosen is None:
                chosen = last or []
                if chosen:
                    log.append(
                        f"{label} – página {page_no}: estrutura de colunas não identificada com certeza; "
                        "conteúdo mantido como texto. Confira esta página."
                    )

            for _, kind, payload in sorted(chosen, key=lambda x: x[0]):
                if kind == "table":
                    cells = [clean_cell(c) for c in payload]
                    nonempty = [c for c in cells if c]
                    if not nonempty:
                        continue
                    text = " ".join(nonempty)
                    if is_header(cells, text, cfg) or is_junk(text, cfg, extra_terms):
                        continue
                    # Linha de nome: só 1 célula preenchida e não numérica -> vai para a 1ª coluna
                    if len(nonempty) == 1 and not looks_numeric(nonempty[0]):
                        cells = [nonempty[0]] + [""] * (len(cells) - 1)
                    rows.append(cells)
                else:
                    text = clean_cell(payload)
                    if not text:
                        continue
                    if token_count(norm(text), [norm(c) for c in cfg["columns"]]) >= min(
                        4, ncols
                    ) or is_junk(text, cfg, extra_terms):
                        continue
                    rows.append([text])

    return rows


def rows_to_dataframe(rows: list[list[str]], cfg: dict, log: list[str], label: str) -> pd.DataFrame:
    if not rows:
        raise ValueError("Nenhuma linha de dados foi encontrada no PDF.")

    expected = cfg["columns"]
    width = max(len(expected), max(len(r) for r in rows))
    padded = [r + [""] * (width - len(r)) for r in rows]
    names = expected + [f"COLUNA_EXTRA_{i}" for i in range(1, width - len(expected) + 1)]
    df = pd.DataFrame(padded, columns=names)

    # Colunas extras: só mantém se tiverem algum conteúdo (nada se perde)
    for col in names[len(expected):]:
        if (df[col] == "").all():
            df = df.drop(columns=col)
        else:
            log.append(
                f"{label}: foi detectada uma coluna além das esperadas ({col}) com conteúdo; ela foi mantida."
            )
    return df


def process_pdf(uploaded_file, cfg: dict, extra_terms: list[str], log: list[str]) -> pd.DataFrame:
    label = uploaded_file.name
    rows = extract_pdf_rows(uploaded_file.getvalue(), cfg, extra_terms, log, label)
    return rows_to_dataframe(rows, cfg, log, label)


# --------------------------------------------------------------------------------------
# Exportação para Excel
# --------------------------------------------------------------------------------------


def to_excel_bytes(df: pd.DataFrame, cfg: dict) -> bytes:
    out = df.astype(object).where(df != "", None)  # vazios viram células vazias
    buffer = io.BytesIO()
    sheet = "Relatorio"

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"

        money_idx = {i for i, c in enumerate(out.columns, start=1) if c in cfg["money"]}
        for i, col in enumerate(out.columns, start=1):
            letter = get_column_letter(i)
            max_len = len(str(col))
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                cell = row[0]
                if cell.value is None:
                    continue
                if i in money_idx and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 70)

    return buffer.getvalue()


def make_filename(tipo: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", norm(tipo)).strip("_") + ".xlsx"


# --------------------------------------------------------------------------------------
# Interface Streamlit
# --------------------------------------------------------------------------------------

st.set_page_config(page_title="Conversor PDF → Excel", page_icon="📊", layout="wide")
st.title("📊 Conversor de Relatórios PDF → Excel")

tipo = st.selectbox("Tipo de relatório", list(CONFIG.keys()))
cfg = CONFIG[tipo]

if cfg["multi"]:
    st.caption("Envie os **2 arquivos** (1ª e 2ª quinzena). Eles serão unidos em uma única planilha.")
else:
    st.caption("Envie **1 arquivo** PDF.")

uploaded = st.file_uploader(
    "Selecione o(s) arquivo(s) PDF",
    type=["pdf"],
    accept_multiple_files=cfg["multi"],
    key=f"uploader_{tipo}",
)

with st.expander("Opções avançadas"):
    convert_numbers = st.checkbox(
        "Converter valores numéricos (formato brasileiro 1.234,56) para número no Excel",
        value=True,
    )
    extra_text = st.text_area(
        "Termos extras a ignorar (um por linha). Linhas que contenham o termo serão descartadas.",
        value="",
        height=90,
    )

extra_terms = [norm(t) for t in extra_text.splitlines() if t.strip()]

if st.button("Converter para Excel", type="primary"):
    files = uploaded if isinstance(uploaded, list) else ([uploaded] if uploaded else [])
    st.session_state.pop("result", None)

    if not files:
        st.error("Envie ao menos um arquivo PDF antes de converter.")
    else:
        if cfg["multi"] and len(files) != 2:
            st.warning(f"Esperado 2 arquivos para o relatório APP; recebidos {len(files)}. Processando mesmo assim.")

        files = sorted(files, key=lambda f: f.name)  # mantém a ordem 1ª → 2ª quinzena pelo nome
        log: list[str] = []
        frames = []
        failed = False

        with st.spinner("Lendo PDF(s)..."):
            for f in files:
                try:
                    frames.append(process_pdf(f, cfg, extra_terms, log))
                except Exception as exc:  # noqa: BLE001
                    failed = True
                    st.error(f"Falha ao ler **{f.name}**: {exc}")

        if not failed:
            try:
                df = pd.concat(frames, ignore_index=True)  # cabeçalho aparece uma única vez

                if convert_numbers:
                    for col in cfg["numeric"]:
                        if col in df.columns:
                            df[col] = df[col].map(parse_br_number)

                excel_bytes = to_excel_bytes(df, cfg)
                preview = df.fillna("").astype(str)

                st.session_state["result"] = {
                    "tipo": tipo,
                    "bytes": excel_bytes,
                    "filename": make_filename(tipo),
                    "preview": preview,
                    "rows": len(df),
                    "files": [f.name for f in files],
                    "log": log,
                }
            except Exception as exc:  # noqa: BLE001
                st.error(f"Falha ao gerar o Excel: {exc}")

result = st.session_state.get("result")
if result and result["tipo"] == tipo:
    st.success(f"Conversão concluída: {result['rows']} linhas (arquivos: {', '.join(result['files'])}).")
    for msg in result["log"]:
        st.warning(msg)

    st.download_button(
        label="⬇️ Baixar Excel (.xlsx)",
        data=result["bytes"],
        file_name=result["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("Pré-visualização")
    st.dataframe(result["preview"], use_container_width=True, height=450)
