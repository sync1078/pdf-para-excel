"""
Conversor de relatórios de vendas (PDF -> Excel)

Tipos suportados:
  - MANUTENÇÃO COMISSIONADA SITE        (1 PDF)
  - MANUTENÇÃO COMISSIONADA APP         (2 PDFs - quinzenas, concatenados)
  - MANUTENÇÃO COMISSIONADA MANAGETOUR  (1 PDF)

Estratégia: NÃO usa detecção de tabelas. Cada página é lida como texto bruto
(`page.extract_text(layout=True)`), dividida em linhas, e cada linha é classificada por
Regex (cliente, serviço, total, lixo). Os registros viram uma lista de dicionários e só
no final são convertidos em DataFrame com colunas FIXAS – portanto nunca surgem colunas
extras nem deslocamento de dados.
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
# Colunas fixas
# --------------------------------------------------------------------------------------

# "% FEE" guarda o percentual que aparece entre a categoria e o VALOR FEE (ex.: "1,00 %").
# Se não quiser essa coluna no Excel, basta remover "% FEE" da lista abaixo.
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
    "% FEE",
    "VALOR FEE",
]
MONEY_SERVICOS = ["TARIFA", "VALOR VENDA", "VALOR FEE"]

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

# --------------------------------------------------------------------------------------
# Utilitários
# --------------------------------------------------------------------------------------


def norm(text: str) -> str:
    """Sem acentos, MAIÚSCULO, espaços colapsados."""
    text = unicodedata.normalize("NFKD", str(text))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().upper()


def collapse(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip()


MONEY = r"-?\d[\d.]*,\d{2}"
NUM_TOKEN = r"-?\d[\d.]*(?:,\d+)?"

_BR_THOUSANDS = re.compile(r"^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$")
_BR_PLAIN = re.compile(r"^-?\d+(?:,\d+)?$")


def parse_br_number(token):
    """'1.234,56' -> 1234.56 | '2' -> 2. Se não for número, devolve o texto original."""
    if not isinstance(token, str):
        return token
    s = token.strip()
    if _BR_THOUSANDS.match(s) or _BR_PLAIN.match(s):
        value = float(s.replace(".", "").replace(",", "."))
        return value if "," in s else int(value)
    return token


# --------------------------------------------------------------------------------------
# Leitura do PDF (texto bruto, preservando espaços)
# --------------------------------------------------------------------------------------


def extract_pages_text(file_bytes: bytes) -> list[str]:
    pages = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if not pdf.pages:
            raise ValueError("O PDF não possui páginas.")
        for page in pdf.pages:
            pages.append(page.extract_text(layout=True) or "")
    return pages


# --------------------------------------------------------------------------------------
# Relatórios APP / SITE
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
    r"^(\d{2}/\d{2}/\d{2,4}\s*)+$",  # linha só com datas do cabeçalho (período)
]

HEADER_WORDS = {
    "SERVICO", "DATA", "ADT", "CHD", "CHD.", "INF", "QTD", "VOUCHER/RECIBO", "VOUCHER",
    "RECIBO", "TARIFA", "VALOR", "VENDA", "CATEGORIA", "FEE", "%",
}

# Linha de cliente/agência: "743133 - RODRIGO ALMEIDA (SITE BROCKER) 2 1 0 3 127,00 1,27"
CLIENT_RE = re.compile(
    rf"^(?P<code>\d{{3,}})\s*-\s*(?P<name>.+?)(?P<tail>(?:\s+{NUM_TOKEN})+)\s*$"
)
# Cliente cujo nome não traz números na mesma linha
CLIENT_NAME_RE = re.compile(r"^\d{3,}\s*-\s*\S")

DATE_RE = re.compile(r"(?<!\d)\d{2}/\d{2}/(?:\d{4}|\d{2})(?!\d)")

# O que vem DEPOIS da data numa linha de serviço
SERVICE_REST_RE = re.compile(
    rf"""^
    (?P<adt>\d+)\s+(?P<chd>\d+)\s+(?P<inf>\d+)\s+(?P<qtd>\d+)\s+
    (?:(?P<voucher>(?!{MONEY}(?:\s|$))\S+)\s+)?
    (?P<tarifa>{MONEY})\s+
    (?P<venda>{MONEY})\s+
    (?P<cat>.*?)\s*
    (?:(?P<pct>\d+(?:,\d+)?\s*%)\s+)?
    (?P<fee>{MONEY})
    \s*$""",
    re.VERBOSE,
)

NUMERIC_ONLY_RE = re.compile(rf"^(?:{NUM_TOKEN})(?:\s+{NUM_TOKEN})*$")


def map_tail(tokens: list[str]):
    """Mapeia os números do fim de linhas de cliente/total para as colunas."""
    n = len(tokens)
    if n < 6 or n > 8:
        return None
    d = {
        "ADT": tokens[0],
        "CHD.": tokens[1],
        "INF": tokens[2],
        "QTD": tokens[3],
        "VALOR VENDA": tokens[-2],
        "VALOR FEE": tokens[-1],
    }
    if n == 7:
        d["TARIFA"] = tokens[4]
    elif n == 8:
        d["VOUCHER/RECIBO"], d["TARIFA"] = tokens[4], tokens[5]
    return d


def parse_service_line(line: str):
    """Procura 'nome + data + valores'. Retorna (nome, data, match) ou None."""
    for m in DATE_RE.finditer(line):
        rest = line[m.end():].strip()
        mm = SERVICE_REST_RE.match(rest)
        if mm:
            return line[: m.start()].strip(), m.group(), mm
    return None


def parse_servicos(pages_text: list[str], label: str, include_totals: bool,
                   extra_terms: list[str], convert: bool):
    """Retorna (records, ignored, warnings) para os relatórios APP/SITE."""
    num = parse_br_number if convert else (lambda x: x)
    records: list[dict] = []
    ignored: list[dict] = []
    warnings: list[str] = []
    buffer: list[str] = []
    stats = {"continuacoes": 0}
    current_page = 0

    def new_record(kind: str, name: str = "") -> dict:
        rec = {c: None for c in COLS_SERVICOS}
        rec["SERVIÇO"] = name
        rec["_kind"] = kind
        rec["_empty_prefix"] = False
        return rec

    def ignore(text: str, reason: str):
        ignored.append({"Arquivo": label, "Página": current_page, "Motivo": reason, "Texto": text})

    def fill_tail(rec: dict, tokens: list[str]) -> bool:
        mapped = map_tail(tokens)
        if mapped is None:
            return False
        for col, tok in mapped.items():
            rec[col] = num(tok)
        return True

    def flush(next_rec):
        """Decide a quem pertencem as linhas de texto soltas (nomes quebrados em várias linhas)."""
        if not buffer:
            return
        lines = list(buffer)
        text = " ".join(lines)
        buffer.clear()
        stats["continuacoes"] += len(lines)

        prev = records[-1] if records else None
        prev_is_svc = prev is not None and prev["_kind"] == "service"
        next_is_svc = next_rec is not None and next_rec["_kind"] == "service"

        if next_is_svc and not next_rec["SERVIÇO"]:
            if prev_is_svc and prev["_empty_prefix"]:
                k = len(lines) // 2  # nome centralizado: metade acima, metade abaixo
                prev["SERVIÇO"] = f"{prev['SERVIÇO']} {' '.join(lines[:k])}".strip()
                next_rec["SERVIÇO"] = " ".join(lines[k:])
            else:
                next_rec["SERVIÇO"] = text
        elif prev_is_svc:
            prev["SERVIÇO"] = f"{prev['SERVIÇO']} {text}".strip()
        elif next_is_svc:
            next_rec["SERVIÇO"] = f"{text} {next_rec['SERVIÇO']}".strip()
        elif prev is not None and prev["_kind"] in ("client", "total"):
            prev["SERVIÇO"] = f"{prev['SERVIÇO']} {text}".strip()
        else:
            rec = new_record("text", text)
            records.append(rec)
            warnings.append(f"{label}: texto solto sem linha de dados associada: '{text}'")

    for page_no, page_text in enumerate(pages_text, start=1):
        current_page = page_no
        for raw in page_text.split("\n"):
            line = collapse(raw)
            if not line:
                continue
            n = norm(line)
            is_client_start = bool(CLIENT_NAME_RE.match(line))

            # 1) Lixo de quebra de página (linhas de cliente são sempre protegidas)
            if not is_client_start:
                if any(re.search(p, n) for p in JUNK_SERVICOS) or (
                    extra_terms and any(t in n for t in extra_terms)
                ):
                    ignore(line, "lixo de cabeçalho/rodapé")
                    continue
                tokens_up = n.split()
                if tokens_up and all(t in HEADER_WORDS for t in tokens_up):
                    ignore(line, "cabeçalho de colunas")
                    continue

            # 2) Linha de cliente/agência
            m = CLIENT_RE.match(line) if is_client_start else None
            if m:
                rec = new_record("client", f"{m.group('code')} - {m.group('name').strip()}")
                if not fill_tail(rec, m.group("tail").split()):
                    rec["SERVIÇO"] = line  # contagem de números inesperada: guarda a linha inteira
                    warnings.append(f"{label} (pág. {page_no}): linha de cliente com números inesperados: '{line}'")
                flush(rec)
                records.append(rec)
                continue
            if is_client_start:
                rec = new_record("client", line)
                flush(rec)
                records.append(rec)
                continue

            # 3) Linha de serviço (tem data dd/mm/aa)
            svc = parse_service_line(line)
            if svc:
                name, date, mm = svc
                rec = new_record("service", name)
                rec["_empty_prefix"] = not name
                rec["DATA SERVIÇO"] = date
                rec["ADT"], rec["CHD."], rec["INF"], rec["QTD"] = (
                    num(mm.group("adt")), num(mm.group("chd")), num(mm.group("inf")), num(mm.group("qtd"))
                )
                rec["VOUCHER/RECIBO"] = mm.group("voucher")
                rec["TARIFA"] = num(mm.group("tarifa"))
                rec["VALOR VENDA"] = num(mm.group("venda"))
                rec["CATEGORIA SERVIÇO"] = mm.group("cat").strip() or None
                rec["% FEE"] = mm.group("pct")
                rec["VALOR FEE"] = num(mm.group("fee"))
                flush(rec)
                records.append(rec)
                continue

            # 4) Linhas de TOTAL (TOTAL ORIGEM, TOTAL RELATÓRIO...)
            if n.startswith("TOTAL"):
                tokens = line.split()
                idx = next((i for i, t in enumerate(tokens) if re.fullmatch(NUM_TOKEN, t)), len(tokens))
                rec = new_record("total", " ".join(tokens[:idx]))
                if tokens[idx:] and not fill_tail(rec, tokens[idx:]):
                    rec["SERVIÇO"] = line
                    warnings.append(f"{label} (pág. {page_no}): linha de total com números inesperados: '{line}'")
                flush(rec)
                if include_totals:
                    records.append(rec)
                else:
                    ignore(line, "linha de total (opção desmarcada)")
                continue

            # 5a) Cliente com nome quebrado: o fim do nome + os números vêm na linha seguinte
            prev = records[-1] if records else None
            if prev is not None and prev["_kind"] in ("client", "total") and prev["ADT"] is None:
                mt = re.match(rf"^(?P<name>.*?\D)(?P<tail>(?:\s+{NUM_TOKEN})+)\s*$", line)
                if mt and map_tail(mt.group("tail").split()):
                    flush(None)  # texto solto pendente pertence ao nome do cliente
                    prev["SERVIÇO"] = f"{prev['SERVIÇO']} {mt.group('name').strip()}".strip()
                    fill_tail(prev, mt.group("tail").split())
                    stats["continuacoes"] += 1
                    continue

            # 5b) Números soltos logo após cliente/total sem números
            if NUMERIC_ONLY_RE.match(line):
                prev = records[-1] if records else None
                if prev is not None and prev["_kind"] in ("client", "total") and prev["ADT"] is None \
                        and fill_tail(prev, line.split()):
                    continue
                warnings.append(f"{label} (pág. {page_no}): linha numérica sem identificação: '{line}'")
                ignore(line, "linha numérica órfã")
                continue

            # 6) Qualquer outro texto: continuação do nome (quebra de linha dentro da célula)
            buffer.append(line)

    flush(None)

    if stats["continuacoes"]:
        warnings.append(
            f"{label}: {stats['continuacoes']} linha(s) de continuação de nome foram anexadas ao serviço/cliente "
            "correspondente. Confira algumas no Excel."
        )
    return records, ignored, warnings


# --------------------------------------------------------------------------------------
# Relatório MANAGETOUR
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


def parse_managetour(pages_text: list[str], label: str, extra_terms: list[str], convert: bool):
    num = parse_br_number if convert else (lambda x: x)
    records, ignored, warnings = [], [], []

    for page_no, page_text in enumerate(pages_text, start=1):
        for raw in page_text.split("\n"):
            line = collapse(raw)
            if not line:
                continue
            n = norm(line)

            m = MT_RE.match(line)
            if m:
                records.append({
                    "Id": num(m.group("id")),
                    "File": m.group("file"),  # texto: preserva '615.128' exatamente
                    "Total Geral (Soma)": num(m.group("a")),
                    "Receita Operacional (Soma)": num(m.group("b")),
                    "Custo Operacao Rateio (Soma)": num(m.group("c")),
                    "Total NET - Previsto (Soma)": num(m.group("d")),
                })
                continue

            if any(re.search(p, n) for p in JUNK_MANAGETOUR) or (
                extra_terms and any(t in n for t in extra_terms)
            ):
                ignored.append({"Arquivo": label, "Página": page_no, "Motivo": "lixo de cabeçalho/rodapé", "Texto": line})
                continue
            if sum(1 for h in HEADER_MT if h in n) >= 3:
                ignored.append({"Arquivo": label, "Página": page_no, "Motivo": "cabeçalho de colunas", "Texto": line})
                continue

            ignored.append({"Arquivo": label, "Página": page_no, "Motivo": "não reconhecida como transação", "Texto": line})
            if len(re.findall(MONEY, line)) >= 2:
                warnings.append(f"{label} (pág. {page_no}): linha com valores não reconhecida – '{line}'")

    return records, ignored, warnings


# --------------------------------------------------------------------------------------
# DataFrame e Excel
# --------------------------------------------------------------------------------------


def records_to_df(records: list[dict], columns: list[str]) -> pd.DataFrame:
    clean = [{c: r.get(c) for c in columns} for r in records]  # descarta chaves internas
    return pd.DataFrame(clean, columns=columns, dtype=object)  # object: mantém int como int


def to_excel_bytes(df: pd.DataFrame, money_cols: list[str]) -> bytes:
    buffer = io.BytesIO()
    sheet = "Relatorio"
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]

        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.freeze_panes = "A2"

        for i, col in enumerate(df.columns, start=1):
            max_len = len(str(col))
            for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                cell = row[0]
                if cell.value is None:
                    continue
                if col in money_cols and isinstance(cell.value, (int, float)):
                    cell.number_format = "#,##0.00"
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(i)].width = min(max_len + 2, 70)
    return buffer.getvalue()


def make_filename(tipo: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "_", norm(tipo)).strip("_") + ".xlsx"


# --------------------------------------------------------------------------------------
# Interface Streamlit
# --------------------------------------------------------------------------------------


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
        convert = st.checkbox(
            "Converter números (formato 1.234,56) para número no Excel", value=True
        )
        include_totals = st.checkbox(
            "Manter linhas de TOTAL (TOTAL ORIGEM / TOTAL RELATÓRIO)",
            value=True,
            disabled=spec["kind"] != "servicos",
        )
        extra_text = st.text_area(
            "Termos extras a ignorar (um por linha). Linhas que contenham o termo são descartadas.",
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
            if spec["multi"] and len(files) != 2:
                st.warning(f"Esperado 2 arquivos para o relatório APP; recebidos {len(files)}. Processando mesmo assim.")

            files = sorted(files, key=lambda f: f.name)  # 1ª → 2ª quinzena pelo nome
            all_records, all_ignored, all_warnings, raw_samples = [], [], [], {}
            failed = False

            with st.spinner("Lendo PDF(s)..."):
                for f in files:
                    try:
                        pages = extract_pages_text(f.getvalue())
                        raw_samples[f.name] = pages[0] if pages else ""
                        if spec["kind"] == "servicos":
                            rec, ign, warn = parse_servicos(pages, f.name, include_totals, extra_terms, convert)
                        else:
                            rec, ign, warn = parse_managetour(pages, f.name, extra_terms, convert)
                        if not rec:
                            raise ValueError(
                                "Nenhuma linha de dados reconhecida. Veja o texto bruto em 'Diagnóstico' abaixo."
                            )
                        all_records += rec
                        all_ignored += ign
                        all_warnings += warn
                    except Exception as exc:  # noqa: BLE001
                        failed = True
                        st.error(f"Falha ao ler **{f.name}**: {exc}")
                        if f.name in raw_samples:
                            with st.expander(f"Diagnóstico – texto bruto da página 1 de {f.name}"):
                                st.code(raw_samples[f.name] or "(vazio)", language=None)

            if not failed:
                try:
                    cols = COLS_SERVICOS if spec["kind"] == "servicos" else COLS_MANAGETOUR
                    money = MONEY_SERVICOS if spec["kind"] == "servicos" else MONEY_MANAGETOUR
                    df = records_to_df(all_records, cols)  # colunas fixas, header uma única vez
                    st.session_state["result"] = {
                        "tipo": tipo,
                        "bytes": to_excel_bytes(df, money),
                        "filename": make_filename(tipo),
                        "preview": df.fillna("").astype(str),
                        "rows": len(df),
                        "files": [f.name for f in files],
                        "warnings": all_warnings,
                        "ignored": pd.DataFrame(all_ignored),
                        "raw": raw_samples,
                    }
                except Exception as exc:  # noqa: BLE001
                    st.error(f"Falha ao gerar o Excel: {exc}")

    result = st.session_state.get("result")
    if result and result["tipo"] == tipo:
        st.success(f"Conversão concluída: {result['rows']} linhas (arquivos: {', '.join(result['files'])}).")

        if result["warnings"]:
            with st.expander(f"⚠️ {len(result['warnings'])} aviso(s) – clique para conferir", expanded=True):
                for w in result["warnings"][:200]:
                    st.write("• " + w)

        st.download_button(
            label="⬇️ Baixar Excel (.xlsx)",
            data=result["bytes"],
            file_name=result["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.subheader("Pré-visualização")
        st.dataframe(result["preview"], use_container_width=True, height=450)

        if not result["ignored"].empty:
            with st.expander(f"Linhas ignoradas ({len(result['ignored'])}) – auditoria"):
                st.dataframe(result["ignored"].astype(str), use_container_width=True, height=300)

        with st.expander("Diagnóstico – texto bruto da página 1 (como o pdfplumber enxergou)"):
            for name, txt in result["raw"].items():
                st.caption(name)
                st.code(txt or "(vazio)", language=None)


if __name__ == "__main__":
    main()
