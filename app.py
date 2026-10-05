import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="Conversor PDF para Excel", layout="wide")

def is_numeric(s):
    """Verifica se a string é composta apenas por números, pontos, vírgulas e sinais."""
    return re.match(r'^-?[\d.,]+$', s) is not None

def parse_line_coords(words):
    """
    Classifica cada palavra lida do PDF com base na sua posição horizontal (Eixo X).
    Isso evita que textos muito longos de 'Serviços' se misturem com as quantidades numéricas.
    """
    # Ordena da esquerda para a direita
    words.sort(key=lambda w: w['x0'])
    
    adt = chd = inf = qtd = voucher = tarifa = venda = categoria = fee = ""
    texto_principal_words = []
    
    for w in words:
        text = w['text']
        x0 = w['x0']
        
        # Mapeamento estrito por margens visuais do PDF
        if 270 <= x0 < 305 and is_numeric(text):
            adt = text
        elif 305 <= x0 < 335 and is_numeric(text):
            chd = text
        elif 335 <= x0 < 365 and is_numeric(text):
            inf = text
        elif 365 <= x0 < 390 and is_numeric(text):
            qtd = text
        elif 390 <= x0 < 470 and is_numeric(text):
            voucher = text
        elif 470 <= x0 < 530 and is_numeric(text):
            tarifa = text
        elif 530 <= x0 < 590 and is_numeric(text):
            venda = text
        elif 590 <= x0 < 720:
            categoria += text + " "
        elif 720 <= x0 < 780:
            pass # Ignora coluna de PERCENTUAL (1,00%) que fica antes do FEE
        elif 780 <= x0 and is_numeric(text):
            fee = text
        else:
            # Se não caiu em nenhuma coluna estrita, pertence ao texto principal
            texto_principal_words.append(text)
            
    texto_principal = " ".join(texto_principal_words)
    return texto_principal, categoria.strip(), adt, chd, inf, qtd, voucher, tarifa, venda, fee

def extract_client_info(texto_principal):
    """Separa o ID, o Nome e o Site da string do cliente."""
    file_id, site = "", ""
    
    # 1. Pega o ID inicial
    m_file = re.search(r'^(\d{5,})', texto_principal)
    if m_file:
        file_id = m_file.group(1)
        texto_principal = texto_principal.replace(file_id, "", 1)
        
    # 2. Pega o Site (última ocorrência entre parênteses)
    sites = re.findall(r'\((.*?)\)', texto_principal)
    if sites:
        site = sites[-1]
        texto_principal = texto_principal.rsplit(f"({site})", 1)[0]
        
    # 3. O que sobrar é o nome
    nome = re.sub(r'^[\s-]*', '', texto_principal).strip()
    return file_id, nome, site

def extract_service_info(texto_principal):
    """Separa a Data e o Nome do Serviço."""
    data_servico = ""
    m_data = re.search(r'^(\d{2}/\d{2}/\d{2})', texto_principal)
    if m_data:
        data_servico = m_data.group(1)
        texto_principal = texto_principal.replace(data_servico, "", 1)
        
    servico = re.sub(r'^[\s-]*', '', texto_principal).strip()
    return data_servico, servico

def process_pdf(pdf_file):
    data = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            
            # Agrupa as palavras em linhas baseando-se na altura (Y)
            lines = []
            words.sort(key=lambda w: w['top'])
            current_line = []
            current_top = None
            
            for w
