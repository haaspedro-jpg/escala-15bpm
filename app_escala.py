import streamlit as st
import pandas as pd
import random
import json
import os
import time
import requests
import re
import unicodedata
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT

# --- CONFIGURAÇÕES DO BANCO DE DADOS NA NUVEM ---
BIN_ID = "6a0dc2716877513b27a34d82"
API_KEY = "$2a$10$mtrG79e1mc862Rn3ZbwgZu7hJhbdHyTb.bZqvA7cmvNKyxM9gOG9G"
URL_DB = f"https://api.jsonbin.io/v3/b/{BIN_ID}"
HEADERS_DB = {'Content-Type': 'application/json', 'X-Master-Key': API_KEY}

# --- DICIONÁRIO DE NOMES DE GUERRA ---
NOMES_DE_GUERRA = {
    "Alan Felipe Buck dos Santos": "FELIPE", "Ana Paula Machado": "MACHADO", "Arthur Amador Silva": "ARTHUR",
    "Carina dos Santos": "CARINA", "Carlos Eduardo Boava": "CARLOS BOAVA", "Carlos Eduardo Soares da Costa Filho": "DA COSTA",
    "Cláudio Márcio Silva Júnior": "CLÁUDIO", "Daniel Barbosa Ramos": "RAMOS", "Danilo Eduardo de Oliveira": "OLIVEIRA",
    "Dener Vinícius Gomes Rodrigo": "DENER", "Edilma Aparecida Sereia da Silva": "SEREIA", "Eduardo Cauan José": "JOSÉ",
    "Eduardo Júnior Marques": "JÚNIOR MARQUES", "Felipe Israel de Almeida Guilherme": "ISRAEL", "Filipe Aparecido dos Santos": "SANTOS",
    "Gabriel Magdiel Reis": "MAGDIEL", "Gabriela Regina Schneider": "SCHNEIDER", "Gustavo Henrique Nogueira": "NOGUEIRA",
    "Gustavo Tivirolli Favaro": "TIVIROLLI", "Hemerson Oliveira Pacheco Júnior": "HEMERSON", "Hilda Heloisa Andrade Cunha": "ANDRADE",
    "Igor Henrique Bibanco": "BIBANCO", "Israel Mendes Martins": "MENDES", "João Victor Bertola da Silva": "BERTOLA",
    "João Victor Rapharred de Oliveira Abrahao": "RAPHARRED", "Jonathan Henrique das Neves": "HENRIQUE",
    "Josué Almeida de Souza": "JOSUÉ", "Kauane Stefany Lopes": "LOPES", "Leonardo Brenes Burque": "BURQUE",
    "Leonardo Cardoso Cosmos": "COSMOS", "Leonardo Santiago Rodrigues": "SANTIAGO", "Letícia Lopes Martelossi": "LETÍCIA",
    "Lirian Borges Barbosa": "LIRIAN", "Lucas Eduardo Dornellas de Oliveira": "EDUARDO", "Lucas Gabriel Souza de Campos": "CAMPOS",
    "Lucas Matheus dos Santos Lopes": "LUCAS LOPES", "Lucas Silva Piccioni": "PICCIONI", "Lucas Vinícius da Silva Panta": "PANTA",
    "Mateus Fujita Araújo": "FUJITA", "Matheus Silva Aleluia": "ALELUIA", "Nathalia Vargas Lopes": "VARGAS",
    "Priscila de Andrade Correia": "CORREIA", "Rafael Ribeiro Garcia": "RAFAEL GARCIA", "Ricardo Vinícius Zamparoni": "ZAMPARONI",
    "Tailer Costa Ribeiro dos Santos": "TAILER", "Thiago Zanin": "THIAGO ZANIN", "Thialis Venâncio dos Santos": "VENÂNCIO",
    "Tiago da Silva Ferreira": "TIAGO FERREIRA", "Victor Gonçalves Brum Silva": "BRUM", "Vitor Hugo Cardoso Almondes": "ALMONDES",
    "Vitor Humberto Ortega Nascimento": "ORTEGA", "Vitor Kauan Amorim dos Reis": "VITOR REIS", "Vitoria Cristina Cortelini": "CORTELINI",
    "Wallison Rodrigo Passerini": "PASSERINI", "Wesley Rodrigues Padovan": "PADOVAN"
}

MESES_PT = {1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL', 5: 'MAIO', 6: 'JUNHO', 
            7: 'JULHO', 8: 'AGOSTO', 9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'}
DIAS_SEMANA_PT = {0: 'segunda', 1: 'terça', 2: 'quarta', 3: 'quinta', 4: 'sexta', 5: 'sábado', 6: 'domingo'}

def padronizar_texto(texto):
    if not isinstance(texto, str): return ""
    t = ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))
    return t.upper().replace(" ", "").strip()

@st.cache_data(ttl=5)
def carregar_banco_nuvem():
    try:
        req = requests.get(URL_DB, headers=HEADERS_DB)
        if req.status_code == 200:
            data = req.json()['record']
            if 'ativo' not in data: data = {'ativo': {}, 'backups': {}, 'escalas_arquivadas': {}}
            if 'escalas_arquivadas' not in data: data['escalas_arquivadas'] = {}
            return data
        return None
    except: return None

def salvar_banco_nuvem(dados_completos):
    try:
        req = requests.put(URL_DB, json=dados_completos, headers=HEADERS_DB)
        return req.status_code == 200
    except: return False

def ler_planilha_segura(caminho_ou_arquivo):
    try:
        if str(caminho_ou_arquivo).endswith('.xlsx'): df = pd.read_excel(caminho_ou_arquivo)
        else:
            try: df = pd.read_csv(caminho_ou_arquivo, sep=None, engine='python', encoding='utf-8')
            except: df = pd.read_csv(caminho_ou_arquivo, sep=None, engine='python', encoding='latin-1')
        df.columns = df.columns.str.strip().str.upper()
        return df
    except: return None

def estilizar_celula_nome(cell, nome_completo, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]; p.alignment = align
    nome_guerra = obter_nome_guerra(nome_completo)
    if not nome_guerra:
        run = p.add_run(nome_completo); formatar_texto_pm(run, 'Arial', 10, bold=False)
        return
    def rm_acento(s): return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c)).upper()
    guerra_words = set(rm_acento(nome_guerra).split()); parts = re.split(r'(\s+)', nome_completo)
    for part in parts:
        if not part.strip(): run = p.add_run(part); formatar_texto_pm(run, 'Arial', 10, bold=False)
        else:
            run = p.add_run(part); formatar_texto_pm(run, 'Arial', 10, bold=(rm_acento(part) in guerra_words))

# --- INTERFACE ---
st.set_page_config(page_title="Gerador de Escala CFP", layout="wide")
st.title("Gerador Automático - Escala de Guarda CFP")

dados_nuvem = carregar_banco_nuvem()
if dados_nuvem is None: st.stop()
historico_ativo = dados_nuvem.get('ativo', {})

# Nova Aba de Importação Manual
with st.expander("➕ Importar Escala Manual"):
    st.info("Suba um CSV com colunas 'NOME' e 'DATA (DD/MM/YYYY)'. O sistema somará os pontos.")
    uploaded_manual = st.file_uploader("Arquivo de escala manual:", type=['csv'])
    if uploaded_manual:
        df_manual = ler_planilha_segura(uploaded_manual)
        if st.button("Confirmar e Atualizar Contagem"):
            for _, row in df_manual.iterrows():
                nome_manual = padronizar_texto(str(row['NOME']))
                data_manual = str(row['DATA'])
                for cpf, dados in historico_ativo.items():
                    if padronizar_texto(dados['nome']) == nome_manual:
                        historico_ativo[cpf]['normal'] += 1
                        historico_ativo[cpf]['datas_servico'].append({"data": data_manual, "tipo": "Manual"})
            dados_nuvem['ativo'] = historico_ativo
            salvar_banco_nuvem(dados_nuvem)
            st.success("Contagem atualizada!")

# Botão de Limpeza no Acervo
if st.button("🗑️ Limpar Todo o Acervo de Escalas Geradas"):
    dados_nuvem['escalas_arquivadas'] = {}
    salvar_banco_nuvem(dados_nuvem)
    st.rerun()

# ... (restante da lógica de geração permanece igual)
