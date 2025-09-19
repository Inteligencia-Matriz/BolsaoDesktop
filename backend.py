# -*- coding: utf-8 -*-
"""
backend.py
-------------------------------------------------
Contém toda a lógica de negócio, acesso ao Google Sheets e geração de PDF
para a aplicação Gestor do Bolsão. Este módulo é independente da interface gráfica.
"""
# --- Importações de Módulos ---
import re
import uuid
from datetime import date, timedelta, datetime
from functools import lru_cache
from pathlib import Path
import sys
import os
import base64 
import json
import requests 
import pytz

import gspread
import pandas as pd
import weasyprint
from google.oauth2.service_account import Credentials

# --------------------------------------------------
# UTILITÁRIOS DE ACESSO AO GOOGLE SHEETS (OTIMIZADOS)
# --------------------------------------------------
SPREAD_URL = "https://docs.google.com/spreadsheets/d/1qBV70qrPswnAUDxnHfBgKEU4FYAISpL7iVP0IM9zU2Q/edit#gid=0"

GCP_CREDS_B64 = '''ew0KICAidHlwZSI6ICJzZXJ2aWNlX2FjY291bnQiLA0KICAicHJvamVjdF9pZCI6ICJyYWl6YS1maXJlYmFzZSIsDQogICJwcml2YXRlX2tleV9pZCI6ICJhNTMwNzA5Y2EyYWRlYmY2YTk
zNTA1ODk2MzMwNzA1MzUzN2ZiN2FhIiwNCiAgInByaXZhdGVfa2V5IjogIi0tLS0tQkVHSU4gUFJJVkFURSBLRVktLS0tLVxuTUlJRXZ3SUJBREFOQmdrcWhraUc5dzBCQVFFRkFBU0NCS2
t3Z2dTbEFnRUFBb0lCQVFDN1pFSmtvMmhrc1dmblxubFpjZHh6OUVVSnU4MTF0YWNFZWpOM0g5eUdqKzVLOTdFYVNCWU4vaDA5WlhPcmtEcGxCQlRiSU9pVVB4ZHBFRFxuMVFGL240SU90e
mMyU0M1dURhTkRDcmloSFZRZHIrekVqQTlzRnhmWExibTZWaDBDY29zRE1TVmdtRFZBMUVFWVxuNjNUV3VrdXpvTnBWRTI2QVQ1WDU2aE91OExhUk9lcUF0elJlcEFleXhpODNMS0w1eWFR
UlhsRWhCVVR1ZlYyM1xuSGx5SEN4SFRoQlpSNzJpMzRVYTVLNmlySDNlWXpsL2RzV1ZLMFBQZzNrTkVId251QkNKbS83b0prNE1lMFZZT1xuVlZKSzdOaTA1VHBqNVJidkpNRkVBZUd0cmR
MWVRGR08rZWdvSEcyVWtGbnhmWEdjMEp3OGNPbG82Zmt6bmZHNlxuaGNlUWFhdGRBZ01CQUFFQ2dnRUFGdld0QzlNWXFKT3krU1V6TDM4WTRaZ2x5TUN6TFBUS1ZqVm1KbFBPaDRLQlxuL0
toMytURWpKVEoxRmRsNFR6bnFwZUdzNGdxUDlFOFVhLzJHY2pwYkwwM2oraWJrWjNBTTA1dEY3Vm1nVTRYWlxuVXVpZFFCOWhPS2hkMC9hV2xkVHVjdHpyNlRhaytiVTM1Nk43dkk1MVZZU
XRGQlR1S2xMMSszbWlZVUlWZ0Z2N1xuemprQXdqUkhKKy9mRkZSMnFTdzZBMnd4ZmVHcWFzS1lZRlJabVNTSVJ1eGpTVUZsV1lDQVMwRWN3QllxWUF4elxuU08zb0pMNFJKVFVwV2U2c2Fm
eDlTeUY3OVU1UUNEZUZGS21laHhHMkxESzN5TmpkLzNlcDR3MUtQdWNaSktZWlxuUjBTN21BbUFpcWxobWEzZlJ3YnB4U0J5aHVzOUlmR1hzeU5wZG1CY3NRS0JnUUQ3c1plVTRaWWFOSGh
5SVkvVVxuR09ISm5lUTFxTWFsL2hHRnVyWWt4Z3RhS3NGQnp3UU81aWVSUlp3V3JaaXdBWEU3YmI5NklsZUp2UnFZNlNMTFxuNzV1Z01BdWdtakhlYlJrV0hZMnY5VlV3RXlCQ3J0cUZLNj
dtWVBsZ2JJRGw1YWsyNGVIZ0hvQ3Z2bXpVdjZOdlxuREtFcWJaNVZYaG9HZDNDd0VzVkR1WGRDY1FLQmdRQyttUWJlSHZhUUpVZ21ueHViM3pJVlZOTUpERDNBSTFmdlxuOXBUZXU3SEhIN
FdxdDk1ZElWU0dDeVRsRXpOWTNqL3p4Q0U2YzFFMUxRWXBDSWY1dHU3aFNBSmJBWVo0dUlaV1xuZmFoVUZWUG5KMy9sb3VWYVRaVGhYQ0VWZjBVQ3N2SThlNURSa1VUUWRUN2J0Z2h1UnlP
TklRZW43cnR5WCtPR1xucTlaVnFvYVZyUUtCZ1FEUjdMUjE2NVVyTkJwRmJ2S3NQemlLMVpNU29qdFZGVXgrRWxWNjVHZHhnL2wrTHZDK1xua0gzdDczWVpnQjY2cGVsUVhOLzRPUTUwQm5
KWm1SRjVzTlpIUyt0V3YyVGFsSG40OVJ0STZFRnVBSFhHeUZuZlxuK3FnODVDTDZwbVQzMm81QkJUTkVuNHhMaUhMekd3ZHdSc05oUk41cmF6b2ZySjBqYmZSejRRdTBNUUtCZ1FDZFxuTT
g0c1NsR0hGcGpwOGZWdG5Ldk1XRWd2Z0Q4MlNIQnhaV25vUTlzZnAyb3lJckZ2RXR5S0tvcmx2ZTV0Ny9IRFxuZHhNSkNMQUVNZnlRdjQ2WGNrQ1k0ekcrS2dYbGNCeXRIYnRHanNqRE1Sc
1dKa01STmtnRGtGOWhRYldEd21CMVxuYmwxRjNKRnJkaWpBUXVXMVAwdWRUWTdvL2NqeFR4RjB0Q3AyUWMzN2lRS0JnUURVenlIQUhFZlg0VG0xRGhyaFxua1JnTFhqZHE3RXdaYUNZdzBC
VUVNZVg1bnRoZEFVZjJ1VnE1cVUzRW15M2pCZ2RQSU1HMGRKOHFGZS8rSkxFQ1xuNGo1Uk80MVhIaHYxRjVLdWhPbWk5Y1RoOWlJcnZCTG9BTm9tV0p3TUZzZmFnc0Q1aUl5ZW5keDYramV
CamN6d1xuYUlxdWRiYUVvRHJNRUFyQXpxdDZTVmdyWXc9PVxuLS0tLS1FTkQgUFJJVkFURSBLRVktLS0tLVxuIiwNCiAgImNsaWVudF9lbWFpbCI6ICJzaW1wbGlmaWNhQHJhaXphLWZpcm
ViYXNlLmlhbS5nc2VydmljZWFjY291bnQuY29tIiwNCiAgImNsaWVudF9pZCI6ICIxMDEzMjUxNTU2NTg0MzU1OTgwMjYiLA0KICAiYXV0aF91cmkiOiAiaHR0cHM6Ly9hY2NvdW50cy5nb
29nbGUuY29tL28vb2F1dGgyL2F1dGgiLA0KICAidG9rZW5fdXJpIjogImh0dHBzOi8vb2F1dGgyLmdvb2dsZWFwaXMuY29tL3Rva2VuIiwNCiAgImF1dGhfcHJvdmlkZXJfeDUwOV9jZXJ0
X3VybCI6ICJodHRwczovL3d3dy5nb29nbGVhcGlzLmNvbS9vYXV0aDIvdjEvY2VydHMiLA0KICAiY2xpZW50X3g1MDlfY2VydF91cmwiOiAiaHR0cHM6Ly93d3cuZ29vZ2xlYXBpcy5jb20
vcm9ib3QvdjEvbWV0YWRhdGEveDUwOS9zaW1wbGlmaWNhJTQwcmFpemEtZmlyZWJhc2UuaWFtLmdzZXJ2aWNlYWNjb3VudC5jb20iLA0KICAidW5pdmVyc2VfZG9tYWluIjogImdvb2dsZW
FwaXMuY29tIg0KfQ=='''

def resource_path(relative_path):
    """ Obtém o caminho absoluto para um recurso, funcionando para dev e para PyInstaller. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_gspread_client():
    """Conecta ao Google Sheets usando credenciais embutidas no código."""
    try:
        decoded_creds_json = base64.b64decode(GCP_CREDS_B64)
        creds_dict = json.loads(decoded_creds_json)
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
        return gspread.authorize(creds)
    except Exception as e:
        raise Exception(f"❌ Erro de autenticação com o Google Sheets a partir das credenciais embutidas: {e}")

client_cache = None
workbook_cache = None

def get_cached_client():
    """Retorna o cliente gspread em cache ou cria um novo."""
    global client_cache
    if client_cache is None:
        client_cache = get_gspread_client()
    return client_cache

def get_cached_workbook():
    """Retorna o workbook (planilha) em cache ou abre um novo."""
    global workbook_cache
    client = get_cached_client()
    if workbook_cache is None and client:
        workbook_cache = client.open_by_url(SPREAD_URL)
    return workbook_cache

@lru_cache(maxsize=32)
def get_ws(title: str):
    """Obtém uma aba (worksheet) pelo título e faz cache."""
    wb = get_cached_workbook()
    if wb:
        try:
            return wb.worksheet(title)
        except gspread.WorksheetNotFound:
            raise gspread.WorksheetNotFound(f"Aba da planilha com o nome '{title}' não foi encontrada.")
    return None

@lru_cache(maxsize=32)
def header_map(ws_title: str):
    """Cria um mapa de 'nome_da_coluna': indice para uma dada aba."""
    ws = get_ws(ws_title)
    if ws:
        headers = ws.row_values(1)
        return {h.strip(): i + 1 for i, h in enumerate(headers) if h and h.strip()}
    return {}

def get_values(ws, a1_range: str):
    """Função auxiliar para leitura de um range específico."""
    return ws.get(a1_range, value_render_option="UNFORMATTED_VALUE")

def find_row_by_id(ws, id_col_idx: int, target_id: str):
    """Encontra o número da linha de um registro pelo seu ID."""
    try:
        col_values = ws.col_values(id_col_idx)[1:]
        for i, value in enumerate(col_values, start=2):
            if str(value) == str(target_id):
                return i
    except Exception:
        return None
    return None

def batch_update_cells(ws, updates):
    """Executa múltiplas atualizações de células em uma única requisição à API."""
    if not updates:
        return
    fixed = []
    sheet_title_safe = ws.title.replace("'", "''")
    for u in updates:
        rng = u.get("range", "")
        if not rng:
            continue
        if "!" not in rng:
            rng = f"'{sheet_title_safe}'!{rng}"
        fixed.append({"range": rng, "values": u.get("values", [[]])})
    body = {"valueInputOption": "USER_ENTERED", "data": fixed}
    ws.spreadsheet.values_batch_update(body)

def ensure_size(ws, min_rows=2000, min_cols=40):
    """Garante que a planilha tenha um tamanho mínimo para evitar erros."""
    try:
        if ws and (ws.row_count < min_rows or ws.col_count < min_cols):
            ws.resize(rows=max(ws.row_count, min_rows), cols=max(ws.col_count, min_cols))
    except Exception:
        pass

def new_uuid():
    """Gera um ID único e curto (12 caracteres) para cada registro."""
    return uuid.uuid4().hex[:12]

def a1_col_letter(col_idx: int) -> str:
    """Converte um índice numérico de coluna (1, 2, 3) para sua letra A1 ('A', 'B', 'C')."""
    return re.sub(r"\d", "", gspread.utils.rowcol_to_a1(1, col_idx))

def batch_get_values_prefixed(ws, ranges, value_render_option="UNFORMATTED_VALUE"):
    """Faz uma leitura em lote de múltiplos ranges em uma única requisição."""
    if not ranges:
        return []
    title_safe = ws.title.replace("'", "''")
    prefixed = [f"'{title_safe}'!{r}" if "!" not in r else r for r in ranges]
    params = {'valueRenderOption': value_render_option}
    resp = ws.spreadsheet.values_batch_get(prefixed, params=params)
    return resp.get("valueRanges", [])

def load_resultados_snapshot():
    """
    Função otimizada para carregar os dados da aba 'Resultados_Bolsao'.
    """
    ws = get_ws("Resultados_Bolsao")
    if not ws:
        return {"rows": [], "id_to_rownum": {}}

    hmap = header_map("Resultados_Bolsao")
    
    columns_needed = [
        "REGISTRO_ID", "Nome do Aluno", "Unidade", "Bolsão", "% Bolsa", 
        "Valor da Mensalidade com Bolsa", "Escola de Origem", "Valor Negociado",
        "Responsável Financeiro", "Telefone", "Aluno Matriculou?", 
        "Observações (Form)", "Data/Hora"
    ]

    col_expectativa = "Expectativa de mensalidade"
    col_expectativa_fallback = "Valor Limite (PIA)"
    if col_expectativa in hmap:
        columns_needed.append(col_expectativa)
    elif col_expectativa_fallback in hmap:
        columns_needed.append(col_expectativa_fallback)

    missing = [c for c in columns_needed if c not in hmap]
    if missing:
        if not (len(missing) == 1 and missing[0] in [col_expectativa, col_expectativa_fallback]):
                 raise RuntimeError(f"Faltam colunas em 'Resultados_Bolsao': {', '.join(missing)}")

    letters = {c: a1_col_letter(hmap[c]) for c in columns_needed if c in hmap}
    ranges = [f"{letters[c]}2:{letters[c]}" for c in columns_needed if c in hmap]

    vranges = batch_get_values_prefixed(ws, ranges)
    series = {}
    valid_columns_from_fetch = [c for c in columns_needed if c in hmap]
    for c, vr in zip(valid_columns_from_fetch, vranges):
        vals = vr.get("values", [])
        series[c] = [row[0] if row else "" for row in vals]

    max_len = max((len(v) for v in series.values()), default=0)
    for c in valid_columns_from_fetch:
        col = series[c]
        if len(col) < max_len:
            col.extend([""] * (max_len - len(col)))

    rows = [{c: series.get(c, [""] * max_len)[i] for c in columns_needed} for i in range(max_len)]

    id_to_rownum = {}
    for i, rid in enumerate(series.get("REGISTRO_ID", []), start=2):
        if rid:
            id_to_rownum[str(rid)] = i

    return {"rows": rows, "id_to_rownum": id_to_rownum}

# --------------------------------------------------
# DADOS DE REFERÊNCIA E CONFIGURAÇÕES (CONSTANTES)
# --------------------------------------------------
BOLSA_MAP = {
    0: .30, 1: .30, 2: .30, 3: .35, 4: .40, 5: .40, 6: .44, 7: .45, 8: .46, 9: .47,
    10: .48, 11: .49, 12: .50, 13: .51, 14: .52, 15: .53, 16: .54, 17: .55, 18: .56, 19: .57,
    20: .60, 21: .65, 22: .70, 23: .80, 24: 1.00,
}
TUITION = {
    "1ª e 2ª Série EM Militar": {"anuidade": 36670.00, "parcela13": 2820.77},
    "1ª e 2ª Série EM Vestibular": {"anuidade": 36670.00, "parcela13": 2820.77},
    "1º ao 5º Ano": {"anuidade": 26654.00, "parcela13": 2050.31},
    "3ª Série (PV/PM)": {"anuidade": 36812.00, "parcela13": 2831.69},
    "3ª Série EM Medicina": {"anuidade": 36812.00, "parcela13": 2831.69},
    "6º ao 8º Ano": {"anuidade": 31354.00, "parcela13": 2411.85},
    "9º Ano EF II Militar": {"anuidade": 34146.00, "parcela13": 2626.62},
    "9º Ano EF II Vestibular": {"anuidade": 34146.00, "parcela13": 2626.62},
    "AFA/EN/EFOMM": {"anuidade": 14802.00, "parcela13": 1138.62},
    "CN/EPCAr": {"anuidade": 8863.00, "parcela13": 681.77},
    "ESA": {"anuidade": 7145.00, "parcela13": 549.62},
    "EsPCEx": {"anuidade": 14802.00, "parcela13": 1138.62},
    "IME/ITA": {"anuidade": 14802.00, "parcela13": 1138.62},
    "Medicina (Pré)": {"anuidade": 14802.00, "parcela13": 1138.62},
    "Pré-Vestibular": {"anuidade": 14802.00, "parcela13": 1138.62},
}
TURMA_DE_INTERESSE_MAP = {
    "1ª série IME ITA Jr": "1ª e 2ª Série EM Militar", "1ª série do EM - Militar": "1ª e 2ª Série EM Militar",
    "1ª série do EM - Pré-Vestibular": "1ª e 2ª Série EM Vestibular", "1º ano do EF1": "1º ao 5º Ano",
    "2ª série IME ITA Jr": "1ª e 2ª Série EM Militar", "2ª série do EM - Militar": "1ª e 2ª Série EM Militar",
    "2ª série do EM - Pré-Vestibular": "1ª e 2ª Série EM Vestibular", "2º ano do EF1": "1º ao 5º Ano",
    "3ª série do EM - AFA EN EFOMM": "3ª Série (PV/PM)", "3ª série do EM - ESA": "3ª Série (PV/PM)",
    "3ª série do EM - EsPCEx": "3ª Série (PV/PM)", "3ª série do EM - IME ITA": "3ª Série (PV/PM)",
    "3ª série do EM - Medicina": "3ª Série EM Medicina", "3ª série do EM - Pré-Vestibular": "3ª Série (PV/PM)",
    "3º ano do EF1": "1º ao 5º Ano", "4º ano do EF1": "1º ao 5º Ano", "5º ano do EF1": "1º ao 5º Ano",
    "6º ano do EF2": "6º ao 8º Ano", "7º ano do EF2": "6º ao 8º Ano", "8º ano do EF2": "6º ao 8º Ano",
    "9º ano do EF2 - Militar": "9º Ano EF II Militar", "9º ano do EF2 - Vestibular": "9º Ano EF II Vestibular",
    "Pré-Militar AFA EN EFOMM": "AFA/EN/EFOMM", "Pré-Militar CN EPCAr": "CN/EPCAr", "Pré-Militar ESA": "ESA",
    "Pré-Militar EsPCEx": "EsPCEx", "Pré-Militar IME ITA": "IME/ITA", "Pré-Vestibular": "Pré-Vestibular",
    "Pré-Vestibular - Medicina": "Medicina (Pré)",
}
UNIDADES_COMPLETAS = [
    "COLEGIO E CURSO MATRIZ EDUCACAO CAMPO GRANDE", "COLEGIO E CURSO MATRIZ EDUCAÇÃO TAQUARA",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO BANGU", "COLEGIO E CURSO MATRIZ EDUCACAO NOVA IGUACU",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO DUQUE DE CAXIAS", "COLEGIO E CURSO MATRIZ EDUCAÇÃO SÃO JOÃO DE MERITI",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO ROCHA MIRANDA", "COLEGIO E CURSO MATRIZ EDUCAÇÃO MADUREIRA",
    "COLEGIO E CURSO MATRIZ EDUCAÇÃO RETIRO DOS ARTISTAS", "COLEGIO E CURSO MATRIZ EDUCACAO TIJUCA",
]
UNIDADES_MAP = {name.replace("COLEGIO E CURSO MATRIZ EDUCACAO", "").replace("COLEGIO E CURSO MATRIZ EDUCAÇÃO", "").strip(): name for name in UNIDADES_COMPLETAS}
UNIDADES_LIMPAS = sorted(list(UNIDADES_MAP.keys()))
DESCONTOS_MAXIMOS_POR_UNIDADE = {
    "RETIRO DOS ARTISTAS": 0.50, "CAMPO GRANDE": 0.6320, "ROCHA MIRANDA": 0.6606,
    "TAQUARA": 0.6755, "NOVA IGUACU": 0.6700, "DUQUE DE CAXIAS": 0.6823,
    "BANGU": 0.6806, "MADUREIRA": 0.7032, "TIJUCA": 0.6800, "SÃO JOÃO DE MERITI": 0.7197,
}

# --------------------------------------------------
# FUNÇÕES DE LÓGICA E UTILITÁRIOS
# --------------------------------------------------
def get_current_brasilia_date() -> date:
    """Obtém a data atual de Brasília a partir de uma API online com fallback."""
    try:
        response = requests.get("http://worldtimeapi.org/api/timezone/America/Sao_Paulo", timeout=3)
        response.raise_for_status()
        data = response.json()
        current_datetime = datetime.fromisoformat(data['datetime'])
        return current_datetime.date()
    except Exception:
        utc_now = datetime.utcnow().replace(tzinfo=pytz.utc)
        br_tz = pytz.timezone("America/Sao_Paulo")
        return utc_now.astimezone(br_tz).date()

@lru_cache(maxsize=1)
def get_bolsao_name_for_date(target_date=None):
    """Verifica a data e retorna o nome do bolsão ou 'Bolsão Avulso'."""
    if target_date is None:
        target_date = get_current_brasilia_date()
    try:
        ws_bolsao = get_ws("Bolsão")
        dates_cells = ws_bolsao.get('A2:A', value_render_option='FORMATTED_STRING')
        names_cells = ws_bolsao.get('C2:C')
        dates_col = [cell[0] for cell in dates_cells if cell]
        names_col = [cell[0] for cell in names_cells if cell]
        for i, date_str in enumerate(dates_col):
            if i < len(names_col) and names_col[i]:
                try:
                    bolsao_date = datetime.strptime(date_str, "%d/%m/%Y").date()
                    if bolsao_date == target_date:
                        return names_col[i]
                except ValueError: continue
        return "Bolsão Avulso"
    except Exception: return "Bolsão Avulso"

# --- FUNÇÃO CORRIGIDA ---
def precos_2026(serie_modalidade: str) -> dict:
    """
    Busca os preços corretos no dicionário TUITION.
    A chave 'parcela13' é o valor da mensalidade e da primeira cota.
    """
    base = TUITION.get(serie_modalidade, {})
    if not base:
        return {"primeira_cota": 0.0, "parcela_mensal": 0.0, "anuidade": 0.0}
    
    # Usa o valor de 'parcela13' diretamente como a mensalidade base.
    valor_mensal = float(base.get("parcela13", 0.0))
    primeira_cota = valor_mensal # A primeira cota é igual à mensalidade base.
    
    # Usa a anuidade do dicionário se existir, caso contrário, calcula.
    anuidade_total = float(base.get("anuidade", primeira_cota + (12 * valor_mensal)))
    
    return {
        "primeira_cota": primeira_cota, 
        "parcela_mensal": valor_mensal, 
        "anuidade": anuidade_total
    }

def calcula_bolsa(acertos: int, serie_modalidade: str | None = None) -> float:
    """
    Calcula o percentual de bolsa com base no número de acertos e na série.
    Contém a regra padrão (24 questões) e a regra especial para o EF1 (10 questões).
    """
    if serie_modalidade == "1º ao 5º Ano":
        a = max(0, min(acertos, 10))
        if a == 0: return 0.0
        if 1 <= a <= 3: return 0.30
        if 4 <= a <= 5: return 0.50
        if 6 <= a <= 8: return 0.60
        return 0.65
    ac = max(0, min(acertos, 24))
    return BOLSA_MAP.get(ac, 0.30)

def format_currency(v: float) -> str:
    """Formata um número float para uma string de moeda brasileira (ex: R$ 1.234,56)."""
    try:
        v_float = float(v)
        return f"R$ {v_float:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except (ValueError, TypeError):
        return str(v)

def parse_brl_to_float(x) -> float:
    """Converte uma string de moeda brasileira (ex: 'R$ 1.234,56') para um float (1234.56)."""
    if isinstance(x, (int, float)):
        return float(x)
    if not x:
        return 0.0
    s = str(x).strip().replace("R$", "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def format_phone_mask(raw: str) -> str:
    """Aplica uma máscara de telefone (##) #####-#### a uma string de dígitos."""
    if raw is None:
        return ""
    digits = re.sub(r"\D", "", str(raw))
    digits = digits[:11]
    if len(digits) >= 11:
        return f"({digits[:2]}) {digits[2:7]}-{digits[7:11]}"
    elif len(digits) == 10:
        return f"({digits[:2]}) {digits[2:6]}-{digits[6:10]}"
    return digits

def gera_pdf_html(ctx: dict) -> bytes:
    """
    Gera um arquivo PDF a partir de um template HTML e um dicionário de dados.
    Retorna os bytes do PDF gerado.
    """
    base_dir = Path(__file__).parent
    html_path = base_dir / "carta.html"
    try:
        with open(html_path, encoding="utf-8") as f:
            html_template = f.read()
        html_renderizado = html_template
        for k, v in ctx.items():
            html_renderizado = html_renderizado.replace(f"{{{{{k}}}}}", str(v))
        html_obj = weasyprint.HTML(string=html_renderizado, base_url=str(base_dir))
        return html_obj.write_pdf()
    except FileNotFoundError:
        raise Exception("Arquivo 'carta.html' ou 'style.css' não encontrado no diretório.")
    except Exception as e:
        raise Exception(f"Erro ao gerar PDF: {e}")

def get_hubspot_data_for_activation():
    """Obtém dados da aba 'Hubspot' para a funcionalidade de carregar candidato."""
    try:
        ws_hub = get_ws("Hubspot")
        if not ws_hub:
            return pd.DataFrame()

        hmap_h = header_map("Hubspot")
        cols_needed = ["Unidade", "Nome do Candidato", "Contato ID", "Status do Contato",
                       "Contato Realizado", "Observações", "Celular Tratado", "Nome",
                       "E-mail", "Turma de Interesse - Geral", "Fonte original"]
        missing_cols = [c for c in cols_needed if c not in hmap_h]
        if missing_cols:
            raise Exception(f"As seguintes colunas necessárias não foram encontradas na aba 'Hubspot': {', '.join(missing_cols)}")

        data = ws_hub.get_all_records(head=1)
        df = pd.DataFrame(data)
        if "Contato Realizado" in df.columns:
            df.rename(columns={"Contato Realizado": "Contato realizado"}, inplace=True)
        return df

    except Exception as e:
        raise Exception(f"❌ Falha ao carregar dados do Hubspot: {e}")

def calcula_valor_minimo(unidade, serie_modalidade):
    """Calcula o valor mínimo de parcela negociável para uma unidade e série."""
    try:
        desconto_maximo = DESCONTOS_MAXIMOS_POR_UNIDADE.get(unidade, 0)
        precos = precos_2026(serie_modalidade)
        valor_anuidade_integral = precos.get("anuidade", 0.0)
        if valor_anuidade_integral > 0 and desconto_maximo > 0:
            valor_minimo_anual = valor_anuidade_integral * (1 - desconto_maximo)
            return valor_minimo_anual / 12
        else:
            return 0.0
    except Exception as e:
        raise Exception(f"❌ Erro ao calcular valor mínimo: {e}")

def gerar_html_material_didatico(unidade: str) -> str:
    """
    Gera o código HTML para as tabelas de material didático
    com base na unidade selecionada.
    """
    precos_gerais = {
        "Medicina": ("R$ 4.009,95", "11x de R$ 364,54"),
        "Pré-Vestibular": ("R$ 4.009,95", "11x de R$ 364,54"),
    }
    
    precos_militares = {
        "AFA/EN/EFOMM": ("R$ 2.333,73", "11x de R$ 212,16"),
        "EPCAR": ("R$ 2.501,36", "11x de R$ 227,40"),
        "ESA": ("R$ 1.111,98", "11x de R$ 101,09"),
        "EsPCEx": ("R$ 2.668,97", "11x de R$ 242,63"),
        "IME/ITA": ("R$ 2.333,73", "11x de R$ 212,16"),
    }

    precos_didatico_padrao = {
        "1ª ao 5ª ano": ("R$ 2.552,80", "11x de R$ 232,07"),
        "6ª ao 8ª ano": ("R$ 2.765,77", "11x de R$ 251,43"),
        "9ª ano Vestibular": ("R$ 2.872,69", "11x de R$ 261,15"),
        "1ª e 2ª série Vestibular": ("R$ 3.399,67", "11x de R$ 309,06"),
        "3ª série": ("R$ 4.009,95", "11x de R$ 364,54"),
    }

    precos_sao_joao = {
        "1ª ao 5ª ano": ("R$ 1.933,56", "11x de R$ 175,78"),
        "6ª ao 8ª ano": ("R$ 2.020,92", "11x de R$ 183,72"),
        "9ª ano Vestibular": ("R$ 2.019,84", "11x de R$ 183,62"),
        "1ª e 2ª série Vestibular": ("R$ 2.474,20", "11x de R$ 224,93"),
        "3ª série": ("R$ 2.932,21", "11x de R$ 266,56"),
    }
    
    precos_retiro = {
        "1ª ao 5ª ano": ("R$ 2.552,80", "11x de R$ 232,07"),
    }
    
    dados_didatico = {}
    
    if unidade == "SÃO JOÃO DE MERITI":
        titulo_didatico = "Material Didático (exclusivo São João de Meriti)"
        dados_didatico = precos_sao_joao
    elif unidade == "RETIRO DOS ARTISTAS":
        titulo_didatico = "Material Didático"
        dados_didatico = precos_didatico_padrao.copy()
        dados_didatico.update(precos_retiro)
    else:
        titulo_didatico = "Material Didático"
        dados_didatico = precos_didatico_padrao

    tabela_didatico_html = f'<table class="pag2"><tr><th colspan="3">{titulo_didatico}</th></tr>'
    for curso, valores in dados_didatico.items():
        tabela_didatico_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_didatico_html += '</table><br>'

    tabela_geral_html = '<table class="pag2"><tr><th colspan="3">Material Didático (geral)</th></tr>'
    for curso, valores in precos_gerais.items():
        tabela_geral_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_geral_html += '</table><br>'
    
    tabela_militares_html = '<table class="pag2"><tr><th colspan="3">Material Militares</th></tr>'
    for curso, valores in precos_militares.items():
        tabela_militares_html += f'<tr><td>{curso}</td><td>{valores[0]}</td><td>{valores[1]}</td></tr>'
    tabela_militares_html += '</table><br>'

    return tabela_didatico_html + tabela_geral_html + tabela_militares_html