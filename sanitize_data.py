# -*- coding: utf-8 -*-
"""
sanitize_data.py
================
Pipeline de higienização científica da base de vendas Porsche.

Agente 1 — Cientista e Analista de Dados
----------------------------------------
Fluxo:
  1. Lê a planilha de trabalho `planilha base porsche (sanitizada).xlsx`.
  2. Varre as colunas originais aplicando regras de tratamento
     (padronização de capitalização, remoção de caracteres invisíveis,
     correção de tipos numéricos/datas, conversões de unidades).
  3. Grava o resultado de cada tratamento ESTRITAMENTE na respectiva
     coluna com sufixo "Sanitized" (mesma linha, mesmo registro).
  4. Exporta:
       - `planilha_porsche_producao.xlsx`  -> somente colunas sanitizadas
       - `dashboard_data.js`               -> ativo JS otimizado do dashboard
  5. Executa checagens de qualidade (schema `prompt do projeto porsche.md`).

Uso:  python sanitize_data.py
"""

import math
import re
from datetime import date, datetime

import pandas as pd

# --------------------------------------------------------------------------
# Configurações gerais
# --------------------------------------------------------------------------
INPUT_FILE = "planilha base porsche (sanitizada).xlsx"
PRODUCTION_XLSX = "planilha_porsche_producao.xlsx"
DASHBOARD_JS = "dashboard_data.js"

KM_TO_MILES = 0.621371
YEAR_MIN, YEAR_MAX = 1990, 2035

INVALID = "INVALID"

# Caracteres invisíveis / não imprimíveis frequentes em planilhas
INVISIBLE_RE = re.compile(r"[\u200b\u200c\u200d\ufeff\u00ad]")


def clean_text(value) -> str:
    """Remove caracteres invisíveis, espaços duplicados e normaliza espaços."""
    if value is None:
        return ""
    s = str(value)
    s = INVISIBLE_RE.sub("", s)
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


# --------------------------------------------------------------------------
# Conversão de números por extenso (inglês)
# --------------------------------------------------------------------------
NUM_WORDS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
    "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19, "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
    "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
}
SCALE_WORDS = {"hundred", "thousand"}


def word2num(text: str, year_mode: bool = False):
    """Converte números por extenso em inglês para inteiro.

    year_mode=True interpreta frases anuais como 'twenty twenty three' -> 2023.
    Retorna None quando o texto não é um número reconhecível.
    """
    tokens = re.findall(r"[a-z]+", str(text).lower())
    if not tokens:
        return None
    if not all(t in NUM_WORDS or t in SCALE_WORDS for t in tokens):
        return None

    if year_mode and "hundred" not in tokens and "thousand" not in tokens:
        # 'twenty twenty' -> 2020 | 'twenty twenty three' -> 2023
        if len(tokens) == 2:
            a, b = NUM_WORDS[tokens[0]], NUM_WORDS[tokens[1]]
            if a >= 20 and b < 100:
                return a * 100 + b
        if len(tokens) == 3:
            a, b, c = (NUM_WORDS[t] for t in tokens)
            if a >= 20 and b >= 20 and c < 10:
                return a * 100 + b + c
        return None

    cur, total = 0, 0
    for t in tokens:
        if t == "hundred":
            cur = (cur or 1) * 100
        elif t == "thousand":
            total += (cur or 1) * 1000
            cur = 0
        else:
            cur += NUM_WORDS[t]
    return total + cur


# --------------------------------------------------------------------------
# Parser numérico (separadores de milhar / decimais US e BR)
# --------------------------------------------------------------------------
def _parse_numeric_string(s: str):
    """Interpreta uma string puramente numérica com '.' e/ou ','.

    Regras:
      - '.' e ',' juntos  -> o último é o separador decimal;
      - apenas ','        -> grupos de 3 são milhar, senão decimal;
      - apenas '.'        -> grupos exatos de 3 são milhar, senão decimal.
    Retorna float ou None.
    """
    s = s.strip()
    if not re.fullmatch(r"[\d.,]+", s):
        return None

    has_dot, has_comma = "." in s, "," in s

    if has_dot and has_comma:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    elif has_dot:
        if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            s = s.replace(".", "")
        elif not re.fullmatch(r"\d+\.\d{1,2}", s):
            return None  # padrão de pontos não reconhecível

    try:
        return float(s)
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------
# Regra 1 — Datas (saída: 'YYYY-MM-DD' ou 'INVALID')
# --------------------------------------------------------------------------
MONTHS = {
    "january": 1, "jan": 1, "february": 2, "feb": 2, "march": 3, "mar": 3,
    "april": 4, "apr": 4, "may": 5, "june": 6, "jun": 6, "july": 7, "jul": 7,
    "august": 8, "aug": 8, "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10, "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}


def _valid_date(y: int, m: int, d: int):
    try:
        return date(y, m, d).isoformat()
    except ValueError:
        return None


def sanitize_date(value):
    """Normaliza datas para ISO 'YYYY-MM-DD'. Datas de calendário inválidas
    (ex.: 2024-13-05, 2024-02-30) tornam-se 'INVALID'."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return INVALID
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    s = clean_text(value).lower().strip()
    if not s:
        return INVALID

    # Formatos ISO: 2024-05-11, 2024/07/11, 2024.07.11 (com hora opcional)
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        iso = _valid_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return iso or INVALID

    # Formatos US: MM/DD/YYYY, MM/DD/YY, MM-DD-YY
    m = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})", s)
    if m:
        mo, dy, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = yr if yr > 99 else (2000 + yr if yr <= 35 else 1900 + yr)
        iso = _valid_date(year, mo, dy)
        return iso or INVALID

    # Mês por extenso: 'July 15th, 2024' | 'Mon DDth YYYY'
    m = re.match(r"^([a-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s+(\d{4})", s)
    if m and m.group(1) in MONTHS:
        iso = _valid_date(int(m.group(3)), MONTHS[m.group(1)], int(m.group(2)))
        return iso or INVALID

    return INVALID


# --------------------------------------------------------------------------
# Regra 2 — Modelos Porsche (rótulo canônico)
# --------------------------------------------------------------------------
CANONICAL_MODELS = {
    "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Turbo",
    "911 Turbo S", "911 GT3", "911 GT3 RS", "911 Dakar", "911 Targa 4",
    "911 Targa 4S", "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster GTS", "718 Spyder RS", "Cayenne",
    "Cayenne S", "Cayenne Coupe", "Cayenne E-Hybrid", "Cayenne Turbo",
    "Cayenne Turbo GT", "Macan", "Macan S", "Macan T", "Macan GTS",
    "Macan Electric", "Panamera", "Panamera 4", "Panamera 4S",
    "Panamera Turbo", "Panamera Turbo S", "Panamera 4 E-Hybrid", "Taycan",
    "Taycan 4S", "Taycan GTS", "Taycan Turbo", "Taycan Turbo S",
    "Taycan Cross Turismo",
}
CANONICAL_LOOKUP = {m.lower(): m for m in CANONICAL_MODELS}


def sanitize_model(value):
    """Normaliza para o rótulo canônico de modelo/trim. Modelos desconhecidos
    são apenas title-cased (nunca descartados)."""
    s = clean_text(value)
    if not s:
        return INVALID
    collapsed = re.sub(r"\s+", " ", s)
    return CANONICAL_LOOKUP.get(collapsed.lower(), collapsed.title())


# --------------------------------------------------------------------------
# Regra 3 — Model Year (4 dígitos ou 'INVALID')
# --------------------------------------------------------------------------
def sanitize_year(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return INVALID
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        year = int(value)
        return year if YEAR_MIN <= year <= YEAR_MAX else INVALID

    s = clean_text(value).lower().strip()

    # Formatos abreviados: '20-23', '20 23' (e anos de 4 dígitos)
    m = re.fullmatch(r"(19|20)[\s\-]?(\d{2})", s)
    if m:
        year = int(m.group(1) + m.group(2))
        return year if YEAR_MIN <= year <= YEAR_MAX else INVALID

    # Ano simples de 4 dígitos
    if re.fullmatch(r"\d{4}", s):
        year = int(s)
        return year if YEAR_MIN <= year <= YEAR_MAX else INVALID

    # Número por extenso (modo ano): 'twenty twenty three'
    year = word2num(s, year_mode=True)
    if year is None:
        year = word2num(s)
    if year is not None:
        return year if YEAR_MIN <= year <= YEAR_MAX else INVALID

    return INVALID


# --------------------------------------------------------------------------
# Regra 4 — Preço de venda (USD com 2 decimais, ex.: '985000.00')
# --------------------------------------------------------------------------
def sanitize_price(value):
    """Normaliza preços para valor USD numérico com 2 casas decimais.
    Suporta símbolos, 'USD/dollars', sufixo 'k', formatos US e BR
    (ex.: '$103.750,00') e valores por extenso."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return INVALID
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(float(value), 2)

    s = clean_text(value).lower()

    # Sufixo 'k': '$59k', '188k USD'
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*k", s)
    if m:
        base = _parse_numeric_string(m.group(1))
        if base is not None:
            return round(base * 1000, 2)

    # Valor por extenso: 'eighty two thousand usd' (ignora palavras de moeda)
    alpha = [t for t in re.findall(r"[a-z]+", s)
             if t not in ("usd", "dollar", "dollars")]
    if alpha:
        num = word2num(" ".join(alpha))
        if num is not None:
            return round(float(num), 2)

    # Extração numérica pura: remove moeda/texto e interpreta separadores
    numeric = re.sub(r"[^0-9.,]", "", s)
    if not numeric:
        return INVALID
    num = _parse_numeric_string(numeric)
    if num is None or num <= 0:
        return INVALID
    return round(num, 2)


# --------------------------------------------------------------------------
# Regra 5 — Quilometragem (milhas inteiras)
# --------------------------------------------------------------------------
ZERO_MILEAGE = {"zero", "0", "new", "new car", "brand new", "zero miles"}


def sanitize_mileage(value):
    """Normaliza para milhas inteiras. Converte KM para milhas quando
    explicitamente rotulado (1 km = 0.621371 mi, arredondado)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return INVALID
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(float(value)))

    s = clean_text(value).lower()
    if not s:
        return INVALID
    if s in ZERO_MILEAGE:
        return 0

    # Conversão explícita de quilômetros: 'KM 15.200', '24.000 km'
    if "km" in s:
        numeric = re.sub(r"[^0-9.,]", "", s)
        num = _parse_numeric_string(numeric)
        if num is not None:
            return int(round(num * KM_TO_MILES))
        return INVALID

    # Valor por extenso: 'twelve thousand miles' (ignora palavras de unidade)
    alpha = [t for t in re.findall(r"[a-z]+", s)
             if t not in ("mi", "mile", "miles")]
    if alpha:
        num = word2num(" ".join(alpha))
        if num is not None:
            return int(num)
        return INVALID

    # Extração numérica: '12,500 miles', '1.100 miles', '19,250 mi.'
    numeric = re.sub(r"[^0-9.,]", "", s)
    if not numeric:
        return INVALID
    num = _parse_numeric_string(numeric)
    if num is None:
        return INVALID

    # Notação decimal curta ('9.5' = 9,5 mil milhas): valores de rodagem
    # nessa faixa (0-100) com decimais representam milhar abreviada.
    if 0 < num < 100 and "." in s:
        num *= 1000
    return int(round(num))

# --------------------------------------------------------------------------
# Regra 6 — Método de pagamento (rótulo controlado)
# --------------------------------------------------------------------------
PAYMENT_RULES = [
    ("ach", "ACH Payment"),
    ("debit", "Debit Card"),
    ("credit", "Credit Card"),
    ("crypto", "Crypto Payment"),
    ("financ", "Financing"),
    ("leas", "Lease"),
    ("wire", "Wire Transfer"),
    ("bank", "Bank Transfer"),
    ("cash", "Cash"),
]


def sanitize_payment(value):
    s = clean_text(value).lower()
    if not s:
        return INVALID
    normalized = re.sub(r"[_\-]+", " ", s)
    for token, label in PAYMENT_RULES:
        if token in normalized:
            return label
    return s.title()


# --------------------------------------------------------------------------
# Regra 7 — Cidade (Title Case, preserva pontuação como 'St. Louis')
# --------------------------------------------------------------------------
def sanitize_city(value):
    s = clean_text(value)
    if not s:
        return INVALID
    return s.title()


# --------------------------------------------------------------------------
# Regra 8 — Estado (código USPS de 2 letras ou 'INVALID')
# --------------------------------------------------------------------------
US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT",
    "delaware": "DE", "florida": "FL", "georgia": "GA", "hawaii": "HI",
    "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME",
    "maryland": "MD", "massachusetts": "MA", "michigan": "MI",
    "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM",
    "new york": "NY", "north carolina": "NC", "north dakota": "ND",
    "ohio": "OH", "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA",
    "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY", "district of columbia": "DC",
}
US_ABBREVIATIONS = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC",
}


def sanitize_state(value):
    s = clean_text(value).lower()
    if not s:
        return INVALID
    if s in US_STATES:
        return US_STATES[s]
    if s.upper() in US_ABBREVIATIONS:
        return s.upper()
    return INVALID


# --------------------------------------------------------------------------
# Regra 9 — Status de entrega (rótulo controlado)
# --------------------------------------------------------------------------
DELIVERY_RULES = [
    ("awaiting pickup", "Awaiting Pickup"),
    ("awaiting delivery", "Awaiting Delivery"),
    ("awaiting review", "Awaiting Review"),
    ("pending approval", "Pending Approval"),
    ("pending review", "Pending Review"),
    ("in transit", "In Transit"),
    ("deliverd", "Delivered"),   # typo recorrente
    ("deliver", "Delivered"),
    ("cancel", "Cancelled"),
    ("ship", "Shipped"),
    ("pending", "Pending"),
]


def sanitize_delivery(value):
    s = clean_text(value).lower()
    if not s:
        return INVALID
    normalized = re.sub(r"[!.,_\-]+", " ", s)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    for token, label in DELIVERY_RULES:
        if token in normalized:
            return label
    return s.title()


# --------------------------------------------------------------------------
# Orquestração do pipeline
# --------------------------------------------------------------------------
# Mapa: coluna original -> coluna sanitizada -> função de tratamento
PIPELINE = {
    "sale_date": ("SaleDateSanitized", sanitize_date),
    "porsche_model": ("PorscheModelSanitized", sanitize_model),
    "model_year": ("ModelYearSanitized", sanitize_year),
    "sale_price": ("SalesPriceSanitized", sanitize_price),
    "vehicle_mileage": ("VehicleMileageSanitized", sanitize_mileage),
    "payment_method": ("PayMethodSanitized", sanitize_payment),
    "city": ("CitySanitized", sanitize_city),
    "state": ("StateSanitized", sanitize_state),
    "delivery_status": ("DeliveryStatusSanitized", sanitize_delivery),
}


def run_pipeline() -> pd.DataFrame:
    df = pd.read_excel(INPUT_FILE, sheet_name="Sanitized")
    print(f"[1/4] Planilha carregada: {INPUT_FILE} ({df.shape[0]} linhas, "
          f"{df.shape[1]} colunas)")

    # Aplica cada regra na sua coluna sanitizada correspondente
    for raw_col, (san_col, func) in PIPELINE.items():
        if raw_col not in df.columns:
            raise KeyError(f"Coluna original ausente: {raw_col}")
        df[san_col] = df[raw_col].apply(func)
        print(f"      - {raw_col:<18} -> {san_col}")

    # Checagens de qualidade (schema do projeto)
    print("[2/4] Checagens de qualidade:")
    sanitized_cols = [c for _, (c, _) in PIPELINE.items()]
    total_invalid = 0
    for col in sanitized_cols:
        blanks = df[col].isna().sum()
        invalids = int((df[col].astype(str) == INVALID).sum())
        total_invalid += invalids
        if blanks:
            raise AssertionError(f"Coluna sanitizada com vazios: {col}")
        print(f"      - {col:<26} INVALID: {invalids:>3} | vazios: {blanks}")
    print(f"      - Total de registros marcados como INVALID: "
          f"{total_invalid} de {len(df) * len(sanitized_cols)} células")

    # Exportação 1: pasta de trabalho atualizada (brutos + sanitizados)
    df.to_excel(INPUT_FILE, sheet_name="Sanitized", index=False)
    print(f"[3/4] Colunas '_sanitized' atualizadas in-loco em: {INPUT_FILE}")

    # Exportação 2: base de produção — SOMENTE colunas sanitizadas
    production = df[sanitized_cols].copy()
    production.to_excel(PRODUCTION_XLSX, index=False)
    print(f"[3/4] Base de produção exportada: {PRODUCTION_XLSX} "
          f"({production.shape[1]} colunas sanitizadas)")

    # Exportação 3: ativo JS otimizado para o dashboard
    records = []
    for _, row in df.iterrows():
        year = row["ModelYearSanitized"]
        price = row["SalesPriceSanitized"]
        mileage = row["VehicleMileageSanitized"]
        records.append({
            "id": str(row["sale_id"]),
            "data": row["SaleDateSanitized"],
            "modelo": row["PorscheModelSanitized"],
            "anoModelo": int(year) if isinstance(year, (int, float)) else None,
            "preco": float(price) if isinstance(price, (int, float)) else 0.0,
            "quilometragem": int(mileage) if isinstance(mileage, (int, float)) else None,
            "cidade": row["CitySanitized"],
            "estado": row["StateSanitized"],
            "formaPagamento": row["PayMethodSanitized"],
            "entrega": row["DeliveryStatusSanitized"],
        })

    with open(DASHBOARD_JS, "w", encoding="utf-8") as fh:
        fh.write("// Gerado automaticamente por sanitize_data.py — nao editar.\n")
        fh.write("// Base de producao sanitizada para o Porsche Intelligence Dashboard.\n")
        fh.write("window.PORSCHE_SANITIZED_DATA = ")
        fh.write(pd.Series(records).to_json(orient="records", force_ascii=False))
        fh.write(";\n")
    print(f"[3/4] Ativo de dados do dashboard gerado: {DASHBOARD_JS} "
          f"({len(records)} registros)")

    print("[4/4] Pipeline concluido com sucesso.")
    return df


if __name__ == "__main__":
    run_pipeline()
