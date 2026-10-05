#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
PORSCHE BRASIL | INTELLIGENCE DATA PIPELINE & SANITIZATION ENGINE
=============================================================================
Autor: Agente Cientista e Analista de Dados Sênior (Porsche Analytics)
Descrição:
    Pipeline robusto, modular e com validação estatística rigorosa para
    ingestão, higienização, normalização e exportação de dados comerciais
    da Porsche.

Objetivos do Pipeline:
    1. Higienizar colunas brutas aplicando regras determinísticas de negócio
       (datas reais, anos válidos, moedas, distâncias, localidades, modelos canônicos).
    2. Gerar a planilha completa 'planilha base porsche (sanitizada).xlsx'
       com colunas originais e colunas sanitizadas adjacentes.
    3. Gerar a planilha oficial de produção 'planilha_porsche_producao.xlsx'
       contendo EXCLUSIVAMENTE as colunas com sufixo '_sanitized'.
    4. Exportar versões leves em CSV e JSON ('planilha_porsche_producao.csv' /
       'planilha_porsche_producao.json') para ingestão assíncrona no dashboard.

Uso:
    python sanitize_data.py [--input CAMINHO] [--out-combined CAMINHO] [--out-prod CAMINHO]
=============================================================================
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ---------------------------------------------------------------------------
# Constantes e Configurações Globais
# ---------------------------------------------------------------------------
DEFAULT_INPUT_PATH = Path("planilha base porsche.xlsx")
DEFAULT_COMBINED_OUTPUT = Path("planilha base porsche (sanitizada).xlsx")
DEFAULT_PROD_OUTPUT_XLSX = Path("planilha_porsche_producao.xlsx")
DEFAULT_PROD_OUTPUT_CSV = Path("planilha_porsche_producao.csv")
DEFAULT_PROD_OUTPUT_JSON = Path("planilha_porsche_producao.json")

INVALID = "INVALID"
KM_TO_MILES_FACTOR = 0.621371

MONTH_MAP: Dict[str, int] = {
    "jan": 1, "january": 1, "janeiro": 1,
    "feb": 2, "february": 2, "fevereiro": 2,
    "mar": 3, "march": 3, "marco": 3, "março": 3,
    "apr": 4, "april": 4, "abril": 4,
    "may": 5, "maio": 5,
    "jun": 6, "june": 6, "junho": 6,
    "jul": 7, "july": 7, "julho": 7,
    "aug": 8, "august": 8, "agosto": 8,
    "sep": 9, "sept": 9, "september": 9, "setembro": 9,
    "oct": 10, "october": 10, "outubro": 10,
    "nov": 11, "november": 11, "novembro": 11,
    "dec": 12, "december": 12, "dezembro": 12,
}

WORDS_TO_NUM: Dict[str, int] = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4,
    "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9,
    "ten": 10, "eleven": 11, "twelve": 12, "thirteen": 13,
    "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17,
    "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30,
    "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90, "hundred": 100, "thousand": 1000,
    "million": 1000000,
}

CANONICAL_PORSCHE_MODELS = [
    # 911 Series
    "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Carrera 4", "911 Carrera 4S",
    "911 Carrera Cabriolet", "911 Turbo", "911 Turbo S", "911 Turbo S Cabriolet",
    "911 GT3", "911 GT3 RS", "911 GT3 Touring", "911 Dakar", "911 S/T",
    "911 Targa 4", "911 Targa 4S", "911 Targa 4 GTS",
    # 718 Series
    "718 Cayman", "718 Cayman S", "718 Cayman GTS", "718 Cayman GTS 4.0", "718 Cayman GT4", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster S", "718 Boxster GTS", "718 Boxster GTS 4.0", "718 Spyder", "718 Spyder RS",
    # Cayenne Series
    "Cayenne", "Cayenne S", "Cayenne GTS", "Cayenne Coupe", "Cayenne S Coupe", "Cayenne GTS Coupe",
    "Cayenne E-Hybrid", "Cayenne S E-Hybrid", "Cayenne Turbo E-Hybrid",
    "Cayenne Turbo", "Cayenne Turbo GT",
    # Macan Series
    "Macan", "Macan T", "Macan S", "Macan GTS", "Macan Electric", "Macan 4 Electric", "Macan Turbo Electric",
    # Panamera Series
    "Panamera", "Panamera 4", "Panamera 4S", "Panamera GTS",
    "Panamera Turbo", "Panamera Turbo S", "Panamera 4 E-Hybrid", "Panamera 4S E-Hybrid", "Panamera Turbo E-Hybrid",
    # Taycan Series
    "Taycan", "Taycan 4", "Taycan 4S", "Taycan GTS", "Taycan Turbo", "Taycan Turbo S", "Taycan Turbo GT",
    "Taycan Cross Turismo", "Taycan 4 Cross Turismo", "Taycan 4S Cross Turismo", "Taycan Turbo Cross Turismo",
    "Taycan Sport Turismo", "Taycan GTS Sport Turismo"
]

_MODEL_LOOKUP = {
    re.sub(r"[^a-z0-9]", "", m.lower()): m for m in CANONICAL_PORSCHE_MODELS
}

US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID",
    "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE",
    "nevada": "NV", "new hampshire": "NH", "new jersey": "NJ",
    "new mexico": "NM", "new york": "NY", "north carolina": "NC",
    "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC",
    "south dakota": "SD", "tennessee": "TN", "texas": "TX", "utah": "UT",
    "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "district of columbia": "DC",
}
US_STATE_CODES = set(US_STATES.values())

PAYMENT_NORMALIZATION_MAP = {
    "credit card": "Credit Card",
    "creditcard": "Credit Card",
    "credit": "Credit Card",
    "credit card payment": "Credit Card",
    "cartao de credito": "Credit Card",
    "debit card": "Debit Card",
    "debitcard": "Debit Card",
    "debit": "Debit Card",
    "cartao de debito": "Debit Card",
    "bank transfer": "Bank Transfer",
    "banktransfer": "Bank Transfer",
    "bank transfer payment": "Bank Transfer",
    "transferencia bancaria": "Bank Transfer",
    "ted": "Bank Transfer",
    "doc": "Bank Transfer",
    "wire transfer": "Wire Transfer",
    "wiretransfer": "Wire Transfer",
    "wire": "Wire Transfer",
    "bank wire": "Wire Transfer",
    "wire bank": "Wire Transfer",
    "financing": "Financing",
    "finance": "Financing",
    "financing plan": "Financing",
    "financiamento": "Financing",
    "lease": "Lease",
    "leasing": "Lease",
    "lease plan": "Lease",
    "cash": "Cash",
    "cash payment": "Cash",
    "dinheiro": "Cash",
    "a vista": "Cash",
    "pix": "Bank Transfer",
    "ach": "ACH Payment",
    "ach payment": "ACH Payment",
    "crypto": "Crypto Payment",
    "crypto payment": "Crypto Payment",
    "cryptocurrency": "Crypto Payment",
    "bitcoin": "Crypto Payment",
}

DELIVERY_STATUS_MAP = {
    "delivered": "Delivered",
    "deliverd": "Delivered",
    "entregue": "Delivered",
    "in transit": "In Transit",
    "intransit": "In Transit",
    "em transito": "In Transit",
    "pending": "Pending",
    "pendente": "Pending",
    "cancelled": "Cancelled",
    "canceled": "Cancelled",
    "cancelado": "Cancelled",
    "awaiting delivery": "Awaiting Delivery",
    "aguardando entrega": "Awaiting Delivery",
    "awaiting pickup": "Awaiting Pickup",
    "aguardando retirada": "Awaiting Pickup",
    "pending approval": "Pending Approval",
    "aprovacao pendente": "Pending Approval",
    "pending review": "Pending Review",
    "revisao pendente": "Pending Review",
    "shipped": "Shipped",
    "despachado": "Shipped",
    "enviado": "Shipped",
    "awaiting review": "Awaiting Review",
}


# ---------------------------------------------------------------------------
# Funções Auxiliares de Sanitização
# ---------------------------------------------------------------------------
def _normalize_string_key(s: str) -> str:
    """Remove acentos, pontuações e converte para minúsculas compactas."""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[\W_]+", " ", s.lower()).strip()
    return re.sub(r"\s+", " ", s)


def _safe_date(year: int, month: int, day: int) -> str:
    """Valida o calendário real antes de gerar a data ISO YYYY-MM-DD."""
    try:
        return dt.date(year, month, day).strftime("%Y-%m-%d")
    except (ValueError, TypeError, OverflowError):
        return INVALID


def _smart_title_case(text: str) -> str:
    """Aplica Title Case preservando siglas e hífens."""
    if not text:
        return ""
    keep_upper = {"GT", "GTS", "RS", "GT3", "GT4", "S", "T", "4", "4S", "E-HYBRID", "USA", "ACH", "PIX", "KM", "ID"}
    words = text.split()
    out = []
    for w in words:
        upper = w.upper()
        if upper in keep_upper:
            out.append(upper)
        elif "-" in w:
            parts = [p.upper() if p.upper() in keep_upper else p.capitalize() for p in w.split("-")]
            out.append("-".join(parts))
        else:
            out.append(w.capitalize())
    return " ".join(out)


def _words_to_number(text: str) -> Optional[float]:
    """Converte números textuais como 'twenty twenty four' ou 'two hundred thousand'."""
    tokens = re.findall(r"[a-z]+", text.lower())
    if not tokens or any(t not in WORDS_TO_NUM for t in tokens):
        return None

    # Caso especial para ano (ex: "twenty twenty four" -> 2024)
    if all(WORDS_TO_NUM[t] < 100 for t in tokens):
        nums = [WORDS_TO_NUM[t] for t in tokens]
        groups: List[int] = []
        i = 0
        while i < len(nums):
            if i + 1 < len(nums) and nums[i] >= 20 and nums[i] % 10 == 0 and nums[i + 1] < 10:
                groups.append(nums[i] + nums[i + 1])
                i += 2
            else:
                groups.append(nums[i])
                i += 1
        if len(groups) == 2:
            return float(groups[0] * 100 + groups[1])
        if len(groups) == 1:
            return float(groups[0])

    # Caso geral com hundred, thousand, million
    total = 0.0
    current = 0.0
    for t in tokens:
        v = float(WORDS_TO_NUM[t])
        if v == 100.0:
            current = max(1.0, current) * 100.0
        elif v == 1000.0:
            current = max(1.0, current) * 1000.0
            total += current
            current = 0.0
        elif v == 1000000.0:
            current = max(1.0, current) * 1000000.0
            total += current
            current = 0.0
        else:
            current += v
    return total + current


def _parse_numeric_value(raw_val: Any) -> Optional[float]:
    """Trata formatos numéricos mistos (ponto, vírgula, milhar, decimais)."""
    if raw_val is None:
        return None
    if isinstance(raw_val, (int, float)) and not isinstance(raw_val, bool):
        return float(raw_val)

    s = str(raw_val).strip()
    if not s:
        return None

    # Verifica se há letras por extenso
    if re.search(r"[a-zA-Z]", s):
        # Limpar símbolos de moeda antes
        clean_text = re.sub(r"[\$,\.]", " ", s)
        clean_text = re.sub(r"\b(usd|dollars?|dolares|reais|brl)\b", " ", clean_text, flags=re.I).strip()
        num_from_text = _words_to_number(clean_text)
        if num_from_text is not None:
            return num_from_text

    # Capturar multiplicador 'k'
    multiplier = 1.0
    if re.search(r"[0-9.,]+\s*k\b", s, flags=re.I):
        multiplier = 1000.0

    # Extrair sequência numérica com separadores
    m = re.search(r"([0-9]+(?:[.,][0-9]+)*)", s)
    if not m:
        return None
    num_str = m.group(1)

    has_dot = "." in num_str
    has_comma = "," in num_str

    if has_dot and has_comma:
        if num_str.rfind(",") > num_str.rfind("."):
            # Formato Europeu / Brasileiro: 100.000,50
            num_str = num_str.replace(".", "").replace(",", ".")
        else:
            # Formato Americano: 100,000.50
            num_str = num_str.replace(",", "")
    elif has_comma:
        # Apenas vírgula: se tem 2 dígitos no final e apenas 1 vírgula, é decimal; senão é milhar
        parts = num_str.rsplit(",", 1)
        if len(parts) == 2 and len(parts[1]) == 2 and num_str.count(",") == 1:
            num_str = num_str.replace(",", ".")
        else:
            num_str = num_str.replace(",", "")
    elif has_dot:
        # Apenas ponto: se tem múltiplos pontos ou o último segmento não tem 2 dígitos, é milhar
        parts = num_str.rsplit(".", 1)
        if num_str.count(".") > 1 or len(parts[1]) != 2:
            num_str = num_str.replace(".", "")

    try:
        return float(num_str) * multiplier
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------------------
# Sanitizadores de Cada Coluna do Schema
# ---------------------------------------------------------------------------
def sanitize_sale_id(value: Any) -> str:
    """Normaliza o identificador da venda."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = str(value).strip()
    digits = re.sub(r"\D", "", s)
    return digits if digits else s


def sanitize_date(value: Any) -> str:
    """Normaliza datas para o padrão ISO YYYY-MM-DD com validação de calendário."""
    if value is None or str(value).strip() == "":
        return INVALID

    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%Y-%m-%d")

    s = str(value).strip()
    if not s:
        return INVALID

    # YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m:
        return _safe_date(int(m.group(1)), int(m.group(2)), int(m.group(3)))

    # MM/DD/YYYY or MM-DD-YYYY
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$", s)
    if m:
        return _safe_date(int(m.group(3)), int(m.group(1)), int(m.group(2)))

    # MM/DD/YY or MM-DD-YY
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2})$", s)
    if m:
        yy = int(m.group(3))
        year = 2000 + yy if yy < 70 else 1900 + yy
        return _safe_date(year, int(m.group(1)), int(m.group(2)))

    # Month DDth, YYYY / Mon DD YYYY / Month DD YYYY
    m = re.match(r"^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?[,\s]+(\d{4})$", s)
    if m:
        mon_key = m.group(1).lower().rstrip(".")
        if mon_key in MONTH_MAP:
            return _safe_date(int(m.group(3)), MONTH_MAP[mon_key], int(m.group(2)))

    # DD de Month de YYYY (formato pt-br)
    m = re.match(r"^(\d{1,2})\s+de\s+([A-Za-z]+)\s+de\s+(\d{4})$", s, flags=re.I)
    if m:
        mon_key = m.group(2).lower()
        if mon_key in MONTH_MAP:
            return _safe_date(int(m.group(3)), MONTH_MAP[mon_key], int(m.group(1)))

    return INVALID


def sanitize_customer_name(value: Any) -> str:
    """Higieniza o nome do cliente em Title Case limpo."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = re.sub(r"[\t\r\n]+", " ", str(value)).strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return INVALID
    return _smart_title_case(s)


def sanitize_porsche_model(value: Any) -> str:
    """Padroniza os modelos e acabamentos conforme o catálogo oficial Porsche."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = " ".join(str(value).split()).strip()
    if not s:
        return INVALID

    norm_key = re.sub(r"[^a-z0-9]", "", s.lower())
    if norm_key in _MODEL_LOOKUP:
        return _MODEL_LOOKUP[norm_key]

    # Checar se inclui o nome canônico
    for c_model in CANONICAL_PORSCHE_MODELS:
        c_norm = re.sub(r"[^a-z0-9]", "", c_model.lower())
        if c_norm == norm_key:
            return c_model

    return _smart_title_case(s)


def sanitize_model_year(value: Any) -> str:
    """Normaliza o ano do modelo para 4 dígitos (1990 <= ano <= 2035)."""
    if value is None or str(value).strip() == "":
        return INVALID

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        y = int(value)
        return str(y) if 1990 <= y <= 2035 else INVALID

    s = str(value).strip()
    if not s:
        return INVALID

    # 4 dígitos exatos
    if re.fullmatch(r"\d{4}", s):
        y = int(s)
        return str(y) if 1990 <= y <= 2035 else INVALID

    # Formatos como '20-24', '20 24', '20.24'
    m = re.fullmatch(r"(\d{2})\s*[-/.\s]\s*(\d{2})", s)
    if m:
        y = int(m.group(1) + m.group(2))
        return str(y) if 1990 <= y <= 2035 else INVALID

    # Formato por extenso
    if re.search(r"[A-Za-z]", s):
        n = _words_to_number(s)
        if n is not None:
            y = int(n)
            return str(y) if 1990 <= y <= 2035 else INVALID

    return INVALID


def sanitize_sale_price(value: Any) -> str:
    """Normaliza o preço de venda para decimal monetário com 2 casas."""
    if value is None or str(value).strip() == "":
        return INVALID

    num = _parse_numeric_value(value)
    if num is None or num <= 0:
        return INVALID

    return f"{num:.2f}"


def sanitize_vehicle_mileage(value: Any) -> str:
    """Normaliza a quilometragem para número inteiro de milhas."""
    if value is None or str(value).strip() == "":
        return INVALID

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(int(round(float(value))))

    s = str(value).strip()
    if not s:
        return INVALID

    low = s.lower()
    if low in {"new", "new car", "zero", "zero miles", "0 mi", "0 miles", "0mi", "novo", "zero km"}:
        return "0"

    is_km = bool(re.search(r"\bkm\b|kilomet", low) or low.startswith("km"))

    # Checar se é puramente textual
    if not re.search(r"\d", s):
        cleaned_text = re.sub(r"\b(mi|miles|mile|km|kilometers|kilometres|milhas)\b", " ", low).strip()
        num_word = _words_to_number(cleaned_text)
        if num_word is None:
            return INVALID
        miles = num_word * (KM_TO_MILES_FACTOR if is_km else 1.0)
        return str(int(round(miles)))

    # Numérico
    num = _parse_numeric_value(s)
    if num is None:
        return INVALID

    if is_km:
        num = num * KM_TO_MILES_FACTOR

    return str(int(round(num)))


def sanitize_payment_method(value: Any) -> str:
    """Padroniza a forma de pagamento nas categorias controladas."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID

    key = _normalize_string_key(s)
    if key in PAYMENT_NORMALIZATION_MAP:
        return PAYMENT_NORMALIZATION_MAP[key]

    return _smart_title_case(s)


def sanitize_city(value: Any) -> str:
    """Normaliza a cidade em Title Case preservando pontuações."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = re.sub(r"\s+", " ", str(value)).strip()
    if not s:
        return INVALID

    parts = s.split(" ")
    out: List[str] = []
    for p in parts:
        low = p.lower()
        if low in {"st.", "st"}:
            out.append("St.")
        elif low in {"mt.", "mt"}:
            out.append("Mt.")
        elif "-" in p:
            out.append("-".join(seg.capitalize() for seg in p.split("-")))
        else:
            out.append(p.capitalize())
    return " ".join(out)


def sanitize_state(value: Any) -> str:
    """Normaliza a sigla do estado para o padrão USPS de 2 letras."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID

    upper = s.upper()
    if upper in US_STATE_CODES:
        return upper

    norm = " ".join(s.lower().split())
    if norm in US_STATES:
        return US_STATES[norm]

    return INVALID


def sanitize_salesperson(value: Any) -> str:
    """Higieniza o nome do vendedor em Title Case."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = re.sub(r"[\t\r\n]+", " ", str(value)).strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return INVALID
    return _smart_title_case(s)


def sanitize_delivery_status(value: Any) -> str:
    """Higieniza o status de entrega com correção de pontuação e typos."""
    if value is None or str(value).strip() == "":
        return INVALID
    s = str(value).strip()
    if not s:
        return INVALID

    key = _normalize_string_key(s)
    if key in DELIVERY_STATUS_MAP:
        return DELIVERY_STATUS_MAP[key]

    return _smart_title_case(key) or INVALID


# ---------------------------------------------------------------------------
# Mapeamento do Pipeline de Colunas
# ---------------------------------------------------------------------------
COLUMN_MAPPING: List[Tuple[str, str, Callable[[Any], str]]] = [
    ("sale_id", "sale_id_sanitized", sanitize_sale_id),
    ("sale_date", "sale_date_sanitized", sanitize_date),
    ("customer_name", "customer_name_sanitized", sanitize_customer_name),
    ("porsche_model", "porsche_model_sanitized", sanitize_porsche_model),
    ("model_year", "model_year_sanitized", sanitize_model_year),
    ("sale_price", "sale_price_sanitized", sanitize_sale_price),
    ("vehicle_mileage", "vehicle_mileage_sanitized", sanitize_vehicle_mileage),
    ("payment_method", "payment_method_sanitized", sanitize_payment_method),
    ("city", "city_sanitized", sanitize_city),
    ("state", "state_sanitized", sanitize_state),
    ("salesperson", "salesperson_sanitized", sanitize_salesperson),
    ("delivery_status", "delivery_status_sanitized", sanitize_delivery_status),
]


# ---------------------------------------------------------------------------
# Processador de Planilhas e Exportação
# ---------------------------------------------------------------------------
def run_sanitization_pipeline(
    input_path: Path,
    combined_output_path: Path,
    prod_output_path: Path,
    prod_csv_path: Optional[Path] = None,
    prod_json_path: Optional[Path] = None
) -> Dict[str, Any]:
    """
    Executa a sanitização completa, salvando a planilha combinada,
    a planilha de produção otimizada, o CSV e o JSON.
    """
    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {input_path}")

    df_raw = pd.read_excel(input_path)
    total_records = len(df_raw)

    print(f"[*] Iniciando sanitização de {total_records} registros a partir de: {input_path}")

    # Validação de presença das colunas esperadas
    missing_cols = [src for src, _, _ in COLUMN_MAPPING if src not in df_raw.columns]
    if missing_cols:
        raise ValueError(f"Colunas obrigatórias ausentes na planilha de entrada: {missing_cols}")

    # Construção dos dados tratados
    sanitized_data: Dict[str, List[Any]] = {}
    invalid_stats: Dict[str, int] = {}

    for src_col, san_col, sanitizer_fn in COLUMN_MAPPING:
        clean_values = []
        invalids = 0
        for val in df_raw[src_col]:
            clean_val = sanitizer_fn(val)
            if clean_val == INVALID:
                invalids += 1
            clean_values.append(clean_val)
        sanitized_data[san_col] = clean_values
        if invalids > 0:
            invalid_stats[san_col] = invalids

    # -----------------------------------------------------------------------
    # 1. Geração da Planilha Combinada (Brutas + Sanitizadas)
    # -----------------------------------------------------------------------
    wb_combined = openpyxl.Workbook()
    ws_combined = wb_combined.active
    ws_combined.title = "Porsche Base Sanitizada"

    # Estilos Porsche
    header_raw_fill = PatternFill("solid", fgColor="1F2937")       # Cinza carvão
    header_san_fill = PatternFill("solid", fgColor="0B5345")       # Verde esmeralda escuro
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10, color="111827")
    invalid_font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    combined_headers: List[str] = []
    for src_col, san_col, _ in COLUMN_MAPPING:
        combined_headers.append(src_col)
        combined_headers.append(san_col)

    ws_combined.append(combined_headers)

    # Estilizar cabeçalho combinado
    for col_idx, h_name in enumerate(combined_headers, start=1):
        cell = ws_combined.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_san_fill if h_name.endswith("_sanitized") else header_raw_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Inserir linhas na planilha combinada
    for row_idx in range(total_records):
        row_vals = []
        for src_col, san_col, _ in COLUMN_MAPPING:
            raw_v = df_raw[src_col].iloc[row_idx]
            san_v = sanitized_data[san_col][row_idx]
            row_vals.append(raw_v)
            row_vals.append(san_v)
        ws_combined.append(row_vals)

        # Formatação de células de dados
        current_excel_row = row_idx + 2
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_combined.cell(row=current_excel_row, column=col_idx)
            cell.border = thin_border
            if val == INVALID:
                cell.font = invalid_font
                cell.fill = PatternFill("solid", fgColor="FEE2E2")
            else:
                cell.font = data_font

    # Ajuste de largura das colunas
    for col_idx in range(1, len(combined_headers) + 1):
        letter = get_column_letter(col_idx)
        ws_combined.column_dimensions[letter].width = max(16, len(combined_headers[col_idx - 1]) + 3)

    ws_combined.freeze_panes = "A2"
    wb_combined.save(combined_output_path)
    print(f"[+] Planilha combinada salva em: {combined_output_path}")

    # -----------------------------------------------------------------------
    # 2. Geração da Planilha Oficial de Produção (Apenas _sanitized)
    # -----------------------------------------------------------------------
    df_prod = pd.DataFrame(sanitized_data)

    wb_prod = openpyxl.Workbook()
    ws_prod = wb_prod.active
    ws_prod.title = "Producao"

    prod_headers = list(df_prod.columns)
    ws_prod.append(prod_headers)

    prod_header_fill = PatternFill("solid", fgColor="0F172A")  # Deep Navy Black
    prod_header_font = Font(name="Segoe UI", size=11, bold=True, color="F8FAFC")

    for col_idx, h_name in enumerate(prod_headers, start=1):
        cell = ws_prod.cell(row=1, column=col_idx)
        cell.font = prod_header_font
        cell.fill = prod_header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx in range(total_records):
        row_vals = [df_prod[c].iloc[row_idx] for c in prod_headers]
        ws_prod.append(row_vals)

        current_excel_row = row_idx + 2
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_prod.cell(row=current_excel_row, column=col_idx)
            cell.border = thin_border
            if val == INVALID:
                cell.font = invalid_font
                cell.fill = PatternFill("solid", fgColor="FEE2E2")
            else:
                cell.font = data_font

    for col_idx in range(1, len(prod_headers) + 1):
        letter = get_column_letter(col_idx)
        ws_prod.column_dimensions[letter].width = max(18, len(prod_headers[col_idx - 1]) + 3)

    ws_prod.freeze_panes = "A2"
    wb_prod.save(prod_output_path)
    print(f"[+] Planilha de produção (apenas _sanitized) salva em: {prod_output_path}")

    # -----------------------------------------------------------------------
    # 3. Exportações CSV e JSON para Ingestão Ultrarrápida no Dashboard
    # -----------------------------------------------------------------------
    if prod_csv_path:
        df_prod.to_csv(prod_csv_path, index=False, encoding="utf-8-sig")
        print(f"[+] Base de produção exportada em CSV: {prod_csv_path}")

    if prod_json_path:
        prod_records = df_prod.to_dict(orient="records")
        with open(prod_json_path, "w", encoding="utf-8") as f:
            json.dump(prod_records, f, ensure_ascii=False, indent=2)
        print(f"[+] Base de produção exportada em JSON: {prod_json_path}")

    # -----------------------------------------------------------------------
    # Estatísticas de Validação
    # -----------------------------------------------------------------------
    valid_prices = [float(p) for p in df_prod["sale_price_sanitized"] if p != INVALID]
    valid_mileage = [int(m) for m in df_prod["vehicle_mileage_sanitized"] if m != INVALID]

    stats_summary = {
        "total_records": total_records,
        "invalid_counts_per_column": invalid_stats,
        "total_revenue_usd": sum(valid_prices),
        "avg_price_usd": sum(valid_prices) / len(valid_prices) if valid_prices else 0.0,
        "avg_mileage_mi": sum(valid_mileage) / len(valid_mileage) if valid_mileage else 0.0,
        "unique_models": len(df_prod["porsche_model_sanitized"].unique()),
        "unique_cities": len(df_prod["city_sanitized"].unique()),
        "unique_payment_methods": len(df_prod["payment_method_sanitized"].unique()),
    }

    return stats_summary


# ---------------------------------------------------------------------------
# CLI Entrypoint
# ---------------------------------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(
        description="Porsche Sales Data Sanitization & Production Engine"
    )
    parser.add_argument(
        "--input", "-i",
        type=Path,
        default=DEFAULT_INPUT_PATH,
        help="Caminho para a planilha Excel bruta de entrada",
    )
    parser.add_argument(
        "--out-combined", "-c",
        type=Path,
        default=DEFAULT_COMBINED_OUTPUT,
        help="Caminho para salvar a planilha combinada (bruta + sanitizada)",
    )
    parser.add_argument(
        "--out-prod", "-p",
        type=Path,
        default=DEFAULT_PROD_OUTPUT_XLSX,
        help="Caminho para salvar a planilha de produção (apenas _sanitized)",
    )
    parser.add_argument(
        "--out-csv",
        type=Path,
        default=DEFAULT_PROD_OUTPUT_CSV,
        help="Caminho para salvar a base de produção em CSV",
    )
    parser.add_argument(
        "--out-json",
        type=Path,
        default=DEFAULT_PROD_OUTPUT_JSON,
        help="Caminho para salvar a base de produção em JSON",
    )

    args = parser.parse_args()

    try:
        stats = run_sanitization_pipeline(
            input_path=args.input,
            combined_output_path=args.out_combined,
            prod_output_path=args.out_prod,
            prod_csv_path=args.out_csv,
            prod_json_path=args.out_json,
        )

        print("\n=================================================================")
        print("RELATÓRIO EXECUTIVO DE HIGIENIZAÇÃO DE DADOS (PORSCHE)")
        print("=================================================================")
        print(f"Total de Registros Processados: {stats['total_records']}")
        print(f"Faturamento Total Calculado:   ${stats['total_revenue_usd']:,.2f}")
        print(f"Ticket Médio por Veículo:      ${stats['avg_price_usd']:,.2f}")
        print(f"Quilometragem Média:           {stats['avg_mileage_mi']:,.0f} milhas")
        print(f"Modelos Únicos Identificados:  {stats['unique_models']}")
        print(f"Cidades Atendidas:             {stats['unique_cities']}")
        print(f"Formas de Pagamento:           {stats['unique_payment_methods']}")
        print("-----------------------------------------------------------------")
        if stats["invalid_counts_per_column"]:
            print("Contagem de Valores 'INVALID' por Coluna Sanitizada:")
            for col, count in stats["invalid_counts_per_column"].items():
                print(f"  * {col}: {count} ocorrências")
        else:
            print("Nenhum valor 'INVALID' produzido. Base 100% íntegra!")
        print("=================================================================\n")
        return 0

    except Exception as e:
        print(f"[ERRO CRÍTICO] Falha no pipeline de sanitização: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
