# TermiFolha.py
#
#  ███████████                                     ███  ███████████          ████  █████
# ░█░░░███░░░█                                    ░░░  ░░███░░░░░░█         ░░███ ░░███
# ░   ░███  ░   ██████  ████████  █████████████   ████  ░███   █ ░   ██████  ░███  ░███████    ██████
#     ░███     ███░░███░░███░░███░░███░░███░░███ ░░███  ░███████    ███░░███ ░███  ░███░░███  ░░░░░███
#     ░███    ░███████  ░███ ░░░  ░███ ░███ ░███  ░███  ░███░░░█   ░███ ░███ ░███  ░███ ░███   ███████
#     ░███    ░███░░░   ░███      ░███ ░███ ░███  ░███  ░███  ░    ░███ ░███ ░███  ░███ ░███  ███░░███
#     █████   ░░██████  █████     █████░███ █████ █████ █████      ░░██████  █████ ████ █████░░████████
#    ░░░░░     ░░░░░░  ░░░░░     ░░░░░ ░░░ ░░░░░ ░░░░░ ░░░░░        ░░░░░░  ░░░░░ ░░░░ ░░░░░  ░░░░░░░░
#
#
# TermiFolha — Folha de Cálculo CLI
#
# Autor: Gabriel Gomes Alves
# Versão: 1.3
#
# Descrição: TermiFolha é uma aplicação de folha de cálculo em linha de comando, permite criar, editar, filtrar e analisar dados de uma folha de cálculo (DataFrame) de forma simples e eficiente.
#
# Funcionalidades:
# - Várias folhas (separadores); juntar, empilhar e tabela dinâmica
# - Criar nova folha de cálculo com colunas personalizadas
# - Adicionar, editar e remover linhas e colunas; procurar e substituir
# - Colunas calculadas por fórmula (recalculadas automaticamente), regras de validação, conversão de tipos
# - Ordenar dados por múltiplas colunas; filtrar com critérios avançados; ver por páginas
# - Estatísticas descritivas, correlação, outliers, normalização e faixas
# - Visualizações (histogramas, linhas, heatmaps, ...) no ecrã ou gravadas em PNG
# - Suporte a tipos de dados: texto, inteiro, decimal e data
# - Importar Excel, CSV, JSON e texto colado; exportar Excel, CSV, JSON, Markdown, HTML, Word e PDF
# - Projectos (.termifolha), Undo/Redo configurável e guardar automático do estado (JSON)
# - Modo de linha de comandos (--importar / --exportar) sem menus
#
# Requisitos: Python 3.12+, numpy, pandas, matplotlib, openpyxl, python-docx, fpdf2 (ver requirements.txt)
#
# Uso: Execute o script e siga as instruções no terminal para interagir com a folha de cálculo.
#
# Nota: Este código é um exemplo de implementação e pode ser expandido com mais funcionalidades ou melhorias na interface.
#
# Copyright (c) 2026 Gabriel Gomes Alves.
#
# Licença: MIT License (https://opensource.org/licenses/MIT)
#
# (ASCII art gerado por https://patorjk.com/software/taag/#p=display&f=DOS+Rebel&t=TermiFolha&x=none&v=4&h=4&w=80&we=false)
#
# 13-11-2026

import argparse
import ast
import copy
import csv
import html
import io
import json
import logging
import operator
import os
import pickle
import re
import sys
import time

import fpdf as fpdf_lib
import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from fpdf import FPDF

# Constantes da aplicação (nome, versão, ficheiros, dimensões UI)
APP_NAME          = "TermiFolha"
APP_TITLE        = "TermiFolha  —  Folha de Cálculo CLI"
APP_VERSION        = "1.3"
APP_AUTHOR         = "Gabriel Gomes Alves"
_PROGRAM_DIR   = os.path.dirname(os.path.abspath(__file__))
STATE_FILE   = os.path.join(_PROGRAM_DIR, ".termifolha_estado.json")
STATE_ENABLED = True  # False no modo não interactivo: a sessão guardada não é lida nem escrita
LEGACY_STATE_FILE = os.path.join(_PROGRAM_DIR, ".termifolha_estado.pkl")  # formato antigo, só lido para migrar
LOG_FILE      = os.path.join(_PROGRAM_DIR, ".termifolha.log")
EXT_EXCEL         = ".xlsx"
EXT_CSV           = ".csv"
EXT_WORD          = ".docx"
EXT_PDF           = ".pdf"
WIDTH           = 72
COLUMN_WIDTH_LIMIT        = 15
DATE_TYPE         = "date"  # sentinel para colunas data (pandas sem dtype Python nativo) - https://pandas.pydata.org/docs/user_guide/missing_data.html

TYPE_CHOICES             = {
    "1": (str,       "Texto"),
    "2": (int,       "Inteiro"),
    "3": (float,     "Decimal"),
    "4": (DATE_TYPE, "Data (DD/MM/AAAA)"),
}
TYPE_NAME         = {col_type: name.lower() for col_type, name in TYPE_CHOICES.values()}
HISTORY_LIMIT       = 20  # valor por omissão (configurável em Definições)
HISTORY_LIMIT_MAX   = 100
PAGE_SIZE           = 20  # linhas por página em "Ver por páginas"
DISPLAY_ROW_LIMIT   = 30  # linhas mostradas no ecrã principal (as restantes vêem-se por páginas)
DEFAULT_SHEET_NAME  = "Folha1"
EXT_PROJECT         = ".termifolha"
UI_INDENT_1   = "  "
UI_INDENT_2   = "    "

log = logging.getLogger("termifolha")  # sem handlers por omissão; main() configura o ficheiro de registo
log.addHandler(logging.NullHandler())

# Estado global: livro de folhas (a folha activa vive em `table`/`col_types`/`meta`) e histórico undo/redo
table = pd.DataFrame()
col_types = {}
meta = {}  # metadados da folha activa (preenchido abaixo por new_meta)
sheets = {DEFAULT_SHEET_NAME: None}  # nome -> folha inactiva; a activa é None (vive nas variáveis acima). A ordem é a dos separadores.
active_name = DEFAULT_SHEET_NAME
settings = {"history_limit": HISTORY_LIMIT, "page_size": PAGE_SIZE}

undo_stack = []
redo_stack = []


def new_meta():
    """Metadados de uma folha: regras de validação, formatos de apresentação, colunas visíveis, fórmulas, linha de totais e textos dos gráficos."""
    return {"rules": {}, "formats": {}, "visible": [], "formulas": {}, "totals": None, "charts": {}}


meta = new_meta()


# Utilitários de terminal e formatação de valores
def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def pause():
    input(f"\n{UI_INDENT_1}Prima [ENTER] para continuar...")


def truncate(value, max_width):
    s = str(value)
    return s if len(s) <= max_width else s[:max_width - 1] + "~"


def format_value(val):
    if pd.isna(val):
        return "--"
    if hasattr(val, "strftime"):
        return val.strftime("%d/%m/%Y")
    return str(val)


def format_number(val):
    if pd.isna(val):
        return "--"
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.4f}".rstrip("0").rstrip(".")
    if isinstance(val, int):
        return str(val)
    return str(val)


# Formatação de valores para apresentação (não altera os dados)
FORMAT_DECIMALS_MAX = 10


def apply_number_format(value, fmt):
    """Formata um número segundo `fmt`: decimals (0-10), thousands (bool), locale ("intl" 1,234.50 | "pt" 1 234,50), prefix, suffix."""
    decimals = fmt.get("decimals", 2)
    if not isinstance(decimals, int) or isinstance(decimals, bool) or not 0 <= decimals <= FORMAT_DECIMALS_MAX:
        decimals = 2
    text = f"{value:,.{decimals}f}" if fmt.get("thousands") else f"{value:.{decimals}f}"
    if fmt.get("locale") == "pt":
        text = text.replace(",", "\x00").replace(".", ",").replace("\x00", " ")
    return f"{fmt.get('prefix', '')}{text}{fmt.get('suffix', '')}"


def format_cell(col, val):
    """Valor de uma célula tal como aparece no ecrã: aplica o formato da coluna (se existir) aos números."""
    fmt = meta["formats"].get(col)
    if fmt and not pd.isna(val) and isinstance(val, (int, float, np.integer, np.floating)) and not isinstance(val, bool):
        try:
            return apply_number_format(val, fmt)
        except (TypeError, ValueError):
            log.warning("Formato inválido na coluna %r: %r", col, fmt)
    return format_value(val)


# Desenho da tabela e cabeçalho no terminal
def column_widths(data):
    widths = {}
    for col in data.columns:
        w = len(str(col))
        if not data.empty:
            for idx in data.index:
                disp_len = len(format_cell(col, data.at[idx, col]))
                if disp_len > w:
                    w = disp_len
        widths[col] = min(w, COLUMN_WIDTH_LIMIT)
    return widths


def draw_data_table(data, totals=None):
    """Desenha `data`; `totals` (coluna -> texto) acrescenta uma linha de totais sob a tabela."""
    if data.columns.empty:
        print(f"{UI_INDENT_1}(tabela vazia — crie uma folha primeiro)")
        return

    widths = column_widths(data)
    if totals:
        for col in data.columns:
            widths[col] = max(widths[col], min(len(totals.get(col, "")), COLUMN_WIDTH_LIMIT))
    index_width = max(3, len(str(len(data))))

    head_line = UI_INDENT_1 + "#".rjust(index_width) + " |"
    sep = UI_INDENT_1 + "-" * index_width + "-+"
    for col in data.columns:
        w = widths[col]
        head_line += " " + truncate(col, w).ljust(w) + " |"
        sep += "-" * w + "--+"

    print(head_line)
    print(sep)

    if data.empty:
        print(UI_INDENT_1 + " " * index_width + f"{UI_INDENT_1}(sem linhas)")
    for idx in data.index:
        line = UI_INDENT_1 + str(idx).rjust(index_width) + " |"
        for col in data.columns:
            shown = truncate(format_cell(col, data.at[idx, col]), widths[col])
            line += " " + shown.ljust(widths[col]) + " |"
        print(line)

    if totals:
        print(sep)
        line = UI_INDENT_1 + "Σ".rjust(index_width) + " |"
        for col in data.columns:
            line += " " + truncate(totals.get(col, ""), widths[col]).ljust(widths[col]) + " |"
        print(line)


def visible_columns():
    """Colunas a mostrar, pela ordem escolhida (todas se não houver escolha)."""
    chosen = [c for c in meta["visible"] if c in table.columns]
    return chosen or list(table.columns)


_TOTALS_LABELS = {"sum": "soma", "mean": "média", "count": "contagem", "min": "mínimo", "max": "máximo"}


def totals_row(columns):
    """Texto da linha de totais por coluna, ou None se não estiver activa. Calcula sobre todas as linhas da folha."""
    function = meta["totals"]
    if function not in _TOTALS_FUNCTIONS:
        return None
    result = {}
    for col in columns:
        series = table[col]
        if function == "count":
            result[col] = str(int(series.notna().sum()))
        elif pd.api.types.is_numeric_dtype(series) and series.notna().any():
            value = getattr(series, function)()
            result[col] = format_cell(col, value) if col in meta["formats"] else format_number(value)
        else:
            result[col] = ""
    return result


def page_bounds(total, size, page):
    """(início, fim, nº de páginas) da página `page` (0-based, limitada ao intervalo válido) de `total` linhas."""
    pages = max(1, -(-total // size))
    page = min(max(page, 0), pages - 1)
    return page * size, min(total, (page + 1) * size), pages


def draw_table():
    view = table[visible_columns()] if not table.columns.empty else table
    hidden = max(0, len(view) - DISPLAY_ROW_LIMIT)
    draw_data_table(view.head(DISPLAY_ROW_LIMIT), totals_row(list(view.columns)))
    if hidden:
        print(f"{UI_INDENT_1}... mais {hidden} linha(s) por mostrar (Análise Básica → Ver por páginas)")
    if len(view.columns) < len(table.columns):
        print(f"{UI_INDENT_1}({len(table.columns) - len(view.columns)} coluna(s) oculta(s); Colunas → Colunas visíveis)")


def header(subtitle="MENU PRINCIPAL", show_table=True):
    clear_screen()
    print("=" * WIDTH)
    print(APP_TITLE.center(WIDTH))
    print("=" * WIDTH)
    if table.columns.empty:
        state_text = "folha vazia"
    else:
        state_text = f"{len(table.columns)} coluna(s)  |  {len(table)} linha(s)"
    names = sheet_names()
    if len(names) > 1:
        state_text += f"  |  folha '{active_name}' ({names.index(active_name) + 1}/{len(names)})"
    print(f"{UI_INDENT_1}Estado: {state_text}")
    print("-" * WIDTH)
    if show_table:
        print()
        draw_table()
    print()
    print("=" * WIDTH)
    print(subtitle.center(WIDTH))
    print("-" * WIDTH)


def print_pandas_table(obj, prefix=UI_INDENT_2):
    for line in obj.to_string().split("\n"):
        print(f"{prefix}{line}")


# Entrada do utilizador: opções, tipos, conversão e prompts
def ask_option(valid_cols):
    while True:
        choice = input(f"{UI_INDENT_1}Opção: ").strip().upper()
        if choice in valid_cols:
            return choice
        print(f"{UI_INDENT_1}Opção inválida.")


def ask_type():
    options = "  ".join(f"[{k}] {name}" for k, (_, name) in TYPE_CHOICES.items())
    while True:
        print(f"{UI_INDENT_1}Tipo: {options}")
        t = input(f"{UI_INDENT_1}Tipo: ").strip()
        if t in TYPE_CHOICES:
            return TYPE_CHOICES[t][0]
        print(f"{UI_INDENT_1}Tipo inválido.")


def convert_value(value, col_type):
    if col_type == DATE_TYPE:
        return pd.to_datetime(value.strip(), format="%d/%m/%Y")
    return col_type(value.strip())


def resolve_column(name_input, source=None):
    source = table.columns if source is None else source
    name_lower = name_input.lower()
    for col in source:
        if str(col).lower() == name_lower:
            return col
    return None


def confirm(message):
    reply = input(f"{UI_INDENT_1}{message} [S/N]: ").strip().lower()
    return reply in ("s", "sim", "y", "yes")


def ask_column(message="Coluna"):
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    value = input(f"{UI_INDENT_1}{message}: ").strip()
    if value.lower() == "sair":
        return None
    col = resolve_column(value)
    if col is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return None
    return col


_ALL = object()  # sentinela: "todas as colunas"


def ask_column_or_all(message="Coluna"):
    """Como pedir_coluna, mas Enter significa todas as colunas (devolve _TODAS). None se cancelar."""
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    value = input(f"{UI_INDENT_1}{message} (Enter = todas): ").strip()
    if value.lower() == "sair":
        return None
    if not value:
        return _ALL
    col = resolve_column(value)
    if col is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return None
    return col


def ask_numeric_column(message="Coluna numérica"):
    num_cols = table.select_dtypes(include="number").columns.tolist()
    if not num_cols:
        print(f"{UI_INDENT_1}Sem colunas numéricas.")
        pause()
        return None
    print(f"{UI_INDENT_1}Colunas numéricas: {', '.join(num_cols)}")
    value = input(f"{UI_INDENT_1}{message}: ").strip()
    if value.lower() == "sair":
        return None
    col = resolve_column(value, source=num_cols)
    if col is None:
        print(f"{UI_INDENT_1}Coluna inválida ou não numérica.")
        pause()
        return None
    return col


def ask_numeric_columns(min_n=1, max_n=None, message="Colunas numéricas (separadas por vírgula)"):
    num_cols = table.select_dtypes(include="number").columns.tolist()
    if len(num_cols) < min_n:
        print(f"{UI_INDENT_1}Necessário pelo menos {min_n} coluna(s) numérica(s).")
        pause()
        return None

    valid_cols = [c for c in num_cols if table[c].notna().any()]
    if len(valid_cols) < min_n:
        print(f"{UI_INDENT_1}Necessário pelo menos {min_n} coluna(s) numérica(s) com dados.")
        pause()
        return None

    print(f"{UI_INDENT_1}Colunas numéricas disponíveis: {', '.join(valid_cols)}")
    if max_n != 1:
        print(f"{UI_INDENT_1}(pode indicar várias colunas separadas por vírgula)")
    value = input(f"{UI_INDENT_1}{message}: ").strip()
    if value.lower() == "sair":
        return None

    requested = [v.strip() for v in value.split(",") if v.strip()]
    if not requested:
        print(f"{UI_INDENT_1}Deve indicar pelo menos uma coluna.")
        pause()
        return None

    chosen = []
    for name in requested:
        col = resolve_column(name, source=valid_cols)
        if col is None:
            print(f"{UI_INDENT_1}Coluna inválida ou sem dados numéricos: '{name}'.")
            pause()
            return None
        if col in chosen:
            print(f"{UI_INDENT_1}Coluna repetida: '{col}'.")
            pause()
            return None
        chosen.append(col)

    if len(chosen) < min_n:
        print(f"{UI_INDENT_1}Deve seleccionar pelo menos {min_n} coluna(s).")
        pause()
        return None
    if max_n is not None and len(chosen) > max_n:
        print(f"{UI_INDENT_1}Deve seleccionar no máximo {max_n} coluna(s).")
        pause()
        return None
    return chosen


def ask_int_in_range(message, minimum, maximum, default=None):
    value = input(message).strip()
    if not value:
        return default
    try:
        number = int(value)
    except ValueError:
        print(f"{UI_INDENT_1}Valor inválido (inteiro esperado).")
        pause()
        return None
    if number < minimum or number > maximum:
        print(f"{UI_INDENT_1}Valor fora do intervalo [{minimum}, {maximum}].")
        pause()
        return None
    return number


def type_of(col):
    return col_types.get(col, str)


_DTYPE_INFER = [
    (pd.api.types.is_integer_dtype,        int),
    (pd.api.types.is_float_dtype,          float),
    (pd.api.types.is_datetime64_any_dtype, DATE_TYPE),
]


def infer_type(series):
    for check, col_type in _DTYPE_INFER:
        if check(series):
            return col_type
    return str


# Aplicação de tipos, persistência em ficheiro e histórico undo/redo
def typed_frame(data, types):
    """Devolve uma cópia de `data` com cada coluna convertida para o tipo declarado (as que não convertem ficam como estão)."""
    data = data.copy()
    for col in data.columns:
        col_type = types.get(col, str)
        try:
            if col_type is int:
                data[col] = pd.to_numeric(data[col], errors="coerce").astype("Int64")
            elif col_type is float:
                data[col] = pd.to_numeric(data[col], errors="coerce")
            elif col_type == DATE_TYPE:
                data[col] = pd.to_datetime(data[col], errors="coerce")
        except Exception:
            log.debug("Coluna %r não convertida para %r", col, col_type, exc_info=True)
    return data


def apply_types():
    global table
    table = typed_frame(table, col_types)


# Persistência: o estado é gravado em JSON (nunca em pickle, que executa código ao carregar).
_JSON_TYPE_NAME = {str: "str", int: "int", float: "float", DATE_TYPE: "date"}
_TYPE_FROM_JSON_NAME = {name: col_type for col_type, name in _JSON_TYPE_NAME.items()}
_STATE_VERSION = 2
_SUPPORTED_STATE_VERSIONS = (1, 2)  # a versão 1 (uma só folha, sem metadados) é lida e convertida
_TOTALS_FUNCTIONS = ("sum", "mean", "count", "min", "max")


def _json_value(v):
    if pd.isna(v):
        return None
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if hasattr(v, "item"):
        return v.item()
    return v


def sheet_names():
    """Nomes das folhas pela ordem dos separadores (a activa está sempre incluída)."""
    names = list(sheets)
    if active_name not in names:
        names.append(active_name)
    return names


def get_sheet(name):
    """Dados da folha `name` ({"table", "types", "meta"}); para a activa devolve as variáveis vivas."""
    if name == active_name:
        return {"table": table, "types": col_types, "meta": meta}
    return sheets[name]


def _sheet_to_dict(name, sheet):
    data, types, sheet_meta = sheet["table"], sheet["types"], sheet["meta"]
    return {
        "nome": name,
        "colunas": [str(c) for c in data.columns],
        "tipos": {str(c): _JSON_TYPE_NAME.get(t, "str") for c, t in types.items()},
        "dados": [[_json_value(v) for v in record] for record in data.itertuples(index=False, name=None)],
        "regras": sheet_meta["rules"],
        "formatos": sheet_meta["formats"],
        "visiveis": sheet_meta["visible"],
        "formulas": sheet_meta["formulas"],
        "totais": sheet_meta["totals"],
        "graficos": sheet_meta["charts"],
    }


def _state_to_dict():
    return {
        "versao": _STATE_VERSION,
        "activa": active_name,
        "definicoes": dict(settings),
        "folhas": [_sheet_to_dict(name, get_sheet(name)) for name in sheet_names()],
    }


def _sheet_from_dict(d, name):
    """Valida e reconstrói (nome, tabela, tipos, metadados) de uma folha. Levanta ValueError se for inválida."""
    if not isinstance(d, dict):
        raise ValueError("folha inválida")
    column_names, type_names, data_rows = d.get("colunas"), d.get("tipos"), d.get("dados")
    if not (isinstance(column_names, list) and all(isinstance(c, str) for c in column_names)):
        raise ValueError("lista de colunas inválida")
    if not (isinstance(type_names, dict) and isinstance(data_rows, list)):
        raise ValueError("estrutura inválida")
    if any(not isinstance(r, list) or len(r) != len(column_names) for r in data_rows):
        raise ValueError("linha com número de valores diferente do número de colunas")
    new_col_types = {c: _TYPE_FROM_JSON_NAME.get(type_names.get(c), str) for c in column_names}
    df = pd.DataFrame(data_rows, columns=column_names) if data_rows else pd.DataFrame(columns=column_names)

    sheet_meta = new_meta()
    for key, json_key in (("rules", "regras"), ("formats", "formatos"), ("formulas", "formulas")):
        value = d.get(json_key, {})
        if not isinstance(value, dict):
            raise ValueError(f"'{json_key}' inválido")
        sheet_meta[key] = {c: v for c, v in value.items() if c in column_names}
    if any(not isinstance(v, dict) for key in ("rules", "formats") for v in sheet_meta[key].values()):
        raise ValueError("regras ou formatos inválidos")
    if any(not isinstance(v, str) for v in sheet_meta["formulas"].values()):
        raise ValueError("fórmulas inválidas")
    visible = d.get("visiveis", [])
    if not (isinstance(visible, list) and all(isinstance(c, str) for c in visible)):
        raise ValueError("colunas visíveis inválidas")
    sheet_meta["visible"] = [c for c in visible if c in column_names]
    totals = d.get("totais")
    sheet_meta["totals"] = totals if totals in _TOTALS_FUNCTIONS else None
    sheet_meta["charts"] = _sanitize_chart_labels(d.get("graficos"))
    return name, df, new_col_types, sheet_meta


def _sanitize_settings(raw):
    """Aceita só definições válidas; o resto fica com o valor por omissão."""
    raw = raw if isinstance(raw, dict) else {}
    clean = {"history_limit": HISTORY_LIMIT, "page_size": PAGE_SIZE}
    limit, page = raw.get("history_limit"), raw.get("page_size")
    if isinstance(limit, int) and not isinstance(limit, bool) and 1 <= limit <= HISTORY_LIMIT_MAX:
        clean["history_limit"] = limit
    if isinstance(page, int) and not isinstance(page, bool) and 5 <= page <= 200:
        clean["page_size"] = page
    return clean


def _state_from_dict(d):
    """Valida e reconstrói o livro: devolve (lista de (nome, tabela, tipos, metadados), nome da activa, definições)."""
    if not isinstance(d, dict) or d.get("versao") not in _SUPPORTED_STATE_VERSIONS:
        raise ValueError("versão de estado desconhecida")
    if d["versao"] == 1:
        return [_sheet_from_dict(d, DEFAULT_SHEET_NAME)], DEFAULT_SHEET_NAME, _sanitize_settings(None)
    raw_sheets = d.get("folhas")
    if not (isinstance(raw_sheets, list) and raw_sheets):
        raise ValueError("lista de folhas inválida")
    sheet_list = []
    for raw in raw_sheets:
        name = raw.get("nome") if isinstance(raw, dict) else None
        if not isinstance(name, str) or sheet_name_error(name, [n for n, *_ in sheet_list]):
            raise ValueError("nome de folha inválido ou repetido")
        sheet_list.append(_sheet_from_dict(raw, name))
    active = d.get("activa")
    if active not in [n for n, *_ in sheet_list]:
        active = sheet_list[0][0]
    return sheet_list, active, _sanitize_settings(d.get("definicoes"))


def _install_workbook(sheet_list, active, new_settings):
    """Substitui todo o livro (folhas, activa, definições) pelo conteúdo já validado."""
    global table, col_types, meta, sheets, active_name
    new_sheets, live = {}, None
    for name, df, types, sheet_meta in sheet_list:
        sheet = {"table": typed_frame(df, types), "types": types, "meta": sheet_meta}
        if name == active:
            live, new_sheets[name] = sheet, None
        else:
            new_sheets[name] = sheet
    table, col_types, meta = live["table"], live["types"], live["meta"]
    sheets, active_name = new_sheets, active
    settings.update(new_settings)


def _write_state_file(path):
    """Grava o livro em JSON de forma atómica (ficheiro temporário + os.replace). Levanta OSError/TypeError se falhar."""
    temp_path = path + ".tmp"
    try:
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(_state_to_dict(), f, ensure_ascii=False)
        os.replace(temp_path, path)
    except BaseException:
        try:
            os.remove(temp_path)
        except OSError:
            pass
        raise


def save_state():
    """Grava o estado de forma atómica. Devolve True se gravou, False se falhou (e avisa o utilizador)."""
    if not STATE_ENABLED:
        return True
    try:
        _write_state_file(STATE_FILE)
        return True
    except Exception as e:
        log.exception("Falha ao guardar o estado")
        print(f"{UI_INDENT_1}Aviso: não foi possível guardar o estado ({e}).")
        print(f"{UI_INDENT_1}As alterações mantêm-se em memória; exporte para Excel para não as perder.")
        return False


# Tipos que um estado antigo (pickle) pode legitimamente referenciar. Tudo o resto é recusado.
_PICKLE_BUILTINS = {"str", "int", "float", "bool", "dict", "list", "tuple", "set", "frozenset", "slice", "object", "complex", "bytes", "range"}
_PICKLE_NUMPY = {"_reconstruct", "ndarray", "dtype", "_frombuffer", "scalar"}
_PICKLE_PANDAS_TOPO = {  # nunca "eval", "read_pickle", ...
    "DataFrame", "Series", "Index", "RangeIndex", "MultiIndex", "DatetimeIndex",
    "StringDtype", "Int64Dtype", "Float64Dtype", "BooleanDtype", "DatetimeTZDtype", "CategoricalDtype",
}


class _SafeUnpickler(pickle.Unpickler):
    """Só deixa reconstruir DataFrames: recusa qualquer outro objecto (os.system, pandas.eval, ...)."""

    def find_class(self, module, name):
        if "." in name:
            raise pickle.UnpicklingError(f"referência recusada: {module}.{name}")
        if module == "builtins" and name in _PICKLE_BUILTINS:
            return super().find_class(module, name)
        if module.startswith("numpy") and name in _PICKLE_NUMPY:
            return super().find_class(module, name)
        if (module == "pandas" and name in _PICKLE_PANDAS_TOPO) or module == "pandas.arrays":
            return super().find_class(module, name)
        if (module.startswith("pandas.core.") or module.startswith("pandas._libs.")) and not module.startswith("pandas.core.computation"):
            return super().find_class(module, name)
        raise pickle.UnpicklingError(f"referência recusada: {module}.{name}")


def _preserve_invalid_file(path, error):
    backup_path = path + ".corrompido"
    log.error("Estado inválido em %s: %s", path, error)
    print(f"{UI_INDENT_1}Aviso: o estado guardado está corrompido ({error}).")
    try:
        os.replace(path, backup_path)
        print(f"{UI_INDENT_1}O ficheiro foi preservado como '{backup_path}'. A começar com uma folha vazia.")
    except OSError as backup_error:
        print(f"{UI_INDENT_1}Não foi possível preservar o ficheiro ({backup_error}).")


def _read_state_file(path):
    """Lê e valida um ficheiro de estado/projecto em JSON. Levanta se for ilegível ou inválido."""
    with open(path, encoding="utf-8") as f:
        return _state_from_dict(json.load(f))


def _read_legacy_pickle(path):
    """Lê o estado da v1.0 (pickle) com um Unpickler restrito e devolve-o no formato do livro."""
    with open(path, "rb") as f:
        old_table, old_col_types = _SafeUnpickler(f).load()
    if not isinstance(old_table, pd.DataFrame) or not isinstance(old_col_types, dict):
        raise ValueError("conteúdo inesperado")
    old_table = old_table.copy()
    old_table.columns = [str(c) for c in old_table.columns]
    types = {str(c): t for c, t in old_col_types.items()}
    return [(DEFAULT_SHEET_NAME, old_table, types, new_meta())], DEFAULT_SHEET_NAME, _sanitize_settings(None)


def load_state():
    if os.path.exists(STATE_FILE):
        path, reader, migrated = STATE_FILE, _read_state_file, False
    elif os.path.exists(LEGACY_STATE_FILE):
        path, reader, migrated = LEGACY_STATE_FILE, _read_legacy_pickle, True
    else:
        return

    try:
        loaded = reader(path)
    except Exception as e:  # noqa: BLE001 - qualquer falha ao ler o estado deve ser avisada, nunca fechar o programa
        _preserve_invalid_file(path, e)
        reset_workbook()
        pause()
        return

    _install_workbook(*loaded)
    if migrated and save_state():
        try:
            os.replace(LEGACY_STATE_FILE, LEGACY_STATE_FILE + ".migrado")
            print(f"{UI_INDENT_1}Estado antigo migrado para o novo formato (cópia em '{LEGACY_STATE_FILE}.migrado').")
            pause()
        except OSError as e:
            log.warning("Não foi possível renomear o estado antigo: %s", e)
    undo_stack.clear()
    redo_stack.clear()
    undo_stack.append(snapshot())


def reset_workbook():
    """Volta a um livro com uma única folha vazia (as definições mantêm-se)."""
    global table, col_types, meta, sheets, active_name
    table, col_types, meta = pd.DataFrame(), {}, new_meta()
    sheets, active_name = {DEFAULT_SHEET_NAME: None}, DEFAULT_SHEET_NAME


# Folhas (separadores): nomes, criação, troca, renomear, eliminar, duplicar
_SHEET_NAME_FORBIDDEN = set("[]:*?/\\")
_SHEET_NAME_MAX = 31  # limite do Excel, para que a exportação nunca falhe por causa do nome


def sheet_name_error(name, existing=None):
    """Mensagem de erro se `name` não servir como nome de folha; None se for válido."""
    existing = sheet_names() if existing is None else existing
    if not name.strip():
        return "o nome não pode ser vazio"
    if name != name.strip():
        return "o nome não pode começar nem acabar com espaços"
    if len(name) > _SHEET_NAME_MAX:
        return f"o nome tem mais de {_SHEET_NAME_MAX} caracteres"
    if any(c in _SHEET_NAME_FORBIDDEN for c in name):
        return "o nome não pode conter [ ] : * ? / \\"
    if any(n.lower() == name.lower() for n in existing):
        return "já existe uma folha com esse nome"
    return None


def unique_sheet_name(base):
    """Devolve `base` ou `base (2)`, `base (3)`... conforme os nomes já usados (cortado para caber no limite)."""
    base = "".join("_" if c in _SHEET_NAME_FORBIDDEN else c for c in base).strip()[:_SHEET_NAME_MAX - 5] or "Folha"
    candidate, n = base, 2
    while sheet_name_error(candidate):
        candidate = f"{base} ({n})"
        n += 1
    return candidate


def switch_sheet(name):
    """Torna `name` a folha activa. As variáveis `table`/`col_types`/`meta` passam a apontar para ela."""
    global table, col_types, meta, active_name
    if name == active_name:
        return
    target = sheets[name]
    sheets[active_name] = {"table": table, "types": col_types, "meta": meta}
    table, col_types, meta = target["table"], target["types"], target["meta"]
    sheets[name] = None
    active_name = name


def add_sheet(name, data=None, types=None, sheet_meta=None, activate=True):
    """Acrescenta uma folha ao livro (vazia se não houver dados). Levanta ValueError se o nome for inválido."""
    error = sheet_name_error(name)
    if error:
        raise ValueError(error)
    sheets[name] = {
        "table": pd.DataFrame() if data is None else data,
        "types": {} if types is None else types,
        "meta": new_meta() if sheet_meta is None else sheet_meta,
    }
    if activate:
        switch_sheet(name)


def rename_sheet(old, new):
    """Renomeia uma folha mantendo a posição. Levanta ValueError se o novo nome for inválido."""
    global sheets, active_name
    others = [n for n in sheet_names() if n != old]
    error = sheet_name_error(new, others)
    if error:
        raise ValueError(error)
    sheets = {(new if n == old else n): v for n, v in sheets.items()}
    if old == active_name:
        active_name = new


def delete_sheet(name):
    """Elimina uma folha (nunca a última). Se for a activa, passa a activa a vizinha."""
    names = sheet_names()
    if len(names) <= 1:
        raise ValueError("o livro tem de ter pelo menos uma folha")
    if name == active_name:
        position = names.index(name)
        switch_sheet(names[position + 1] if position + 1 < len(names) else names[position - 1])
    del sheets[name]


def duplicate_sheet(name, new_name):
    """Copia a folha `name` para uma nova folha `new_name` (fica activa)."""
    source = get_sheet(name)
    copy_ = {"table": source["table"].copy(), "types": dict(source["types"]), "meta": copy.deepcopy(source["meta"])}
    add_sheet(new_name, copy_["table"], copy_["types"], copy_["meta"])


# Histórico (desfazer/refazer): cada ponto guarda o livro inteiro
def snapshot():
    """Cópia independente do livro (folha activa, outras folhas e qual é a activa)."""
    def copy_sheet(sheet):
        return {"table": sheet["table"].copy(), "types": dict(sheet["types"]), "meta": copy.deepcopy(sheet["meta"])}

    return {
        "live": copy_sheet({"table": table, "types": col_types, "meta": meta}),
        "sheets": {n: (None if sh is None else copy_sheet(sh)) for n, sh in sheets.items()},
        "active": active_name,
    }


def restore_snapshot(snap):
    """Repõe o livro guardado em `snap` (trabalha sobre uma cópia, o ponto fica intacto para desfazer/refazer)."""
    global table, col_types, meta, sheets, active_name
    live = snap["live"]
    table, col_types, meta = live["table"].copy(), dict(live["types"]), copy.deepcopy(live["meta"])
    sheets = {
        n: (None if sh is None else {"table": sh["table"].copy(), "types": dict(sh["types"]), "meta": copy.deepcopy(sh["meta"])})
        for n, sh in snap["sheets"].items()
    }
    active_name = snap["active"]


def record_checkpoint():
    redo_stack.clear()
    undo_stack.append(snapshot())
    del undo_stack[:max(0, len(undo_stack) - (settings["history_limit"] + 1))]


def commit():
    """Fecha uma alteração: recalcula as fórmulas, grava o estado e regista o ponto de desfazer."""
    recalculate_formulas()
    save_state()
    record_checkpoint()


def undo_last():
    if len(undo_stack) <= 1:
        print(f"{UI_INDENT_1}Nada para desfazer.")
        pause()
        return
    redo_stack.append(undo_stack.pop())
    restore_snapshot(undo_stack[-1])
    save_state()
    print(f"{UI_INDENT_1}Desfeito. ({len(undo_stack) - 1} passo(s) atrás)")
    pause()


def redo_last():
    if not redo_stack:
        print(f"{UI_INDENT_1}Nada para refazer.")
        pause()
        return
    state = redo_stack.pop()
    undo_stack.append(state)
    restore_snapshot(state)
    save_state()
    print(f"{UI_INDENT_1}Refeito. ({len(redo_stack)} passo(s) para refazer)")
    pause()


# Gestão da folha: criar, importar, exportar, dummy, limpar
def new_sheet():
    global table, col_types, meta
    header("NOVA FOLHA")
    print(f"{UI_INDENT_1}Defina as colunas (nome vazio ou 'sair' para terminar | 'sair' sem colunas para cancelar)\n")

    column_names = []
    new_col_types = {}

    while True:
        name = input(f"{UI_INDENT_1}Nome da coluna {len(column_names) + 1}: ").strip()
        if name.lower() == "sair":
            if not column_names:
                print(f"\n{UI_INDENT_1}Operação cancelada.")
                pause()
                return
            break
        if not name:
            if not column_names:
                print(f"{UI_INDENT_1}Precisa de pelo menos uma coluna.")
                continue
            break
        if any(c.lower() == name.lower() for c in column_names):
            print(f"{UI_INDENT_1}Coluna '{name}' já existe. Por favor escolha um nome diferente.")
            continue
        new_col_types[name] = ask_type()
        column_names.append(name)
        print(f"{UI_INDENT_1}+ coluna '{name}' do tipo '{TYPE_NAME.get(new_col_types[name], 'texto')}' adicionada.\n")

    table = pd.DataFrame(columns=column_names)
    col_types = new_col_types
    meta = new_meta()
    commit()
    print(f"\n{UI_INDENT_1}Folha criada com {len(column_names)} coluna(s).")
    pause()


def confirm_replace_sheet():
    """Se já existir folha, pede confirmação antes de a substituir. Devolve False se o utilizador recusar."""
    if table.columns.empty or confirm("Folha actual será substituída. Continuar?"):
        return True
    print(f"{UI_INDENT_1}Operação cancelada.")
    pause()
    return False


def import_excel():
    header("IMPORTAR EXCEL")
    print()
    if not confirm_replace_sheet():
        return
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return

    if not name.endswith(EXT_EXCEL):
        name += EXT_EXCEL

    try:
        frames = pd.read_excel(name, sheet_name=None)
        if len(frames) == 1:
            adopt_dataframe(next(iter(frames.values())), ask_integers=True)
        else:
            adopt_workbook(frames, ask_integers=True)
        suffix = f" ({len(frames)} folhas)" if len(frames) > 1 else ""
        print(f"\n{UI_INDENT_1}'{name}' importado{suffix}: {len(table)} linhas, {len(table.columns)} colunas na folha activa.")
    except FileNotFoundError:
        print(f"\n{UI_INDENT_1}Ficheiro '{name}' não encontrado.")
    except Exception as e:
        log.exception("Falha em import_excel")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def _integers_with_gaps(series):
    """True se for decimal com células vazias e todos os valores presentes forem inteiros (ex.: 25.0, 40.0, NaN)."""
    valid_values = series.dropna()
    return pd.api.types.is_float_dtype(series) and 0 < len(valid_values) < len(series) and bool((valid_values % 1 == 0).all())


def _prepare_dataframe(df, detect_dates=False, ask_integers=False):
    """Normaliza um DataFrame importado e infere os tipos das colunas. Devolve (df, tipos); não toca no estado global."""
    df = df.copy()
    # Cabeçalhos numéricos (ex.: anos) são tratados como texto no resto da aplicação.
    df.columns = [str(c) for c in df.columns]
    if ask_integers:
        # O pandas converte inteiros com células vazias em decimais; só o utilizador sabe se era intencional.
        for col in df.columns:
            if _integers_with_gaps(df[col]) and confirm(f"A coluna '{col}' só tem valores inteiros (e células vazias). Tratar como inteiro?"):
                df[col] = df[col].astype("Int64")
    if detect_dates:
        for col in df.columns:
            if infer_type(df[col]) is str:
                dates = _as_dates(df[col])
                if dates is not None:
                    df[col] = dates
    return df, {col: infer_type(df[col]) for col in df.columns}


def adopt_dataframe(df, detect_dates=False, ask_integers=False):
    """Torna `df` a folha activa (importação de Excel/CSV).

    O estado global só é alterado no fim, por isso uma falha a meio não deixa
    a folha actual meio substituída.
    """
    global table, col_types, meta
    prepared, new_col_types = _prepare_dataframe(df, detect_dates, ask_integers)
    table, col_types, meta = prepared, new_col_types, new_meta()
    apply_types()
    commit()


def adopt_workbook(frames, detect_dates=False, ask_integers=False):
    """Substitui o livro inteiro: cada item de `frames` ({nome: DataFrame}) passa a ser uma folha; a primeira fica activa."""
    prepared = []
    for name, df in frames.items():
        data, types = _prepare_dataframe(df, detect_dates, ask_integers)
        prepared.append((name, data, types, new_meta()))
    _install_workbook(prepared, prepared[0][0], dict(settings))
    commit()


DATE_INPUT_FORMATS = ("%d/%m/%Y", "%Y-%m-%d")  # DD/MM/AAAA (a da aplicação) e ISO (habitual em JSON)


def _as_dates(series):
    """Converte uma coluna de texto em datas se TODOS os valores forem DD/MM/AAAA (ou AAAA-MM-DD); senão None."""
    valid_values = series.dropna()
    if valid_values.empty:
        return None
    for date_format in DATE_INPUT_FORMATS:
        try:
            pd.to_datetime(valid_values, format=date_format)
        except (ValueError, TypeError):
            continue
        return pd.to_datetime(series, format=date_format, errors="coerce")
    return None


def _separator_of(sample):
    """Detecta o separador entre os mais comuns; ',' se não for possível (ex.: uma só coluna)."""
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def _csv_separator(name, file_encoding):
    with open(name, encoding=file_encoding, newline="") as f:
        return _separator_of(f.read(8192))


def _infer_columns(raw):
    """Inferência de tipos a partir do texto cru de cada coluna (ver _numeric_or_text)."""
    for col in raw.columns:
        raw[col] = _numeric_or_text(raw[col])
    return raw


def parse_csv_text(text):
    """Texto CSV/TSV (por exemplo copiado do Excel) -> DataFrame com os tipos inferidos. Levanta ValueError se não houver dados."""
    if not text.strip():
        raise ValueError("não há dados")
    try:
        raw = pd.read_csv(io.StringIO(text), sep=_separator_of(text[:8192]), dtype=str)
    except pd.errors.EmptyDataError:
        raise ValueError("não há dados") from None
    return _infer_columns(raw)


def _numeric_or_text(series):
    """Inferência a partir do texto cru: inteiros (com falhas) -> Int64, decimais -> float64, resto -> texto.

    Códigos com zeros à esquerda (ex.: 007) ficam como texto para não perder os zeros.
    """
    series = series.str.strip()
    valid_values = series.dropna()
    if valid_values.empty or valid_values.str.contains(r"^[+-]?0\d").any():
        return series
    if valid_values.str.fullmatch(r"[+-]?\d+(,\d+)?").all() and valid_values.str.contains(",").any():  # vírgula decimal (Excel PT)
        series = series.str.replace(",", ".", regex=False)
        valid_values = series.dropna()
    if valid_values.str.fullmatch(r"[+-]?\d+").all():
        return pd.to_numeric(series).astype("Int64")
    converted = pd.to_numeric(series, errors="coerce")
    if converted[series.notna()].notna().all():
        return converted
    return series


def _read_csv(name):
    """Lê um CSV (UTF-8 ou Latin-1, separador detectado) e infere os tipos das colunas."""
    for file_encoding in ("utf-8-sig", "latin-1"):
        try:
            sep = _csv_separator(name, file_encoding)
            raw = pd.read_csv(name, sep=sep, encoding=file_encoding, dtype=str)
        except UnicodeDecodeError:
            continue
        except pd.errors.EmptyDataError:
            raise ValueError("o ficheiro está vazio") from None
        return _infer_columns(raw)
    raise ValueError("não foi possível descodificar o ficheiro (esperado UTF-8 ou Latin-1)")


def import_csv():
    header("IMPORTAR CSV")
    print()
    if not confirm_replace_sheet():
        return
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return

    if not name.endswith(EXT_CSV):
        name += EXT_CSV

    try:
        adopt_dataframe(_read_csv(name), detect_dates=True)
        print(f"\n{UI_INDENT_1}'{name}' importado: {len(table)} linhas, {len(table.columns)} colunas.")
    except FileNotFoundError:
        print(f"\n{UI_INDENT_1}Ficheiro '{name}' não encontrado.")
    except Exception as e:
        log.exception("Falha em import_csv")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def write_csv(path, style="intl"):
    """Grava a folha activa em CSV: "intl" (',' e '.') ou "pt" (';' e ',', como o Excel português)."""
    sep, decimal = (";", ",") if style == "pt" else (",", ".")
    # utf-8-sig: o Excel reconhece os acentos; datas em DD/MM/AAAA, como na aplicação.
    table.to_csv(path, index=False, sep=sep, decimal=decimal, encoding="utf-8-sig", date_format="%d/%m/%Y")


def export_csv():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem dados para exportar.")
        pause()
        return

    header("EXPORTAR CSV")
    print()
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.endswith(EXT_CSV):
        name += EXT_CSV

    print(f"{UI_INDENT_1}Formato: [1] Internacional (',' e '.')  [2] Excel português (';' e ',')")
    style = "pt" if input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() == "2" else "intl"

    try:
        write_csv(name, style)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_csv")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def write_word(path):
    doc = Document()
    doc.add_heading(f"{APP_NAME} — Relatório", 0)
    doc.add_paragraph(f"Dimensões: {len(table)} linhas × {len(table.columns)} colunas")

    doc.add_heading("Colunas e tipos", level=1)
    for col in table.columns:
        doc.add_paragraph(f"• {col}: {TYPE_NAME.get(type_of(col), 'texto')}")

    if not table.empty:
        doc.add_heading("Dados", level=1)
        cols = list(table.columns)
        tbl = doc.add_table(rows=1 + len(table), cols=len(cols))
        tbl.style = "Table Grid"
        for j, col in enumerate(cols):
            tbl.rows[0].cells[j].text = str(col)
        for i, (_, row) in enumerate(table.iterrows()):
            for j, col in enumerate(cols):
                tbl.rows[i + 1].cells[j].text = format_value(row[col])

    num_cols = table.select_dtypes(include="number").columns.tolist()
    if num_cols:
        doc.add_heading("Estatísticas básicas (colunas numéricas)", level=1)
        for col in num_cols:
            series = table[col].dropna()
            doc.add_paragraph(
                f"{col}:  média={format_number(series.mean())}  "
                f"mín={format_number(series.min())}  "
                f"máx={format_number(series.max())}  "
                f"desvio={format_number(series.std())}  "
                f"mediana={format_number(series.median())}"
            )

    doc.save(path)


def export_word():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem dados para exportar.")
        pause()
        return

    header("EXPORTAR WORD")
    print()
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.endswith(EXT_WORD):
        name += EXT_WORD

    try:
        write_word(name)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_word")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def _latin1(text):
    """Fontes core do PDF só suportam Latin-1: o resto é substituído por '?'."""
    return str(text).encode("latin-1", "replace").decode("latin-1")


def _pdf_text(pdf, text):
    """Texto pronto a escrever no PDF: intacto com fonte Unicode, senão reduzido a Latin-1."""
    return str(text) if getattr(pdf, "tf_unicode", False) else _latin1(text)


def _configure_pdf_font(pdf):
    """Devolve o nome da família a usar. Com fpdf2 regista a DejaVu Sans (vem com o matplotlib) para suportar Unicode.

    O fpdf antigo (1.7) escreve ficheiros de cache na pasta da fonte, por isso aí mantém-se Helvetica (Latin-1).
    """
    pdf.tf_unicode = False
    try:
        if int(str(getattr(fpdf_lib, "FPDF_VERSION", "1")).split(".")[0]) < 2:
            return "Helvetica"
        folder = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
        pdf.add_font("DejaVu", "", os.path.join(folder, "DejaVuSans.ttf"))
        pdf.add_font("DejaVu", "B", os.path.join(folder, "DejaVuSans-Bold.ttf"))
    except Exception:
        log.warning("Fonte Unicode indisponível; o PDF usa Helvetica (Latin-1)", exc_info=True)
        return "Helvetica"
    pdf.tf_unicode = True
    return "DejaVu"


def _pdf_line(pdf, height, text, align=""):
    """Escreve uma linha e avança. Funciona em fpdf e fpdf2 (evita o `ln=True` deprecado)."""
    pdf.cell(0, height, _pdf_text(pdf, text), align=align)
    pdf.ln(height)


def write_pdf(path):
    pdf = FPDF()
    font_family = _configure_pdf_font(pdf)
    pdf.add_page()

    # Título
    pdf.set_font(font_family, "B", 16)
    pdf.set_text_color(16, 41, 230)
    _pdf_line(pdf, 10, f"{APP_NAME} - Relatório", align="C")

    # Dimensões
    pdf.set_font(font_family, size=11)
    pdf.set_text_color(0, 0, 0)
    _pdf_line(pdf, 8, f"Dimensões: {len(table)} linhas × {len(table.columns)} colunas")

    # Colunas e tipos
    pdf.ln(3)
    pdf.set_font(font_family, "B", 12)
    pdf.set_text_color(14, 122, 230)
    _pdf_line(pdf, 8, "Colunas e tipos")
    pdf.set_font(font_family, size=10)
    pdf.set_text_color(0, 0, 0)
    for col in table.columns:
        type_label = TYPE_NAME.get(type_of(col), "texto")
        _pdf_line(pdf, 7, f"  {col}: {type_label}")

    # Dados
    if not table.empty:
        pdf.ln(3)
        pdf.set_font(font_family, "B", 12)
        pdf.set_text_color(14, 122, 230)
        _pdf_line(pdf, 8, "Dados")
        pdf.set_text_color(0, 0, 0)

        cols = list(table.columns)
        col_w = min(35, (pdf.w - 20) / len(cols))
        pdf.set_font(font_family, "B", 8)
        for col in cols:
            pdf.cell(col_w, 7, _pdf_text(pdf, str(col)[:14]), border=1)
        pdf.ln()
        pdf.set_font(font_family, size=8)
        for _, row in table.iterrows():
            for col in cols:
                pdf.cell(col_w, 6, _pdf_text(pdf, str(format_value(row[col]))[:14]), border=1)
            pdf.ln()

    # Estatísticas
    num_cols = table.select_dtypes(include="number").columns.tolist()
    if num_cols:
        pdf.ln(3)
        pdf.set_font(font_family, "B", 12)
        pdf.set_text_color(14, 122, 230)
        _pdf_line(pdf, 8, "Estatísticas básicas (colunas numéricas)")
        pdf.set_font(font_family, size=10)
        pdf.set_text_color(0, 0, 0)
        for col in num_cols:
            series = table[col].dropna()
            line = (
                f"{col}: média={format_number(series.mean())}  "
                f"mín={format_number(series.min())}  "
                f"máx={format_number(series.max())}  "
                f"desvio={format_number(series.std())}  "
                f"mediana={format_number(series.median())}"
            )
            _pdf_line(pdf, 7, line)

    pdf.output(path)


def export_pdf():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem dados para exportar.")
        pause()
        return

    header("EXPORTAR PDF")
    print()
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.endswith(EXT_PDF):
        name += EXT_PDF

    try:
        write_pdf(name)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_pdf")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def write_excel(path, all_sheets=False):
    """Grava a folha activa (ou todas, cada uma num separador) em Excel."""
    if not all_sheets:
        table.to_excel(path, index=False)
        return
    with pd.ExcelWriter(path) as writer:
        for name in sheet_names():
            get_sheet(name)["table"].to_excel(writer, sheet_name=name, index=False)


def export_excel():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem dados para exportar.")
        pause()
        return

    header("EXPORTAR EXCEL")
    print()
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return

    if not name.endswith(EXT_EXCEL):
        name += EXT_EXCEL

    try:
        all_sheets = len(sheet_names()) > 1 and confirm("Exportar todas as folhas do livro (cada uma num separador)?")
        write_excel(name, all_sheets)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_excel")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


# JSON, texto colado, Markdown, HTML e ficheiros de projecto
PASTE_MAX_LINES = 100_000
PASTE_END_WORD = "fim"


def _json_cell(value):
    """Valor JSON (escalar) -> valor de célula; estruturas aninhadas viram texto JSON."""
    if isinstance(value, bool):
        return str(value).lower()
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return value


def _column_from_json_values(values):
    """Lista de valores JSON de uma coluna -> Series: inteiros (com nulos) em Int64, números em float, o resto em texto."""
    present = [v for v in values if v is not None]
    if present and all(isinstance(v, int) and not isinstance(v, bool) for v in present):
        return pd.Series(pd.array(values, dtype="Int64"))
    if present and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in present):
        return pd.Series(values, dtype="float64")
    return pd.Series([None if v is None else (v if isinstance(v, str) else _plain_text(v)) for v in values], dtype=object)


def json_to_dataframe(data):
    """Estrutura JSON -> DataFrame. Aceita uma lista de objectos (um por linha) ou um objecto de listas (uma por coluna).

    Levanta ValueError se a estrutura não for suportada ou estiver vazia.
    """
    if isinstance(data, dict) and data and all(isinstance(v, list) for v in data.values()):
        lengths = {len(v) for v in data.values()}
        if len(lengths) != 1:
            raise ValueError("as listas de cada coluna têm comprimentos diferentes")
        columns = {str(k): [_json_cell(v) for v in values] for k, values in data.items()}
    elif isinstance(data, list) and data and all(isinstance(r, dict) for r in data):
        names = list(dict.fromkeys(str(k) for r in data for k in r))
        columns = {name: [_json_cell(r.get(name)) for r in data] for name in names}
    else:
        raise ValueError("estrutura JSON não suportada (esperava uma lista de objectos ou um objecto de listas) ou vazia")
    if not columns or not any(columns.values()):
        raise ValueError("o JSON não tem linhas")
    return pd.DataFrame({name: _column_from_json_values(values) for name, values in columns.items()})


def read_json_file(path):
    with open(path, encoding="utf-8-sig") as f:
        return json_to_dataframe(json.load(f))


def _export_value(value):
    """Valor de célula para JSON de exportação: em falta -> null, datas DD/MM/AAAA, números como tipos nativos."""
    if pd.isna(value):
        return None
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return value.item() if hasattr(value, "item") else value


def write_json(path):
    """Grava a folha activa como lista de objectos JSON (um por linha)."""
    records = [{str(c): _export_value(v) for c, v in zip(table.columns, row, strict=True)} for row in table.itertuples(index=False, name=None)]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def _markdown_cell(text):
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def write_markdown(path):
    """Grava a folha activa como tabela Markdown (com os formatos de apresentação)."""
    columns = list(table.columns)
    numeric = [pd.api.types.is_numeric_dtype(table[c]) for c in columns]
    lines = [f"# {_markdown_cell(active_name)}", "", f"{len(table)} linhas × {len(columns)} colunas", ""]
    lines.append("| " + " | ".join(_markdown_cell(str(c)) for c in columns) + " |")
    lines.append("| " + " | ".join("---:" if n else "---" for n in numeric) + " |")
    for row in table.itertuples(index=False, name=None):
        lines.append("| " + " | ".join(_markdown_cell(format_cell(c, v)) for c, v in zip(columns, row, strict=True)) + " |")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")


_HTML_STYLE = (
    "body{font-family:system-ui,sans-serif;margin:2rem}table{border-collapse:collapse}"
    "th,td{border:1px solid #bbb;padding:.3rem .6rem}th{background:#f0f0f0;text-align:left}td.num{text-align:right}"
)


def write_html(path):
    """Grava a folha activa como página HTML autónoma (todo o texto é escapado; sem scripts nem recursos externos)."""
    columns = list(table.columns)
    numeric = [pd.api.types.is_numeric_dtype(table[c]) for c in columns]
    title = html.escape(active_name)
    rows = ["<tr>" + "".join(f"<th>{html.escape(str(c))}</th>" for c in columns) + "</tr>"]
    for row in table.itertuples(index=False, name=None):
        cells = "".join(
            f'<td class="num">{html.escape(format_cell(c, v))}</td>' if n else f"<td>{html.escape(format_cell(c, v))}</td>"
            for c, v, n in zip(columns, row, numeric, strict=True)
        )
        rows.append(f"<tr>{cells}</tr>")
    document = (
        '<!doctype html>\n<html lang="pt">\n<head>\n<meta charset="utf-8">\n'
        f"<title>{title}</title>\n<style>{_HTML_STYLE}</style>\n</head>\n<body>\n"
        f"<h1>{title}</h1>\n<p>{len(table)} linhas × {len(columns)} colunas</p>\n<table>\n" + "\n".join(rows) + "\n</table>\n</body>\n</html>\n"
    )
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(document)


def _export_dialog(title, extension, writer):
    """Diálogo comum de exportação: pede o nome, acrescenta a extensão e chama `writer(caminho)`."""
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem dados para exportar.")
        pause()
        return
    header(title)
    print()
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.lower().endswith(extension):
        name += extension
    try:
        writer(name)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em %s", title)
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def export_json():
    _export_dialog("EXPORTAR JSON", ".json", write_json)


def export_markdown():
    _export_dialog("EXPORTAR MARKDOWN", ".md", write_markdown)


def export_html():
    _export_dialog("EXPORTAR HTML", ".html", write_html)


def import_json():
    header("IMPORTAR JSON")
    print()
    if not confirm_replace_sheet():
        return
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.lower().endswith(".json"):
        name += ".json"
    try:
        adopt_dataframe(read_json_file(name), detect_dates=True)
        print(f"\n{UI_INDENT_1}'{name}' importado: {len(table)} linhas, {len(table.columns)} colunas.")
    except FileNotFoundError:
        print(f"\n{UI_INDENT_1}Ficheiro '{name}' não encontrado.")
    except (ValueError, json.JSONDecodeError) as e:
        print(f"\n{UI_INDENT_1}JSON inválido: {e}")
    except Exception as e:
        log.exception("Falha em import_json")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def import_pasted():
    header("COLAR DADOS (CSV / EXCEL)")
    print()
    if not confirm_replace_sheet():
        return
    print(f"{UI_INDENT_1}Cole os dados (a primeira linha são os nomes das colunas; vale copiar do Excel).")
    print(f"{UI_INDENT_1}Termine com uma linha vazia ou escrevendo '{PASTE_END_WORD}'. 'sair' cancela.\n")
    lines = []
    while len(lines) < PASTE_MAX_LINES:
        line = input()
        if line.strip().lower() == "sair":
            return
        if not line.strip() or line.strip().lower() == PASTE_END_WORD:
            break
        lines.append(line)
    try:
        adopt_dataframe(parse_csv_text("\n".join(lines)), detect_dates=True)
        print(f"\n{UI_INDENT_1}Colados {len(table)} linhas e {len(table.columns)} colunas.")
    except ValueError as e:
        print(f"\n{UI_INDENT_1}Não foi possível importar: {e}.")
    except Exception as e:
        log.exception("Falha em import_pasted")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def save_project():
    _export_dialog("GUARDAR PROJECTO", EXT_PROJECT, _write_state_file)


def open_project():
    header("ABRIR PROJECTO")
    print()
    if not confirm_replace_sheet():
        return
    name = input(f"{UI_INDENT_1}Nome do ficheiro (sem extensão): ").strip()
    if not name:
        return
    if not name.lower().endswith(EXT_PROJECT):
        name += EXT_PROJECT
    try:
        sheet_list, active, _file_settings = _read_state_file(name)
        _install_workbook(sheet_list, active, dict(settings))  # as definições são da instalação, não do projecto
        commit()
        print(f"\n{UI_INDENT_1}Projecto '{name}' aberto: {len(sheet_names())} folha(s).")
    except FileNotFoundError:
        print(f"\n{UI_INDENT_1}Ficheiro '{name}' não encontrado.")
    except (ValueError, json.JSONDecodeError) as e:
        print(f"\n{UI_INDENT_1}Projecto inválido: {e}")
    except Exception as e:
        log.exception("Falha em open_project")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


def create_dummy_sheet():
    global table, col_types, meta

    header("CRIAR FOLHA DUMMY")
    print()
    if not table.columns.empty:
        if not confirm("Folha actual será substituída. Continuar?"):
            print(f"\n{UI_INDENT_1}  Operação cancelada.")
            pause()
            return

    table = pd.DataFrame({
        "Nome": ["João", "Maria", "Pedro", "Ana", "Carlos", "Sofia", "João", "Rita"],
        "Idade": [25, 30, 35, None, 45, 28, 25, 33],
        "Salario": [1500.50, 2000.00, 1800.75, 2200.00, 3500.00, 1700.25, 1500.50, 1950.00],
        "Cidade": ["Lisboa", "Porto", "Faro", "Coimbra", "Lisboa", "Porto", "Lisboa", "Braga"],
        "Admissao": pd.to_datetime([
            "15/01/2020", "01/06/2019", "10/03/2021", "20/11/2018",
            "30/05/2017", "14/02/2022", "15/01/2020", "25/09/2019"
        ], format="%d/%m/%Y"),
    })
    col_types = {"Nome": str, "Idade": int, "Salario": float, "Cidade": str, "Admissao": DATE_TYPE}
    meta = new_meta()
    apply_types()
    commit()

    print(f"\n{UI_INDENT_1}Folha dummy criada: 8 linhas | 5 colunas")
    print(f"{UI_INDENT_1}Inclui:")
    print(f"{UI_INDENT_2}• Tipos variados: texto, inteiro, decimal, data")
    print(f"{UI_INDENT_2}• 1 valor em falta na coluna 'Idade'")
    print(f"{UI_INDENT_2}• 1 linha duplicada (João, 25, Lisboa)")
    print(f"{UI_INDENT_2}• Cidades repetidas (ideal para 'groupby')")
    pause()


def clear_sheet():
    global table, col_types, meta
    if table.columns.empty:
        print(f"{UI_INDENT_1}Folha já está vazia.")
        pause()
        return

    header("LIMPAR FOLHA")
    print()
    if not confirm("Confirma apagar toda a folha actual (perdem-se dados)?"):
        print(f"\n{UI_INDENT_1}  Operação cancelada.")
        pause()
        return

    table, col_types, meta = pd.DataFrame(), {}, new_meta()
    if len(sheet_names()) > 1:
        commit()  # as outras folhas do livro têm de ser preservadas
    else:
        try:
            if os.path.exists(STATE_FILE):
                os.remove(STATE_FILE)
        except OSError:
            log.warning("Não foi possível remover o ficheiro de estado", exc_info=True)
        record_checkpoint()
    print(f"\n{UI_INDENT_1}Folha limpa.")
    pause()


# Operações sobre linhas: adicionar, editar célula, remover
def add_row():
    global table
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return

    header("ADICIONAR LINHA")
    print(f"{UI_INDENT_1}('sair' em qualquer campo para terminar a introdução de linhas)\n")

    rows_added = 0

    while True:
        new_row = {}
        cancelled = False

        for col in table.columns:
            if col in meta["formulas"]:
                continue  # coluna calculada: é preenchida pela fórmula
            col_type = type_of(col)
            type_label = TYPE_NAME.get(col_type, "texto")
            while True:
                val = input(f"{UI_INDENT_1}{col} ({type_label}): ").strip()
                if val.lower() == "sair":
                    cancelled = True
                    break
                try:
                    converted = convert_value(val, col_type)
                except (ValueError, OverflowError):
                    print(f"{UI_INDENT_1}Valor inválido para {type_label}.")
                    continue
                problem = rule_violation(col, converted)
                if problem:
                    print(f"{UI_INDENT_1}Regra da coluna '{col}': {problem}.")
                    continue
                new_row[col] = converted
                break
            if cancelled:
                break

        if cancelled:
            break

        new_df = pd.DataFrame([new_row])
        table = pd.concat([table, new_df], ignore_index=True)
        apply_types()
        commit()
        rows_added += 1
        print(f"\n{UI_INDENT_1}Linha adicionada. (total introduzido: {rows_added})\n")

    if rows_added > 0:
        print(f"\n{UI_INDENT_1}{rows_added} linha(s) adicionada(s) com sucesso.")
    else:
        print(f"\n{UI_INDENT_1}Nenhuma linha adicionada.")
    pause()


def edit_cell():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados para editar.")
        pause()
        return

    header("EDITAR CÉLULA")
    print(f"{UI_INDENT_1}('sair' em qualquer campo para cancelar)\n")

    index_text = input(f"{UI_INDENT_1}Índice da linha: ").strip()
    if index_text.lower() == "sair":
        print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return
    try:
        idx = int(index_text)
    except ValueError:
        print(f"{UI_INDENT_1}Índice deve ser inteiro.")
        pause()
        return

    if idx not in table.index:
        print(f"{UI_INDENT_1}Índice inválido.")
        pause()
        return

    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    col_text = input(f"{UI_INDENT_1}Coluna: ").strip()
    if col_text.lower() == "sair":
        print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return
    col = resolve_column(col_text)
    if col is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return

    if col in meta["formulas"]:
        print(f"{UI_INDENT_1}'{col}' é uma coluna calculada ({meta['formulas'][col]}). Converta-a em valores para a poder editar.")
        pause()
        return

    col_type = type_of(col)
    current = table.at[idx, col]
    print(f"{UI_INDENT_1}Valor actual: {format_value(current)}")

    while True:
        loaded = input(f"{UI_INDENT_1}Novo valor: ").strip()
        if loaded.lower() == "sair":
            print(f"\n{UI_INDENT_1}Operação cancelada.")
            pause()
            return
        try:
            converted = convert_value(loaded, col_type)
        except (ValueError, OverflowError):
            print(f"{UI_INDENT_1}Valor inválido.")
            continue
        problem = rule_violation(col, converted)
        if problem:
            print(f"{UI_INDENT_1}Regra da coluna '{col}': {problem}.")
            continue
        table.at[idx, col] = converted
        break

    commit()
    print(f"\n{UI_INDENT_1}Célula actualizada.")
    pause()


def remove_row():
    global table
    if table.empty:
        print(f"{UI_INDENT_1}Sem linhas para remover.")
        pause()
        return

    header("REMOVER LINHA")
    print()

    try:
        idx = int(input(f"{UI_INDENT_1}Índice da linha a remover: ").strip())
    except ValueError:
        print(f"{UI_INDENT_1}Índice deve ser inteiro.")
        pause()
        return

    if idx not in table.index:
        print(f"{UI_INDENT_1}Índice inválido.")
        pause()
        return

    print()
    draw_data_table(table.loc[[idx]])
    print()
    if not confirm(f"Confirma remover linha {idx}?"):
        print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return

    table = table.drop(index=idx).reset_index(drop=True)
    commit()
    print(f"\n{UI_INDENT_1}Linha {idx} removida.")
    pause()


# Fórmulas: avaliador próprio sobre a árvore sintáctica (ast). Nunca usa eval/exec: só os nós e funções listados abaixo são aceites,
# por isso uma fórmula vinda de um ficheiro de estado adulterado não consegue executar código.
class FormulaError(ValueError):
    """Fórmula inválida (sintaxe, nome desconhecido ou operação não permitida)."""


_FORMULA_MAX_LENGTH = 500
_FORMULA_MAX_NODES = 200
_FORMULA_BINARY = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv, ast.FloorDiv: operator.floordiv, ast.Mod: operator.mod, ast.Pow: operator.pow}
_FORMULA_COMPARE = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt, ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge}
_FORMULA_NODES = (
    ast.Expression, ast.BinOp, ast.UnaryOp, ast.BoolOp, ast.Compare, ast.Call, ast.Name, ast.Constant, ast.Load,
    ast.And, ast.Or, ast.Not, ast.USub, ast.UAdd, *_FORMULA_BINARY, *_FORMULA_COMPARE,
)
_FORMULA_AGGREGATES = {"soma": "sum", "sum": "sum", "media": "mean", "average": "mean", "avg": "mean", "contar": "count", "count": "count"}  # nome -> agregação
_FORMULA_ARITY = {  # nome -> (mínimo, máximo) de argumentos
    "abs": (1, 1), "round": (1, 2), "sqrt": (1, 1), "log": (1, 1), "exp": (1, 1), "floor": (1, 1), "ceil": (1, 1),
    "min": (1, 2), "max": (1, 2), "se": (3, 3), "year": (1, 1), "month": (1, 1), "day": (1, 1), "days": (2, 2),
    **dict.fromkeys(_FORMULA_AGGREGATES, (1, 1)),
}
_FORMULA_FLOAT_FUNCTIONS = {"sqrt", "log", "exp", "media", "average", "avg"}  # produzem sempre decimais

# Referências a células: coluna por letra (A = 1.ª coluna) ou por nome entre parênteses rectos, mais o nº da linha como no ecrã (#, a começar em 0).
# `B2` é uma célula, `B2:B9` um intervalo (só dentro de soma/media/contar/min/max). São sempre fixas; o `$` é aceite e ignorado.
_CELL_PATTERN = r"\$?[A-Za-z]{1,3}\$?\d+"
_BRACKET_CELL_PATTERN = r"\[[^\[\]]+\]\$?\d+"
_REF_PATTERN = rf"(?:{_BRACKET_CELL_PATTERN}|{_CELL_PATTERN})"
_FORMULA_TOKEN = re.compile(
    r"""(?P<text>'[^']*'|"[^"]*")"""  # texto entre aspas: fica como está
    rf"""|(?P<ref>(?<![\w.]){_REF_PATTERN}(?:\s*:\s*{_REF_PATTERN})?(?![\w(\[.]))"""
    r"""|(?P<column>\[[^\[\]]+\])"""
)


def column_letter(index):
    """Letra da coluna na posição `index` (0 -> A, 25 -> Z, 26 -> AA)."""
    letters = ""
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _cell_endpoint(token, columns):
    """('B2' ou '[Salario]2') -> (nome da coluna, nº da linha). Levanta FormulaError."""
    token = token.strip()
    if token.startswith("["):
        name, row = re.fullmatch(r"\[([^\[\]]+)\]\$?(\d+)", token).groups()
        column = resolve_column(name.strip(), source=columns)
        if column is None:
            raise FormulaError(f"coluna inexistente: [{name.strip()}]")
        return column, int(row)
    letters, row = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", token).groups()
    index = 0
    for char in letters.upper():
        index = index * 26 + ord(char) - 64
    if index > len(columns):
        raise FormulaError(f"a coluna {letters.upper()} não existe (a folha tem {len(columns)} coluna(s))")
    return columns[index - 1], int(row)


def _is_column_name(token, columns):
    """Uma 'célula' como `Q1` que é afinal o nome de uma coluna existente (o nome ganha à referência)."""
    return ":" not in token and not token.startswith("[") and resolve_column(token, source=columns) is not None


def canonical_formula(expression, columns):
    """Troca as referências por letra (`B2`) por referências por nome (`[Salario]2`), que não mudam se as colunas mudarem de posição."""
    def replace(match):
        token = match["ref"]
        if token is None or _is_column_name(token, columns):
            return match.group(0)
        parts = []
        for part in re.findall(_REF_PATTERN, token):
            column, row = _cell_endpoint(part, columns)
            if "[" in column or "]" in column:
                raise FormulaError(f"a coluna '{column}' tem parênteses rectos no nome e não pode ser referida")
            parts.append(f"[{column}]{row}")
        return ":".join(parts)

    return _FORMULA_TOKEN.sub(replace, expression.strip())


def _parse_formula(expression, columns):
    """Valida `expression` e devolve (árvore, {nome no texto: coluna}, colunas referidas, {nome no texto: extremos da célula/intervalo}). Levanta FormulaError."""
    text = expression.strip()
    if not text:
        raise FormulaError("a fórmula está vazia")
    if len(text) > _FORMULA_MAX_LENGTH:
        raise FormulaError(f"a fórmula tem mais de {_FORMULA_MAX_LENGTH} caracteres")

    names = {}
    refs = {}

    def replace(match):  # [Nome com espaços] e referências a células -> identificadores seguros
        if match["text"] is not None:
            return match["text"]
        if match["column"] is not None:
            name = match["column"][1:-1].strip()
            column = resolve_column(name, source=columns)
            if column is None:
                raise FormulaError(f"coluna inexistente: [{name}]")
            key = f"__col{len(names)}__"
            names[key] = column
            return key
        token = match["ref"]
        if _is_column_name(token, columns):
            return token
        key = f"__ref{len(refs)}__"
        refs[key] = tuple(_cell_endpoint(part, columns) for part in re.findall(_REF_PATTERN, token))
        return key

    text = _FORMULA_TOKEN.sub(replace, text)
    try:
        tree = ast.parse(text, mode="eval")
    except SyntaxError:
        raise FormulaError("sintaxe inválida") from None
    nodes = list(ast.walk(tree))
    if len(nodes) > _FORMULA_MAX_NODES:
        raise FormulaError("fórmula demasiado complexa")

    call_names = {id(n.func) for n in nodes if isinstance(n, ast.Call)}
    aggregate_args = {  # onde um intervalo é permitido: soma/media/contar e min/max com um só argumento
        id(arg) for n in nodes if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) for arg in n.args
        if n.func.id.lower() in _FORMULA_AGGREGATES or (n.func.id.lower() in ("min", "max") and len(n.args) == 1)
    }
    for node in nodes:
        if not isinstance(node, _FORMULA_NODES):
            raise FormulaError(f"expressão não permitida ({type(node).__name__})")
        if isinstance(node, ast.Constant) and (isinstance(node.value, bool) or not isinstance(node.value, (int, float, str))):
            raise FormulaError("só são aceites números e texto")
        if isinstance(node, ast.Compare) and len(node.ops) != 1:
            raise FormulaError("comparações encadeadas não são suportadas")
        if isinstance(node, ast.Call):
            func = node.func
            if not isinstance(func, ast.Name) or func.id.lower() not in _FORMULA_ARITY or node.keywords:
                raise FormulaError("função desconhecida ou chamada inválida")
            low, high = _FORMULA_ARITY[func.id.lower()]
            if not low <= len(node.args) <= high:
                raise FormulaError(f"'{func.id}' espera entre {low} e {high} argumento(s)")
        elif isinstance(node, ast.Name) and node.id in refs:
            if len(refs[node.id]) > 1 and id(node) not in aggregate_args:
                raise FormulaError("um intervalo só pode ser usado dentro de soma, media, contar, min ou max")
        elif isinstance(node, ast.Name) and id(node) not in call_names and node.id not in names:
            column = resolve_column(node.id, source=columns)
            if column is None:
                raise FormulaError(f"coluna ou função desconhecida: '{node.id}'")
            names[node.id] = column
    referenced = {names[n.id] for n in nodes if isinstance(n, ast.Name) and id(n) not in call_names and n.id in names}
    referenced |= {column for ends in refs.values() for column, _ in ends}
    for ends in refs.values():
        if len(ends) > 1:  # um intervalo abrange todas as colunas entre os extremos
            positions = sorted(list(columns).index(column) for column, _ in ends)
            referenced |= set(list(columns)[positions[0]:positions[-1] + 1])
    return tree, names, referenced, refs


def _formula_operand(data, col, types):
    """Coluna pronta a calcular: números como float64 (vazios = NaN), datas como datetime, texto como object."""
    col_type = types.get(col, str)
    if col_type in (int, float):
        return pd.to_numeric(data[col], errors="coerce").astype("float64")
    if col_type == DATE_TYPE:
        return pd.to_datetime(data[col], errors="coerce")
    return data[col].astype(object)


def _formula_reference(ends, data, types):
    """Valor de uma célula (escalar) ou de um intervalo (Series só com os valores) da folha. Levanta FormulaError se a linha não existir."""
    columns = list(data.columns)
    low, high = sorted(row for _, row in (ends[0], ends[-1]))
    if high >= len(data):
        raise FormulaError(f"a linha {high} não existe (linhas 0 a {len(data) - 1})" if len(data) else "a folha não tem linhas")
    if len(ends) == 1:
        return _formula_operand(data, ends[0][0], types).iloc[low]
    first, last = sorted(columns.index(column) for column, _ in ends)
    parts = [_formula_operand(data, columns[p], types).iloc[low:high + 1].reset_index(drop=True) for p in range(first, last + 1)]
    if len(parts) == 1:
        return parts[0]
    return pd.Series(np.concatenate([part.to_numpy(dtype=object) for part in parts]), dtype=object)


def _formula_aggregate(function, value):
    """soma/media/contar/min/max de uma coluna, de um intervalo ou de um valor. Devolve um escalar."""
    values = value if isinstance(value, pd.Series) else pd.Series([value])
    if function == "count":
        return np.float64(values.notna().sum())
    if pd.api.types.is_datetime64_any_dtype(values):
        if function in ("min", "max"):
            return getattr(values, function)()
        raise FormulaError("não se pode somar nem calcular a média de datas")
    numbers = pd.to_numeric(values, errors="coerce").dropna()
    return np.float64(getattr(numbers, function)()) if len(numbers) or function == "sum" else np.float64("nan")


def _formula_function(name, args, index):
    """Aplica a função `name` (já validada) a `args`; devolve Series/escalar."""
    def as_series(value):
        return value if isinstance(value, pd.Series) else pd.Series(value, index=index)

    if name in _FORMULA_AGGREGATES:
        return _formula_aggregate(_FORMULA_AGGREGATES[name], args[0])
    if name in ("min", "max") and len(args) == 1:
        return _formula_aggregate(name, args[0])
    if name == "round":
        decimals = int(args[1]) if len(args) > 1 else 0
        return as_series(args[0]).round(decimals)
    if name == "se":
        condition = as_series(args[0])
        mask = condition.fillna(False).astype(bool) if condition.dtype != object else condition.eq(True)
        return pd.Series(np.where(mask, as_series(args[1]), as_series(args[2])), index=index)
    if name in ("year", "month", "day"):
        values = pd.to_datetime(as_series(args[0]), errors="coerce")
        return getattr(values.dt, name).astype("float64")
    if name == "days":
        return (pd.to_datetime(as_series(args[0]), errors="coerce") - pd.to_datetime(as_series(args[1]), errors="coerce")).dt.days.astype("float64")
    functions = {"abs": np.abs, "sqrt": np.sqrt, "log": np.log, "exp": np.exp, "floor": np.floor, "ceil": np.ceil, "min": np.minimum, "max": np.maximum}
    result = functions[name](*[as_series(a) if isinstance(a, pd.Series) else a for a in args])
    return as_series(result)


def evaluate_formula(expression, data, types):
    """Avalia `expression` sobre `data`. Devolve (Series, tipo da coluna: int/float/str/DATE_TYPE). Levanta FormulaError."""
    tree, names, referenced, refs = _parse_formula(expression, list(data.columns))
    operands = {key: _formula_operand(data, col, types) for key, col in names.items()}
    operands.update({key: _formula_reference(ends, data, types) for key, ends in refs.items()})

    def value_of(node):
        if isinstance(node, ast.Constant):
            return np.float64(node.value) if isinstance(node.value, (int, float)) else node.value  # float: evita potências gigantes em inteiros Python
        if isinstance(node, ast.Name):
            return operands[node.id]
        if isinstance(node, ast.BinOp):
            return _FORMULA_BINARY[type(node.op)](value_of(node.left), value_of(node.right))
        if isinstance(node, ast.UnaryOp):
            operand = value_of(node.operand)
            return {ast.USub: operator.neg, ast.UAdd: operator.pos, ast.Not: operator.invert}[type(node.op)](operand)
        if isinstance(node, ast.Compare):
            return _FORMULA_COMPARE[type(node.ops[0])](value_of(node.left), value_of(node.comparators[0]))
        if isinstance(node, ast.BoolOp):
            combine = operator.and_ if isinstance(node.op, ast.And) else operator.or_
            result = value_of(node.values[0])
            for other in node.values[1:]:
                result = combine(result, value_of(other))
            return result
        return _formula_function(node.func.id.lower(), [value_of(a) for a in node.args], data.index)  # ast.Call

    try:
        with np.errstate(all="ignore"):
            value = value_of(tree.body)
            series = value if isinstance(value, pd.Series) else pd.Series(value, index=data.index)
            series = series.replace([np.inf, -np.inf], np.nan) if pd.api.types.is_float_dtype(series) else series
    except FormulaError:
        raise
    except Exception as e:  # noqa: BLE001 - o pandas pode falhar de muitas formas; tudo vira FormulaError
        raise FormulaError(f"operação inválida ({e})") from None

    uses_division = any(isinstance(n, ast.BinOp) and isinstance(n.op, ast.Div) for n in ast.walk(tree))
    uses_float_function = any(isinstance(n, ast.Call) and n.func.id.lower() in _FORMULA_FLOAT_FUNCTIONS for n in ast.walk(tree))
    if pd.api.types.is_datetime64_any_dtype(series):
        return series, DATE_TYPE
    if pd.api.types.is_bool_dtype(series):
        return series.astype("int64"), int
    if pd.api.types.is_numeric_dtype(series):
        present = series.dropna()
        integral = bool((present % 1 == 0).all())
        all_int_inputs = bool(referenced) and all(types.get(c) in (int, DATE_TYPE) for c in referenced)  # year/month/day/days de datas dão inteiros
        return series.astype("float64"), (int if integral and all_int_inputs and not uses_division and not uses_float_function else float)
    return series.astype(object), str


def formula_references(expression):
    """Colunas da folha activa que `expression` referencia (vazio se for inválida)."""
    try:
        return _parse_formula(expression, list(table.columns))[2]  # (árvore, nomes, colunas referidas, células)
    except FormulaError:
        return set()


def formula_users(col):
    """Colunas calculadas cuja fórmula usa `col`."""
    return [c for c, expression in meta["formulas"].items() if c != col and col in formula_references(expression)]


def recalculate_formulas():
    """Recalcula as colunas calculadas da folha activa, pela ordem de criação. Uma fórmula que falhe deixa a coluna em branco e avisa."""
    for col, expression in list(meta["formulas"].items()):
        if col not in table.columns:
            continue
        try:
            series, col_type = evaluate_formula(expression, table, col_types)
        except FormulaError as e:
            log.warning("Fórmula de %r falhou: %s", col, e)
            print(f"{UI_INDENT_1}Aviso: a fórmula de '{col}' falhou ({e}); a coluna fica em branco.")
            series, col_type = pd.Series(np.nan, index=table.index, dtype="float64"), float
        table[col] = series
        col_types[col] = col_type
    if meta["formulas"]:
        apply_types()


# Operações sobre colunas: adicionar, remover, renomear
def forget_column_meta(col):
    """Remove os metadados (regras, formato, visibilidade, fórmula) de uma coluna eliminada."""
    for key in ("rules", "formats", "formulas"):
        meta[key].pop(col, None)
    meta["visible"] = [c for c in meta["visible"] if c != col]


def rename_column_meta(old, new):
    """Passa os metadados de uma coluna para o novo nome."""
    for key in ("rules", "formats", "formulas"):
        if old in meta[key]:
            meta[key][new] = meta[key].pop(old)
    meta["visible"] = [new if c == old else c for c in meta["visible"]]


def add_column():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return

    header("ADICIONAR COLUNA")
    print(f"{UI_INDENT_1}('sair' para cancelar)\n")

    name = input(f"{UI_INDENT_1}Nome da nova coluna: ").strip()
    if not name or name.lower() == "sair":
        return
    if resolve_column(name) is not None:
        print(f"{UI_INDENT_1}Coluna '{name}' já existe. Por favor escolha um nome diferente.")
        pause()
        return

    col_type = ask_type()
    col_types[name] = col_type
    type_label = TYPE_NAME.get(col_type, "texto")

    if table.empty:
        table[name] = pd.Series(dtype=object)
        apply_types()
        commit()
        print(f"\n{UI_INDENT_1}Coluna '{name}' adicionada.")
        pause()
        return

    values = []
    for idx in table.index:
        while True:
            val = input(f"{UI_INDENT_1}Valor para linha {idx} ({type_label}) ('sair' cancela, Enter = em branco): ").strip()
            if val.lower() == "sair":
                col_types.pop(name, None)
                print(f"\n{UI_INDENT_1}Operação cancelada.")
                pause()
                return
            if not val:
                values.append(None)  # em branco = valor em falta (apresentado como "--")
                break
            try:
                values.append(convert_value(val, col_type))
                break
            except (ValueError, OverflowError):
                print(f"{UI_INDENT_1}Valor inválido para {type_label}.")

    table[name] = values
    apply_types()
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' adicionada.")
    pause()


def remove_column():
    global table
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem colunas para remover.")
        pause()
        return

    header("REMOVER COLUNA")
    print()
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    name = resolve_column(input(f"{UI_INDENT_1}Coluna a remover: ").strip())

    if name is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return

    users = formula_users(name)
    if users:
        print(f"{UI_INDENT_1}'{name}' é usada pela fórmula de: {', '.join(users)}. Converta essas colunas em valores primeiro.")
        pause()
        return

    if not confirm(f"Confirma remover coluna '{name}'?"):
        print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return

    table = table.drop(columns=[name])
    col_types.pop(name, None)
    forget_column_meta(name)
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' removida.")
    pause()


def rename_column():
    global table
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem colunas para renomear.")
        pause()
        return

    header("RENOMEAR COLUNA")
    print()
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    name = resolve_column(input(f"{UI_INDENT_1}Coluna a renomear: ").strip())

    if name is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return

    users = formula_users(name)
    if users:
        print(f"{UI_INDENT_1}'{name}' é usada pela fórmula de: {', '.join(users)}. Converta essas colunas em valores primeiro.")
        pause()
        return

    while True:
        new_name = input(f"{UI_INDENT_1}Novo nome ('sair' para cancelar): ").strip()
        if new_name.lower() == "sair":
            print(f"\n{UI_INDENT_1}Operação cancelada.")
            pause()
            return
        if not new_name:
            print(f"{UI_INDENT_1}Nome não pode ser vazio.")
        elif resolve_column(new_name) is not None:
            print(f"{UI_INDENT_1}Já existe uma coluna com esse nome.")
        else:
            break

    table = table.rename(columns={name: new_name})
    if name in col_types:
        col_types[new_name] = col_types.pop(name)
    rename_column_meta(name, new_name)
    commit()
    print(f"\n{UI_INDENT_1}'{name}' renomeada para '{new_name}'.")
    pause()


# Análise básica: ordenar, filtrar (numérico/data/texto), ver início/fim, amostra
def sort_rows():
    global table
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados para ordenar.")
        pause()
        return

    header("ORDENAR")
    print(f"{UI_INDENT_1}Adicione colunas em sequência ('sair' para terminar)\n")
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}\n")

    sort_cols = []
    asc_order = []

    while True:
        value = input(f"{UI_INDENT_1}Coluna {len(sort_cols) + 1}: ").strip()
        if value.lower() == "sair":
            if not sort_cols:
                print(f"\n{UI_INDENT_1}Operação cancelada.")
                pause()
                return
            break
        if not value:
            if not sort_cols:
                continue
            break
        col = resolve_column(value)
        if col is None:
            print(f"{UI_INDENT_1}Coluna não existe.")
            continue
        if col in sort_cols:
            print(f"{UI_INDENT_1}Coluna já adicionada.")
            continue
        direction = input(f"{UI_INDENT_2}Direcção [A]scendente / [D]escendente: ").strip().upper()
        sort_cols.append(col)
        asc_order.append(direction != "D")

    table = table.sort_values(by=sort_cols, ascending=asc_order, ignore_index=True)
    commit()

    description = ", ".join(f"{c} ({'asc' if a else 'desc'})" for c, a in zip(sort_cols, asc_order, strict=True))
    print(f"\n{UI_INDENT_1}Ordenado por: {description}")
    pause()


def filter_numeric(col):
    print()
    print(f"{UI_INDENT_1}Operador: [1] ==  [2] >  [3] <  [4] >=  [5] <=  [6] intervalo [min,max]")
    op = input(f"{UI_INDENT_1}Opção: ").strip()
    try:
        if op == "6":
            mn = float(input(f"{UI_INDENT_1}Mínimo: ").strip())
            mx = float(input(f"{UI_INDENT_1}Máximo: ").strip())
            return table[table[col].between(mn, mx)]
        val = float(input(f"{UI_INDENT_1}Valor: ").strip())
        if op == "1":
            return table[table[col] == val]
        elif op == "2":
            return table[table[col] > val]
        elif op == "3":
            return table[table[col] < val]
        elif op == "4":
            return table[table[col] >= val]
        elif op == "5":
            return table[table[col] <= val]
        print(f"{UI_INDENT_1}Operador inválido.")
        pause()
        return None
    except ValueError:
        print(f"{UI_INDENT_1}Valor inválido.")
        pause()
        return None


def filter_date(col):
    print()
    print(f"{UI_INDENT_1}Operador: [1] Antes de  [2] Depois de  [3] Entre duas datas")
    op = input(f"{UI_INDENT_1}Opção: ").strip()
    try:
        if op == "3":
            d1 = pd.to_datetime(input(f"{UI_INDENT_1}De (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
            d2 = pd.to_datetime(input(f"{UI_INDENT_1}Até (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
            return table[(table[col] >= d1) & (table[col] <= d2)]
        d = pd.to_datetime(input(f"{UI_INDENT_1}Data (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
        if op == "1":
            return table[table[col] < d]
        elif op == "2":
            return table[table[col] > d]
        print(f"{UI_INDENT_1}Operador inválido.")
        pause()
        return None
    except (ValueError, TypeError, OverflowError):
        print(f"{UI_INDENT_1}Data inválida.")
        pause()
        return None


def filter_text(col):
    print()
    print(f"{UI_INDENT_1}Operador: [1] Contém  [2] Igual a  [3] Começa com  [4] Termina com")
    op = input(f"{UI_INDENT_1}Opção: ").strip()
    if op not in ("1", "2", "3", "4"):
        print(f"{UI_INDENT_1}Operador inválido.")
        pause()
        return None
    term = input(f"{UI_INDENT_1}Termo: ").strip()
    series = table[col].astype(str)
    term_lower = term.lower()
    if op == "2":
        return table[series.str.lower() == term_lower]
    elif op == "3":
        return table[series.str.lower().str.startswith(term_lower, na=False)]
    elif op == "4":
        return table[series.str.lower().str.endswith(term_lower, na=False)]
    return table[series.str.contains(term, case=False, na=False, regex=False)]


def filter_rows():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados para filtrar.")
        pause()
        return

    header("FILTRAR / PESQUISAR")
    print()
    col = ask_column("Coluna a filtrar")
    if col is None:
        return

    series = table[col]
    if pd.api.types.is_numeric_dtype(series):
        result = filter_numeric(col)
    elif pd.api.types.is_datetime64_any_dtype(series):
        result = filter_date(col)
    else:
        result = filter_text(col)

    if result is None:
        print(f"{UI_INDENT_1}Filtro inválido ou cancelado.")
        pause()
        return

    print()
    if result.empty:
        print(f"{UI_INDENT_1}Nenhum resultado encontrado.")
    else:
        print(f"{UI_INDENT_1}{len(result)} resultado(s) em '{col}':\n")
        draw_data_table(result)
    pause()


def view_head_tail():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    header("INÍCIO / FIM")
    print()
    print(f"{UI_INDENT_1}Mostra as primeiras ou últimas N linhas da tabela.")
    print()
    n_text = input(f"{UI_INDENT_1}N — nº de linhas a mostrar (Enter = 5): ").strip()
    try:
        n = int(n_text) if n_text else 5
        if n <= 0:
            raise ValueError
    except ValueError:
        n = 5

    view_mode = input(f"{UI_INDENT_1}[I] Início  [F] Fim: ").strip().upper()
    if view_mode == "F":
        print(f"\n{UI_INDENT_1}Últimas {n} linhas:\n")
        draw_data_table(table.tail(n))
    else:
        print(f"\n{UI_INDENT_1}Primeiras {n} linhas:\n")
        draw_data_table(table.head(n))
    pause()


def random_sample():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    header("AMOSTRA ALEATÓRIA")
    print()
    n_text = input(f"{UI_INDENT_1}Tamanho: nº de linhas (1-{len(table)}) ou percentagem, ex. 20% (Enter = 5): ").strip()
    try:
        if n_text.endswith("%"):
            pct = float(n_text[:-1])
            if pct <= 0 or pct > 100:
                raise ValueError
            n = max(1, round(len(table) * pct / 100))
        else:
            n = int(n_text) if n_text else 5
            if n <= 0 or n > len(table):
                raise ValueError
    except ValueError:
        n = min(5, len(table))

    result = table.sample(n=n)
    print(f"\n{UI_INDENT_1} Amostra aleatória de {n} linha(s):\n")
    draw_data_table(result)
    pause()


# Estatísticas: etiquetas, métodos individuais e menu
STAT_LABELS = {
    "count":  "Nº de valores",
    "mean":   "Média",
    "std":    "Desvio padrão",
    "min":    "Mínimo",
    "25%":    "1º Quartil (25%)",
    "50%":    "Mediana (50%)",
    "75%":    "3º Quartil (75%)",
    "max":    "Máximo",
    "unique": "Valores únicos",
    "top":    "Valor mais frequente",
    "freq":   "Frequência do top",
    "first":  "Data mais antiga",
    "last":   "Data mais recente",
}


def _summary_all_numeric():
    numeric_df = table.select_dtypes(include="number")
    if numeric_df.columns.empty:
        print(f"{UI_INDENT_1}Sem colunas numéricas.")
        pause()
        return
    desc = numeric_df.describe()
    widths = {c: max(len(str(c)), 10) for c in desc.columns}
    print()
    print(f"{UI_INDENT_1}Resumo estatístico de todas as colunas numéricas:\n")
    print(UI_INDENT_2 + " " * 22 + " " + " ".join(str(c).rjust(widths[c]) for c in desc.columns))
    for stat in desc.index:
        cells = " ".join(format_number(float(desc.at[stat, c])).rjust(widths[c]) for c in desc.columns)
        print(UI_INDENT_2 + STAT_LABELS.get(stat, stat).ljust(22) + " " + cells)
    pause()


def stat_summary():
    col = ask_column_or_all()
    if col is None:
        return
    if col is _ALL:
        _summary_all_numeric()
        return

    desc = table[col].describe()
    col_kind = "numérica" if pd.api.types.is_numeric_dtype(table[col]) else (
        "data" if pd.api.types.is_datetime64_any_dtype(table[col]) else "texto"
    )
    print()
    print(f"{UI_INDENT_1}Resumo estatístico de '{col}' (coluna {col_kind}):\n")

    for stat, val in desc.items():
        label = STAT_LABELS.get(stat, stat)
        if pd.isna(val):
            value_str = "--"
        elif hasattr(val, "strftime"):
            value_str = val.strftime("%d/%m/%Y")
        elif isinstance(val, float):
            value_str = format_number(val)
        else:
            value_str = str(val)
        print(UI_INDENT_2 + label.ljust(22) + " " + value_str)
    pause()


def stat_basic():
    col = ask_numeric_column()
    if col is None:
        return
    print()
    print(f"{UI_INDENT_1}Estatísticas base de '{col}':")
    print(f"{UI_INDENT_2}Média:   {format_number(table[col].mean())}")
    print(f"{UI_INDENT_2}Soma:    {format_number(table[col].sum())}")
    print(f"{UI_INDENT_2}Mínimo:  {format_number(table[col].min())}")
    print(f"{UI_INDENT_2}Máximo:  {format_number(table[col].max())}")
    print(f"{UI_INDENT_2}Mediana: {format_number(table[col].median())}")
    pause()


def stat_std_var():
    col = ask_numeric_column()
    if col is None:
        return
    print()
    print(f"{UI_INDENT_1}Desvio padrão de '{col}': {format_number(table[col].std())}")
    print(f"{UI_INDENT_1}Variância de '{col}':    {format_number(table[col].var())}")
    pause()


def stat_mode():
    col = ask_column()
    if col is None:
        return
    modes = table[col].mode()
    print()
    if modes.empty:
        print(f"{UI_INDENT_1}Sem moda calculável.")
    else:
        values = ", ".join(str(m) for m in modes.tolist())
        print(f"{UI_INDENT_1}Moda de '{col}': {values}")
    pause()


def stat_quantiles():
    col = ask_numeric_column()
    if col is None:
        return
    percentiles_text = input(f"{UI_INDENT_1}Percentis separados por vírgula (ex: 0.25,0.5,0.75) [Enter = quartis]: ").strip()
    if not percentiles_text:
        percentiles = [0.25, 0.5, 0.75]
    else:
        try:
            percentiles = [float(p.strip()) for p in percentiles_text.split(",")]
            if any(p < 0 or p > 1 for p in percentiles):
                raise ValueError
        except ValueError:
            print(f"{UI_INDENT_1}Percentis inválidos (devem ser entre 0 e 1). A usar quartis.")
            percentiles = [0.25, 0.5, 0.75]

    quantiles = table[col].quantile(percentiles)
    print()
    print(f"{UI_INDENT_1}Percentis de '{col}':")
    for p, v in quantiles.items():
        print(UI_INDENT_2 + f"{p * 100:.1f}%".rjust(6) + ": " + format_number(v))
    pause()


def stat_counts():
    col = ask_column()
    if col is None:
        return
    n_text = input(f"{UI_INDENT_1}Top N (Enter = 10): ").strip()
    try:
        n = int(n_text) if n_text else 10
        if n <= 0:
            raise ValueError
    except ValueError:
        n = 10

    counts = table[col].value_counts().head(n)
    print()
    print(f"{UI_INDENT_1}Top {n} valores em '{col}':\n")
    if counts.empty:
        print(f"{UI_INDENT_2}(vazio)")
    else:
        print_pandas_table(counts)
    pause()


def stat_unique():
    col = ask_column_or_all()
    if col is None:
        return
    if col is _ALL:
        print()
        print(f"{UI_INDENT_1}Valores únicos por coluna:\n")
        for c in table.columns:
            print(UI_INDENT_2 + str(c).ljust(22) + " " + f"{table[c].nunique()} (de {int(table[c].notna().sum())} não-nulos)")
        pause()
        return
    n_unique = table[col].nunique()
    n_total = int(table[col].notna().sum())
    print()
    print(f"{UI_INDENT_1}Valores únicos em '{col}': {n_unique} (de {n_total} não-nulos)")
    pause()


def stat_skew_kurt():
    col = ask_numeric_column()
    if col is None:
        return
    print()
    print(f"{UI_INDENT_1}Assimetria (skew) de '{col}': {format_number(table[col].skew())}")
    print(f"{UI_INDENT_1}Curtose (kurt) de '{col}':    {format_number(table[col].kurt())}")
    pause()



def statistics_menu():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    while True:
        header("ESTATÍSTICAS")
        print(f"{UI_INDENT_1}[1] Describe (resumo completo)")
        print(f"{UI_INDENT_1}[2] Estatísticas base (média/soma/mín/máx/mediana)")
        print(f"{UI_INDENT_1}[3] Desvio padrão + variância")
        print(f"{UI_INDENT_1}[4] Moda")
        print(f"{UI_INDENT_1}[5] Quartis / Percentis")
        print(f"{UI_INDENT_1}[6] Top N valores mais frequentes")
        print(f"{UI_INDENT_1}[7] Nº de valores únicos")
        print(f"{UI_INDENT_1}[8] Assimetria + Curtose")
        print(f"{UI_INDENT_1}[9] Correlação (Pearson e Spearman)")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"})

        if choice == "0":
            break
        elif choice == "1":
            stat_summary()
        elif choice == "2":
            stat_basic()
        elif choice == "3":
            stat_std_var()
        elif choice == "4":
            stat_mode()
        elif choice == "5":
            stat_quantiles()
        elif choice == "6":
            stat_counts()
        elif choice == "7":
            stat_unique()
        elif choice == "8":
            stat_skew_kurt()
        elif choice == "9":
            stat_correlation()


# Análise avançada: info do DataFrame, groupby, duplicados, missing data
def info_summary():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Sem folha.")
        pause()
        return

    header("INFO / RESUMO")
    print()
    print(f"{UI_INDENT_1}Dimensões: {len(table)} linhas × {len(table.columns)} colunas")
    print(f"{UI_INDENT_1}Memória:   {table.memory_usage(deep=True).sum() / 1024:.2f} KB")
    print()
    print(UI_INDENT_1 + "Coluna".ljust(22) + " " + "Tipo".ljust(14) + " " + "Não-nulos".rjust(12) + " " + "Missing %".rjust(11))
    print(UI_INDENT_1 + "-" * 62)
    n_total = len(table)
    for col in table.columns:
        n_non_null = int(table[col].notna().sum())
        dtype_str = str(table[col].dtype)
        miss_pct = (1 - n_non_null / n_total) * 100 if n_total > 0 else 0.0
        print(UI_INDENT_1 + str(col).ljust(22) + " " + dtype_str.ljust(14) + " " + str(n_non_null).rjust(12) + " " + f"{miss_pct:.1f}%".rjust(11))
    pause()


def group_by():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    header("AGRUPAR (groupby)")
    print()

    group_col = ask_column("Coluna para agrupar")
    if group_col is None:
        return

    num_col = ask_numeric_column("Coluna numérica a agregar")
    if num_col is None:
        return

    print()
    print(f"{UI_INDENT_1}Agregação: [1] Soma  [2] Média  [3] Contagem  [4] Mínimo  [5] Máximo")
    op = input(f"{UI_INDENT_1}Opção: ").strip()
    agg_map = {"1": "sum", "2": "mean", "3": "count", "4": "min", "5": "max"}
    agg = agg_map.get(op, "mean")

    result = table.groupby(group_col)[num_col].agg(agg).reset_index()
    result.columns = [group_col, f"{num_col} ({agg})"]
    print()
    print(f"{UI_INDENT_1}groupby('{group_col}')['{num_col}'].{agg}():\n")
    draw_data_table(result)
    pause()


def duplicates():
    global table
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    header("DUPLICADOS")
    print()

    dup_mask = table.duplicated(keep=False)
    n_dup = int(dup_mask.sum())

    if n_dup == 0:
        print(f"{UI_INDENT_1}Nenhum duplicado encontrado.")
        pause()
        return

    print(f"{UI_INDENT_1}{n_dup} linha(s) duplicada(s):\n")
    draw_data_table(table[dup_mask])
    print()

    if confirm("Remover duplicados (mantém primeira ocorrência)?"):
        n_before = len(table)
        table = table.drop_duplicates(ignore_index=True)
        apply_types()
        commit()
        print(f"\n{UI_INDENT_1}{n_before - len(table)} duplicado(s) removido(s).")
    else:
        print(f"\n{UI_INDENT_1}Nenhuma alteração.")
    pause()


def value_for_column(col, value):
    """Ajusta um valor calculado (média, mediana) ao tipo da coluna.

    Colunas inteiras não aceitam fracções, por isso o valor é arredondado.
    """
    if pd.isna(value):
        return value
    if pd.api.types.is_integer_dtype(table[col]):
        return round(value)
    return value


def missing_data():
    global table
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    header("DADOS EM FALTA")
    print()

    counts = table.isna().sum()
    n_total = len(table)

    print(f"{UI_INDENT_1}Missing por coluna (total {n_total} linhas):\n")
    for col in table.columns:
        n_miss = int(counts[col])
        pct = (n_miss / n_total * 100) if n_total else 0
        print(UI_INDENT_2 + str(col).ljust(22) + " " + str(n_miss).rjust(5) + f"{UI_INDENT_1} (" + f"{pct:.1f}%".rjust(6) + ")")
    print()

    if int(counts.sum()) == 0:
        print(f"{UI_INDENT_1}Sem missing data. Nada a tratar.")
        pause()
        return

    print(f"{UI_INDENT_1}Acções:")
    print(f"{UI_INDENT_2}[1] Preencher missing (fillna)")
    print(f"{UI_INDENT_2}[2] Remover linhas com missing (dropna)")
    print(f"{UI_INDENT_2}[0] Voltar")
    op = input(f"{UI_INDENT_1}Opção: ").strip()

    if op == "1":
        col = ask_column("Coluna a preencher")
        if col is None:
            return
        print()
        print(f"{UI_INDENT_2}[1] Valor específico")
        print(f"{UI_INDENT_2}[2] Média (se numérica)")
        print(f"{UI_INDENT_2}[3] Mediana (se numérica)")
        print(f"{UI_INDENT_2}[4] Forward-fill (ffill)")
        sub_choice = input(f"{UI_INDENT_1}Opção: ").strip()

        if sub_choice == "1":
            value = input(f"{UI_INDENT_1}Valor a usar: ").strip()
            col_type = type_of(col)
            try:
                table[col] = table[col].fillna(convert_value(value, col_type))
            except (ValueError, TypeError, OverflowError) as e:
                print(f"{UI_INDENT_1}Erro: {e}")
                pause()
                return
        elif sub_choice == "2":
            if not pd.api.types.is_numeric_dtype(table[col]):
                print(f"{UI_INDENT_1}Coluna não numérica.")
                pause()
                return
            table[col] = table[col].fillna(value_for_column(col, table[col].mean()))
        elif sub_choice == "3":
            if not pd.api.types.is_numeric_dtype(table[col]):
                print(f"{UI_INDENT_1}Coluna não numérica.")
                pause()
                return
            table[col] = table[col].fillna(value_for_column(col, table[col].median()))
        elif sub_choice == "4":
            table[col] = table[col].ffill()
        else:
            return

        apply_types()
        commit()
        print(f"\n{UI_INDENT_1}Coluna '{col}' preenchida.")
    elif op == "2":
        if not confirm(f"Remover linhas com qualquer missing ({int(counts.sum())} células)?"):
            return
        n_before = len(table)
        table = table.dropna().reset_index(drop=True)
        commit()
        print(f"\n{UI_INDENT_1}{n_before - len(table)} linha(s) removida(s).")
    pause()


# Gráficos (matplotlib + numpy): dispersão, área, histograma, boxplot, heatmap, tendência
plot_png_dir = None  # pasta onde gravar os gráficos em PNG; None = mostrá-los numa janela


PLOT_TEXT_MAX = 120
PLOT_ALL = "*"  # chave dos textos que valem para todos os gráficos da folha
_PLOT_FIELDS = {"title": "Título", "xlabel": "Rótulo do eixo X", "ylabel": "Rótulo do eixo Y"}
_PLOT_NAMES = {  # nome interno (o de finish_plot) -> nome mostrado
    "bars": "Barras", "lines": "Linhas", "pie": "Circular", "scatter": "Dispersão", "area": "Área", "histogram": "Histograma",
    "boxplot": "Boxplot", "correlation_heatmap": "Heatmap de correlação", "moving_average": "Linhas com média móvel",
    "polynomial_trend": "Tendência polinomial",
}


def clean_plot_text(text):
    """Texto de título/rótulo seguro para o matplotlib: sem espaços nas pontas, limitado, e sem '$' (que o matplotlib leria como fórmula)."""
    return " ".join(text.split())[:PLOT_TEXT_MAX].replace("$", r"\$")


def _sanitize_chart_labels(raw):
    """Aceita só textos de gráficos válidos ({gráfico ou '*': {campo: texto}}); o resto é ignorado."""
    clean = {}
    if not isinstance(raw, dict):
        return clean
    for key, fields in raw.items():
        if (key != PLOT_ALL and key not in _PLOT_NAMES) or not isinstance(fields, dict):
            continue
        texts = {f: clean_plot_text(v) for f, v in fields.items() if f in _PLOT_FIELDS and isinstance(v, str)}
        texts = {f: v for f, v in texts.items() if v}
        if texts:
            clean[key] = texts
    return clean


def chart_labels(name):
    """Textos a usar no gráfico `name` na folha activa: os de «todos os gráficos» e, por cima, os desse gráfico."""
    labels = dict(meta["charts"].get(PLOT_ALL, {}))
    labels.update(meta["charts"].get(name, {}))
    return labels


def apply_plot_labels(name):
    """Aplica ao gráfico aberto os textos guardados para a folha (o gráfico circular não tem eixos)."""
    labels = chart_labels(name)
    if labels and plt.get_fignums():
        axes = plt.gcf().axes[0]
        if "title" in labels:
            axes.set_title(labels["title"])
        if name != "pie" and "xlabel" in labels:
            axes.set_xlabel(labels["xlabel"])
        if name != "pie" and "ylabel" in labels:
            axes.set_ylabel(labels["ylabel"])


def _describe_plot_labels():
    """Linhas de texto com os títulos e rótulos guardados na folha activa."""
    lines = []
    for key, fields in meta["charts"].items():
        who = "todos os gráficos" if key == PLOT_ALL else _PLOT_NAMES[key].lower()
        lines.append(f"{who}: " + "; ".join(f"{_PLOT_FIELDS[f].lower()} «{v}»" for f, v in fields.items()))
    return lines


def set_plot_labels():
    """Define (ou limpa) o título e os rótulos dos eixos de um gráfico, ou de todos. Ficam gravados com a folha."""
    header("TÍTULO E RÓTULOS DOS GRÁFICOS")
    print(f"{UI_INDENT_1}Os textos ficam gravados com a folha e usam-se sempre que o gráfico for desenhado.\n")
    saved = _describe_plot_labels()
    if saved:
        print(f"{UI_INDENT_1}Guardados nesta folha:")
        for line in saved:
            print(f"{UI_INDENT_2}{line}")
        print()
    print(f"{UI_INDENT_1}[0] Todos os gráficos")
    for i, label in enumerate(_PLOT_NAMES.values(), 1):
        print(f"{UI_INDENT_1}[{i}] {label}")
    print(f"{UI_INDENT_1}[L] Apagar todos os textos guardados")
    choice = input(f"\n{UI_INDENT_1}Para que gráfico ('sair' cancela): ").strip().upper()
    if not choice or choice == "SAIR":
        return
    if choice == "L":
        if not meta["charts"]:
            print(f"\n{UI_INDENT_1}Não há textos guardados nesta folha.")
        elif confirm("Apagar os títulos e rótulos guardados nesta folha?"):
            meta["charts"] = {}
            commit()
            print(f"\n{UI_INDENT_1}Textos apagados. Os gráficos usam os textos automáticos.")
        else:
            print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return
    keys = [PLOT_ALL, *_PLOT_NAMES]
    if not choice.isdigit() or int(choice) >= len(keys):
        print(f"{UI_INDENT_1}Opção inválida.")
        pause()
        return
    key = keys[int(choice)]
    current = dict(meta["charts"].get(key, {}))
    print(f"\n{UI_INDENT_1}Enter = manter o que está, '-' = apagar (volta ao texto automático).\n")
    for field, label in _PLOT_FIELDS.items():
        shown = f" [{current[field]}]" if field in current else ""
        answer = input(f"{UI_INDENT_1}{label}{shown}: ").strip()
        if answer == "-":
            current.pop(field, None)
        elif answer:
            text = clean_plot_text(answer)
            if text:
                current[field] = text
    if current:
        meta["charts"][key] = current
    else:
        meta["charts"].pop(key, None)
    commit()
    print(f"\n{UI_INDENT_1}Guardado na folha '{active_name}'.")
    pause()


def finish_plot(name):
    """Termina um gráfico: mostra-o numa janela ou, com a gravação em PNG activa, grava-o e fecha-o."""
    apply_plot_labels(name)
    if plot_png_dir is None:
        plt.show()
        return
    base = os.path.join(plot_png_dir, f"{name}_{time.strftime('%Y%m%d_%H%M%S')}")
    path, n = base + ".png", 2
    while os.path.exists(path):  # dois gráficos no mesmo segundo não se sobrepõem
        path, n = f"{base}_{n}.png", n + 1
    try:
        plt.savefig(path, dpi=150, bbox_inches="tight")
        print(f"{UI_INDENT_1}Gráfico guardado em '{path}'.")
    except OSError as e:
        log.exception("Falha ao gravar o gráfico")
        print(f"{UI_INDENT_1}Erro ao gravar o gráfico: {e}")
    finally:
        plt.close("all")
    pause()


def toggle_png_export():
    global plot_png_dir
    if plot_png_dir is not None:
        plot_png_dir = None
        print(f"{UI_INDENT_1}Os gráficos voltam a abrir numa janela.")
        pause()
        return
    folder = input(f"{UI_INDENT_1}Pasta onde gravar os PNG (Enter = graficos): ").strip() or "graficos"
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError as e:
        print(f"{UI_INDENT_1}Não foi possível usar a pasta ({e}).")
        pause()
        return
    plot_png_dir = folder
    print(f"{UI_INDENT_1}Os gráficos passam a ser gravados em '{folder}' (sem abrir janela).")
    pause()


def plot_scatter():
    cols = ask_numeric_columns(
        min_n=2,
        max_n=2,
        message="Colunas X,Y para dispersão (ex.: idade,salario)"
    )
    if cols is None:
        return
    col_x, col_y = cols
    pairs = table[[col_x, col_y]].dropna()
    if pairs.empty:
        print(f"{UI_INDENT_1}Sem pares de valores válidos para dispersão.")
        pause()
        return

    x = pairs[col_x].to_numpy(dtype=float)
    y = pairs[col_y].to_numpy(dtype=float)
    plt.figure()
    plt.scatter(x, y, color="purple", alpha=0.7, edgecolor="black")
    plt.xlabel(col_x)
    plt.ylabel(col_y)
    plt.title(f"Dispersão: {col_x} vs {col_y}")
    plt.grid(linestyle="--", alpha=0.5)
    plt.tight_layout()
    finish_plot("scatter")


def plot_area():
    cols = ask_numeric_columns(min_n=1, message="Coluna(s) para área")
    if cols is None:
        return
    x = np.arange(len(table))
    plt.figure()
    for c in cols:
        vals = table[c].fillna(0).to_numpy(dtype=float)
        cum = np.cumsum(vals)
        plt.plot(x, cum, marker="o", label=c)
        plt.fill_between(x, cum, alpha=0.2)
    plt.xlabel("Índice da linha")
    plt.ylabel("Soma acumulada (np.cumsum)")
    plt.title("Evolução da soma acumulada")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    finish_plot("area")


# Gráficos analíticos: histograma, boxplot, heatmap, média móvel, tendência
def plot_histogram():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    result = ask_numeric_columns(min_n=1, max_n=1, message="Coluna para histograma")
    if result is None:
        return
    col = result[0]
    vals = table[col].dropna().to_numpy(dtype=float)
    if len(vals) == 0:
        print(f"{UI_INDENT_1}Sem valores numéricos.")
        pause()
        return
    q1, median, q3 = np.percentile(vals, [25, 50, 75])
    bins_auto = max(5, int(np.sqrt(len(vals))))
    bins_max = max(10, min(100, len(vals)))
    bins = ask_int_in_range(f"{UI_INDENT_1}Nº de bins [{2}-{bins_max}] (Enter = {bins_auto}): ", 2, bins_max, default=bins_auto)
    if bins is None:
        return
    plt.figure()
    plt.hist(vals, bins=bins, color="teal", edgecolor="black", alpha=0.75)
    plt.axvline(q1, color="orange", linestyle="--", label=f"Q1 ({format_number(q1)})")
    plt.axvline(median, color="red", linestyle="--", label=f"Mediana ({format_number(median)})")
    plt.axvline(q3, color="purple", linestyle="--", label=f"Q3 ({format_number(q3)})")
    plt.xlabel(col)
    plt.ylabel("Frequência")
    plt.title(f"Distribuição de '{col}'")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    finish_plot("histogram")


def plot_boxplot():
    cols = ask_numeric_columns(min_n=1, message="Coluna(s) para boxplot")
    if cols is None:
        return
    data_rows = [table[c].dropna().to_numpy(dtype=float) for c in cols]
    plt.figure()
    plt.boxplot(data_rows, tick_labels=cols)
    for i, arr in enumerate(data_rows, start=1):
        if len(arr) == 0:
            continue
        q1, _, q3 = np.percentile(arr, [25, 50, 75])
        plt.text(i, q3, f"Q3={format_number(q3)}", ha="center", va="bottom", fontsize=8)
        plt.text(i, q1, f"Q1={format_number(q1)}", ha="center", va="top", fontsize=8)
    plt.xlabel("Coluna")
    plt.ylabel("Valor")
    plt.title("Boxplot — quartis e mediana por coluna")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.tight_layout()
    finish_plot("boxplot")


def plot_correlation_heatmap():
    cols = ask_numeric_columns(min_n=2, message="Colunas para correlação")
    if cols is None:
        return
    data_rows = table[cols].dropna()
    if len(data_rows) < 2:
        print(f"{UI_INDENT_1}Necessário pelo menos 2 linhas com dados válidos nas colunas seleccionadas.")
        pause()
        return
    matrix = data_rows.corr().to_numpy()
    if np.isnan(matrix).all():
        print(f"{UI_INDENT_1}Correlação indefinida para as colunas seleccionadas.")
        pause()
        return
    n = len(cols)
    plt.figure()
    plt.imshow(matrix, cmap="coolwarm", vmin=-1, vmax=1)
    plt.colorbar(label="Correlação")
    positions = np.arange(n)
    plt.xticks(positions, cols, rotation=45, ha="right")
    plt.yticks(positions, cols)
    for i in range(n):
        for j in range(n):
            plt.text(j, i, f"{matrix[i, j]:.2f}", ha="center", va="center", color="black")
    plt.title("Heatmap de correlação")
    plt.tight_layout()
    finish_plot("correlation_heatmap")


def plot_moving_average():
    cols = ask_numeric_columns(min_n=1, message="Coluna(s) para média móvel")
    if cols is None:
        return
    n_rows = len(table)
    if n_rows < 2:
        print(f"{UI_INDENT_1}Necessário pelo menos 2 linhas.")
        pause()
        return
    max_window = min(50, n_rows)
    default_window = min(5, max_window)
    window = ask_int_in_range(f"{UI_INDENT_1}Janela da média móvel [2-{max_window}] (Enter = {default_window}): ", 2, max_window, default=default_window)
    if window is None:
        return
    kernel = np.ones(window) / window
    x = np.arange(n_rows)
    plt.figure()
    for c in cols:
        vals = table[c].to_numpy(dtype=float)
        if np.isnan(vals).all():
            print(f"{UI_INDENT_1}Coluna '{c}' sem valores válidos — ignorada.")
            continue
        if np.isnan(vals).any():
            vals = np.where(np.isnan(vals), np.nanmean(vals), vals)
        smoothed = np.convolve(vals, kernel, mode="same")
        plt.plot(x, smoothed, marker="o", label=c)
    plt.xlabel("Índice da linha")
    plt.ylabel("Valor suavizado")
    plt.title(f"Média móvel (np.convolve, janela={window})")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    finish_plot("moving_average")


def plot_polynomial_trend():
    cols = ask_numeric_columns(min_n=1, message="Coluna(s) para tendência")
    if cols is None:
        return
    n = len(table)
    if n < 2:
        print(f"{UI_INDENT_1}Necessário pelo menos 2 linhas.")
        pause()
        return
    max_degree = min(5, n - 1)
    default_degree = min(2, max_degree)
    degree = ask_int_in_range(f"{UI_INDENT_1}Grau do polinómio [1-{max_degree}] (Enter = {default_degree}): ", 1, max_degree, default=default_degree)
    if degree is None:
        return
    x = np.arange(n)
    plt.figure()
    for c in cols:
        y = table[c].to_numpy(dtype=float)
        if np.isnan(y).all():
            print(f"{UI_INDENT_1}Coluna '{c}' sem valores válidos — ignorada.")
            continue
        if np.isnan(y).any():
            y = np.where(np.isnan(y), np.nanmean(y), y)
        coefs = np.polyfit(x, y, degree)
        poly = np.poly1d(coefs)
        trend = poly(x)
        plt.scatter(x, y, alpha=0.3, s=20)
        plt.plot(x, trend, linewidth=2, label=c)
    plt.xlabel("Índice da linha")
    plt.ylabel("Valor")
    plt.title("Tendência polinomial (np.polyfit)")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    finish_plot("polynomial_trend")


# Gráficos com eixos personalizados (escolha de eixo X e agregações)
def plot_bars():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    print()
    print(f"{UI_INDENT_1}Eixo X: [1] Índice da linha  [2] Coluna da tabela")
    op_x = input(f"{UI_INDENT_1}Opção: ").strip()

    if op_x == "2":
        col_x = ask_column("Coluna para eixo X")
        if col_x is None:
            return
        col_y = ask_numeric_column("Coluna numérica para eixo Y")
        if col_y is None:
            return
        print()
        print(f"{UI_INDENT_1}Agregação: [1] Soma  [2] Média  [3] Contagem  [4] Mínimo  [5] Máximo")
        agg_map = {"1": "sum", "2": "mean", "3": "count", "4": "min", "5": "max"}
        agg = agg_map.get(input(f"{UI_INDENT_1}Opção: ").strip(), "mean")
        result = table.groupby(col_x)[col_y].agg(agg)
        x_vals = [str(v) for v in result.index]
        y_vals = result.values
        xlabel = str(col_x)
        ylabel = f"{col_y} ({agg})"
        title = f"Barras: {col_y} por {col_x} ({agg})"
    else:
        col_y = ask_numeric_column("Coluna numérica para eixo Y")
        if col_y is None:
            return
        x_vals = list(table.index.astype(str))
        y_vals = table[col_y].to_numpy(dtype=float)
        xlabel = "Índice"
        ylabel = str(col_y)
        title = f"Barras: {col_y} por índice"

    plt.figure()
    plt.bar(x_vals, y_vals, color="steelblue", edgecolor="black")
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(title)
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.tight_layout()
    finish_plot("bars")


def plot_lines():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    print()
    print(f"{UI_INDENT_1}Eixo X: [1] Índice da linha  [2] Coluna da tabela")
    op_x = input(f"{UI_INDENT_1}Opção: ").strip()

    if op_x == "2":
        col_x = ask_column("Coluna para eixo X")
        if col_x is None:
            return
        x_vals = table[col_x].tolist()
        xlabel = str(col_x)
    else:
        x_vals = list(table.index)
        xlabel = "Índice"

    cols_y = ask_numeric_columns(min_n=1, message="Coluna(s) para eixo Y")
    if cols_y is None:
        return

    plt.figure()
    for c in cols_y:
        y_vals = table[c].to_numpy(dtype=float)
        plt.plot(x_vals, y_vals, marker="o", label=c)
    plt.xlabel(xlabel)
    plt.ylabel("Valor")
    plt.title("Linhas — eixo X personalizado")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    finish_plot("lines")


def plot_pie():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return

    col_labels = ask_column("Coluna de categorias (labels)")
    if col_labels is None:
        return
    value_col = ask_numeric_column("Coluna numérica (valores)")
    if value_col is None:
        return

    print()
    print(f"{UI_INDENT_1}Agregação: [1] Soma  [2] Média  [3] Contagem")
    agg_map = {"1": "sum", "2": "mean", "3": "count"}
    agg = agg_map.get(input(f"{UI_INDENT_1}Opção: ").strip(), "sum")

    result = table.groupby(col_labels)[value_col].agg(agg)
    result = result[result > 0]
    if result.empty:
        print(f"{UI_INDENT_1}Sem valores positivos para desenhar.")
        pause()
        return

    plt.figure()
    plt.pie(result.values, labels=[str(label_text) for label_text in result.index], autopct="%1.1f%%", startangle=90)
    plt.title(f"Circular: {value_col} por {col_labels} ({agg})")
    plt.tight_layout()
    finish_plot("pie")


# Menu de gráficos: orquestra todas as visualizações disponíveis
def menu_charts():
    while True:
        header("GRÁFICOS (matplotlib + numpy)")
        print(f"{UI_INDENT_1}[1] Barras (eixos à escolha)")
        print(f"{UI_INDENT_1}[2] Linhas (eixo X à escolha)")
        print(f"{UI_INDENT_1}[3] Circular (categorias reais)")
        print(f"{UI_INDENT_1}[4] Dispersão (2 colunas)")
        print(f"{UI_INDENT_1}[5] Área (soma acumulada)")
        print(f"{UI_INDENT_1}[6] Histograma (1 coluna + quartis)")
        print(f"{UI_INDENT_1}[7] Boxplot (quartis e mediana)")
        print(f"{UI_INDENT_1}[8] Heatmap de correlação")
        print(f"{UI_INDENT_1}[9] Linhas com média móvel")
        print(f"{UI_INDENT_1}[A] Tendência polinomial")
        print(f"{UI_INDENT_1}[B] Gravar gráficos em PNG: {'ligado (' + plot_png_dir + ')' if plot_png_dir else 'desligado'}")
        saved = f" ({len(meta['charts'])} guardado(s) nesta folha)" if meta["charts"] else ""
        print(f"{UI_INDENT_1}[C] Título e rótulos dos gráficos{saved}")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "0"})
        if choice == "0":
            break
        elif choice == "1":
            plot_bars()
        elif choice == "2":
            plot_lines()
        elif choice == "3":
            plot_pie()
        elif choice == "4":
            plot_scatter()
        elif choice == "5":
            plot_area()
        elif choice == "6":
            plot_histogram()
        elif choice == "7":
            plot_boxplot()
        elif choice == "8":
            plot_correlation_heatmap()
        elif choice == "9":
            plot_moving_average()
        elif choice == "A":
            plot_polynomial_trend()
        elif choice == "B":
            toggle_png_export()
        elif choice == "C":
            set_plot_labels()


# Regras de validação por coluna: mínimo/máximo (números e datas), valores permitidos e obrigatório
def _is_blank(value):
    return (isinstance(value, str) and not value.strip()) or (not isinstance(value, str) and pd.isna(value))


def _rule_bound(raw, col_type):
    """Converte o limite guardado (texto) para o tipo da coluna. Levanta ValueError se for inválido."""
    return convert_value(str(raw), col_type) if col_type == DATE_TYPE else float(str(raw).replace(",", "."))


def _is_allowed(value, allowed, col_type):
    if col_type in (int, float):
        try:
            return any(float(value) == float(a.replace(",", ".")) for a in allowed)
        except ValueError:
            return False
    shown = format_value(value) if col_type == DATE_TYPE else str(value)
    return shown.lower() in [a.lower() for a in allowed]


def rule_violation(col, value):
    """Mensagem se `value` (já no tipo da coluna; vazio = em branco) violar as regras da coluna; None se estiver tudo bem."""
    rule = meta["rules"].get(col)
    if not rule:
        return None
    if _is_blank(value):
        return "o valor é obrigatório" if rule.get("required") else None
    col_type = type_of(col)
    allowed = rule.get("allowed")
    if allowed and not _is_allowed(value, allowed, col_type):
        return f"valor fora dos permitidos ({', '.join(allowed)})"
    if col_type in (int, float, DATE_TYPE):
        for key, label, violates in (("min", "mínimo", operator.lt), ("max", "máximo", operator.gt)):
            if rule.get(key) is None:
                continue
            try:
                if violates(value, _rule_bound(rule[key], col_type)):
                    return f"{'abaixo' if key == 'min' else 'acima'} do {label} ({rule[key]})"
            except (ValueError, TypeError):
                log.warning("Limite inválido na regra de %r: %r", col, rule[key])
    return None


def find_violations():
    """Todas as células da folha activa que violam regras: lista de (índice da linha, coluna, valor, mensagem)."""
    found = []
    for col, rule in meta["rules"].items():
        if col in table.columns and rule:
            for idx, value in table[col].items():
                problem = rule_violation(col, value)
                if problem:
                    found.append((idx, col, value, problem))
    return found


def _ask_rule_bound(label, col, col_type):
    """Pede um limite; Enter = sem limite. Devolve (texto, ok). ok=False se for inválido."""
    raw = input(f"{UI_INDENT_1}{label} de '{col}' ({'DD/MM/AAAA' if col_type == DATE_TYPE else 'número'}; Enter = sem limite): ").strip()
    if not raw:
        return None, True
    try:
        _rule_bound(raw, col_type)
    except ValueError:
        print(f"{UI_INDENT_1}Valor inválido.")
        pause()
        return None, False
    return raw, True


def set_rule():
    col = ask_column("Coluna a validar")
    if col is None:
        return
    col_type = type_of(col)
    rule = {}
    if col_type in (int, float, DATE_TYPE):
        low, ok = _ask_rule_bound("Mínimo", col, col_type)
        if not ok:
            return
        high, ok = _ask_rule_bound("Máximo", col, col_type)
        if not ok:
            return
        if low is not None and high is not None and _rule_bound(low, col_type) > _rule_bound(high, col_type):
            print(f"{UI_INDENT_1}O mínimo é maior do que o máximo.")
            pause()
            return
        rule.update({k: v for k, v in (("min", low), ("max", high)) if v is not None})
    allowed_text = input(f"{UI_INDENT_1}Valores permitidos, separados por vírgula (Enter = qualquer): ").strip()
    allowed = [a.strip() for a in allowed_text.split(",") if a.strip()]
    if allowed:
        if col_type is not str and any(not _is_valid_for_type(a, col_type) for a in allowed):
            print(f"{UI_INDENT_1}Há valores permitidos que não são do tipo da coluna.")
            pause()
            return
        rule["allowed"] = allowed
    if confirm("Valor obrigatório (não aceitar em branco)?"):
        rule["required"] = True
    if not rule:
        print(f"{UI_INDENT_1}Nenhuma regra definida.")
        pause()
        return
    meta["rules"][col] = rule
    commit()
    existing = [v for v in find_violations() if v[1] == col]
    if existing:
        print(f"\n{UI_INDENT_1}Regra guardada para '{col}'. {len(existing)} valor(es) existente(s) violam-na.")
    else:
        print(f"\n{UI_INDENT_1}Regra guardada para '{col}'. Os dados existentes cumprem-na.")
    pause()


def _is_valid_for_type(text, col_type):
    try:
        convert_value(text.replace(",", ".") if col_type in (int, float) else text, col_type)
        return True
    except (ValueError, TypeError):
        return False


def remove_rule():
    if not meta["rules"]:
        print(f"{UI_INDENT_1}Não há regras definidas.")
        pause()
        return
    col = ask_column("Coluna cujas regras remover")
    if col is None:
        return
    if meta["rules"].pop(col, None) is None:
        print(f"{UI_INDENT_1}'{col}' não tem regras.")
    else:
        commit()
        print(f"{UI_INDENT_1}Regras de '{col}' removidas.")
    pause()


def check_rules():
    violations = find_violations()
    print()
    if not violations:
        print(f"{UI_INDENT_1}Todos os valores cumprem as regras definidas.")
    else:
        print(f"{UI_INDENT_1}{len(violations)} violação(ões):")
        for idx, col, value, problem in violations[:20]:
            print(f"{UI_INDENT_2}linha {idx}, '{col}' = {format_value(value)}: {problem}")
        if len(violations) > 20:
            print(f"{UI_INDENT_2}... e mais {len(violations) - 20}.")
    pause()


def menu_rules():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    while True:
        header("REGRAS DE VALIDAÇÃO")
        if meta["rules"]:
            for col, rule in meta["rules"].items():
                parts = [f"{k}={v}" for k, v in rule.items() if k != "allowed"] + ([f"permitidos={'|'.join(rule['allowed'])}"] if rule.get("allowed") else [])
                print(f"{UI_INDENT_2}{col}: {', '.join(parts)}")
        else:
            print(f"{UI_INDENT_2}(sem regras)")
        print(f"\n{UI_INDENT_1}[1] Definir regra  [2] Remover regras de uma coluna  [3] Verificar a folha  [0] Voltar")
        print("-" * WIDTH)
        choice = ask_option({"1", "2", "3", "0"})
        if choice == "0":
            break
        elif choice == "1":
            set_rule()
        elif choice == "2":
            remove_rule()
        elif choice == "3":
            check_rules()


# Converter o tipo de uma coluna
def _plain_text(value):
    """Texto sem perdas de um valor: inteiros sem '.0', decimais com todos os dígitos, datas DD/MM/AAAA."""
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        return str(int(value)) if float(value).is_integer() else repr(float(value))
    return str(value)


def convert_series(series, source_type, target_type):
    """Converte `series` de `source_type` para `target_type`. Devolve (série nova, máscara das células com valor que falharam).

    Levanta ValueError se a combinação não fizer sentido (número <-> data).
    """
    numeric = (int, float)
    if (source_type == DATE_TYPE and target_type in numeric) or (source_type in numeric and target_type == DATE_TYPE):
        raise ValueError("não é possível converter entre números e datas (passe por texto)")
    present = series.notna()
    text = series.map(lambda v: _plain_text(v).strip() if pd.notna(v) else None)
    if target_type is str:
        return text.astype(object), pd.Series(False, index=series.index)
    if target_type == DATE_TYPE:
        converted = pd.to_datetime(text, format="%d/%m/%Y", errors="coerce")
        return converted, present & converted.isna()
    numbers = pd.to_numeric(text.str.replace(",", ".", regex=False), errors="coerce")
    ok = numbers.notna()
    if target_type is int:
        ok = ok & (numbers % 1 == 0)  # 1.5 não pode ser inteiro: falha em vez de arredondar em silêncio
        return numbers.where(ok).astype("Int64"), present & ~ok
    return numbers.astype("float64"), present & ~ok


def convert_column_type():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    header("CONVERTER TIPO DA COLUNA")
    print()
    col = ask_column("Coluna a converter")
    if col is None:
        return
    if col in meta["formulas"]:
        print(f"{UI_INDENT_1}'{col}' é uma coluna calculada; o tipo vem da fórmula.")
        pause()
        return
    source_type = type_of(col)
    print(f"{UI_INDENT_1}Tipo actual: {TYPE_NAME.get(source_type, 'texto')}")
    target_type = ask_type()
    if target_type == source_type:
        print(f"{UI_INDENT_1}A coluna já é desse tipo.")
        pause()
        return
    try:
        converted, failed = convert_series(table[col], source_type, target_type)
    except ValueError as e:
        print(f"{UI_INDENT_1}{e}.")
        pause()
        return
    n_failed = int(failed.sum())
    print(f"\n{UI_INDENT_1}{int(table[col].notna().sum()) - n_failed} valor(es) convertido(s), {n_failed} não convertido(s).")
    if n_failed:
        examples = ", ".join(f"linha {i}: '{format_value(table.at[i, col])}'" for i in list(failed[failed].index)[:5])
        print(f"{UI_INDENT_1}Não convertidos (ficam em branco): {examples}{' ...' if n_failed > 5 else ''}")
        if not confirm("Continuar mesmo assim?"):
            print(f"\n{UI_INDENT_1}Operação cancelada.")
            pause()
            return
    table[col] = converted
    col_types[col] = target_type
    if meta["rules"].pop(col, None) or meta["formats"].pop(col, None):
        print(f"{UI_INDENT_1}As regras e o formato de '{col}' foram removidos (já não se aplicam ao novo tipo).")
    apply_types()
    commit()
    print(f"\n{UI_INDENT_1}'{col}' é agora {TYPE_NAME.get(target_type, 'texto')}.")
    pause()


# Procurar e substituir
def _replace_in_text(series, find, replacement, whole_cell, case_sensitive):
    """Colunas de texto: troca parte do texto (ou a célula inteira, se `whole_cell`). Devolve (série nova, máscara)."""
    pattern = re.compile(("^" + re.escape(find) + "$") if whole_cell else re.escape(find), 0 if case_sensitive else re.IGNORECASE)
    text = series.astype(object)
    match = text.map(lambda v: isinstance(v, str) and pattern.search(v) is not None)
    new = text.map(lambda v: pattern.sub(lambda _: replacement, v) if isinstance(v, str) else v)
    return new, match


def _replace_whole_value(series, col_type, find, replacement, col):
    """Colunas numéricas e de datas: troca as células cujo valor é igual a `find`. Devolve (série nova, máscara)."""
    shown = series.map(lambda v: _plain_text(v) if pd.notna(v) else None)
    if col_type in (int, float):
        try:
            match = pd.to_numeric(shown, errors="coerce").eq(float(find.strip().replace(",", ".")))
        except ValueError:
            match = pd.Series(False, index=series.index)
    else:
        match = shown.eq(find.strip())
    if not match.any():
        return series, match
    if replacement.strip():
        try:
            new_value = convert_value(replacement.replace(",", ".") if col_type in (int, float) else replacement, col_type)
        except (ValueError, TypeError):
            raise ValueError(f"'{replacement}' não é válido para a coluna '{col}' ({TYPE_NAME.get(col_type, 'texto')})") from None
    else:
        new_value = None  # substituir por nada = deixar em branco
    return series.astype(object).where(~match, new_value), match


def plan_replace(find, replacement, columns, whole_cell=False, case_sensitive=False):
    """Calcula (sem alterar a folha) as substituições: devolve {coluna: (série nova, máscara das células alteradas)}.

    Em colunas de texto pode substituir parte do texto; nas numéricas e de datas só substitui células iguais.
    Levanta ValueError se `replacement` não for válido para o tipo de alguma coluna.
    """
    plan = {}
    for col in columns:
        col_type = type_of(col)
        if col_type is str:
            new, match = _replace_in_text(table[col], find, replacement, whole_cell, case_sensitive)
        else:
            new, match = _replace_whole_value(table[col], col_type, find, replacement, col)
        if match.any():
            plan[col] = (new, match)
    return plan


def replace_violations(plan):
    """Violações de regras que a aplicação de `plan` causaria: lista de (índice, coluna, valor, mensagem)."""
    found = []
    for col, (new, match) in plan.items():
        for idx in match[match].index:
            problem = rule_violation(col, new[idx])
            if problem:
                found.append((idx, col, new[idx], problem))
    return found


def find_replace():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    header("PROCURAR E SUBSTITUIR")
    print()
    scope = ask_column_or_all("Coluna")
    if scope is None:
        return
    columns = [c for c in table.columns if c not in meta["formulas"]] if scope is _ALL else [scope]
    if scope is not _ALL and scope in meta["formulas"]:
        print(f"{UI_INDENT_1}'{scope}' é uma coluna calculada e não pode ser alterada.")
        pause()
        return
    find = input(f"{UI_INDENT_1}Procurar ('sair' cancela): ")
    if not find.strip() or find.lower() == "sair":
        return
    replacement = input(f"{UI_INDENT_1}Substituir por (Enter = vazio): ")
    whole_cell = confirm("Só células inteiras iguais (em vez de parte do texto)?")
    case_sensitive = confirm("Distinguir maiúsculas de minúsculas?")
    try:
        plan = plan_replace(find, replacement, columns, whole_cell, case_sensitive)
    except ValueError as e:
        print(f"{UI_INDENT_1}{e}.")
        pause()
        return
    if not plan:
        print(f"\n{UI_INDENT_1}Nada encontrado.")
        pause()
        return
    print()
    for col, (_new, match) in plan.items():
        print(f"{UI_INDENT_2}{col}: {int(match.sum())} célula(s)")
    violations = replace_violations(plan)
    if violations:
        print(f"\n{UI_INDENT_1}A substituição violaria regras de validação ({len(violations)} célula(s)); nada foi alterado.")
        for idx, col, value, problem in violations[:5]:
            print(f"{UI_INDENT_2}linha {idx}, '{col}' = {format_value(value)}: {problem}")
        pause()
        return
    if not confirm(f"Substituir {sum(int(m.sum()) for _, m in plan.values())} célula(s)?"):
        print(f"\n{UI_INDENT_1}Operação cancelada.")
        pause()
        return
    for col, (new, _match) in plan.items():
        table[col] = new
    apply_types()
    commit()
    print(f"\n{UI_INDENT_1}Substituição concluída.")
    pause()


# Análise: correlação, outliers, normalização e faixas
CORRELATION_MIN_PAIRS = 3
IQR_FACTOR = 1.5
ZSCORE_LIMIT = 3.0


def correlation_strength(r):
    """Descrição em português da intensidade e do sentido de um coeficiente de correlação."""
    if pd.isna(r):
        return "indefinida"
    size = abs(r)
    for limit, label in ((0.8, "muito forte"), (0.6, "forte"), (0.4, "moderada"), (0.2, "fraca")):
        if size >= limit:
            return f"{label} {'positiva' if r > 0 else 'negativa'}"
    return "muito fraca ou nula"


def correlation(col_x, col_y, method="pearson"):
    """Coeficiente de correlação ("pearson" ou "spearman") entre duas colunas numéricas, só com as linhas onde ambas têm valor.

    Devolve (r, nº de pares). Levanta ValueError se houver menos de 3 pares ou se uma das colunas for constante.
    """
    pairs = table[[col_x, col_y]].dropna().astype("float64")
    if len(pairs) < CORRELATION_MIN_PAIRS:
        raise ValueError(f"são precisos pelo menos {CORRELATION_MIN_PAIRS} pares de valores (há {len(pairs)})")
    if pairs[col_x].nunique() < 2 or pairs[col_y].nunique() < 2:
        raise ValueError("uma das colunas não varia, por isso a correlação não está definida")
    if method == "spearman":
        pairs = pairs.rank()  # Spearman = Pearson sobre as ordens (evita depender do scipy)
    elif method != "pearson":
        raise ValueError(f"método desconhecido: {method}")
    return float(pairs[col_x].corr(pairs[col_y])), len(pairs)


def stat_correlation():
    cols = ask_numeric_columns(min_n=2, message="Colunas (2 para o detalhe, mais para a matriz), separadas por vírgula")
    if cols is None:
        return
    print()
    if len(cols) == 2:
        try:
            for method, label in (("pearson", "Pearson (relação linear)"), ("spearman", "Spearman (relação monótona, por ordens)")):
                r, n = correlation(cols[0], cols[1], method)
                print(f"{UI_INDENT_1}{label}: {r:.4f}  ->  {correlation_strength(r)}  ({n} pares)")
        except ValueError as e:
            print(f"{UI_INDENT_1}{e}.")
    else:
        print(f"{UI_INDENT_1}Matriz de correlação de Pearson:\n")
        print_pandas_table(table[cols].astype("float64").corr().round(3))
    pause()


def detect_outliers(col, method="iqr", factor=None):
    """Máscara (por linha da folha) dos valores anómalos de uma coluna numérica e os limites usados.

    "iqr": fora de [Q1 - k*IQR, Q3 + k*IQR] (k=1.5); "zscore": a mais de k desvios-padrão da média (k=3).
    Devolve (máscara, limite inferior, limite superior); valores em falta nunca são anómalos.
    """
    series = pd.to_numeric(table[col], errors="coerce").astype("float64")
    if method == "iqr":
        factor = IQR_FACTOR if factor is None else factor
        q1, q3 = series.quantile(0.25), series.quantile(0.75)
        low, high = q1 - factor * (q3 - q1), q3 + factor * (q3 - q1)
    elif method == "zscore":
        factor = ZSCORE_LIMIT if factor is None else factor
        mean, std = series.mean(), series.std()
        low, high = mean - factor * std, mean + factor * std
    else:
        raise ValueError(f"método desconhecido: {method}")
    if pd.isna(low) or pd.isna(high):
        return pd.Series(False, index=table.index), low, high
    return (series < low) | (series > high), low, high


def ask_float(message, default, minimum):
    """Pede um número decimal >= `minimum`; Enter = `default`. Devolve None se for inválido."""
    value = input(message).strip().replace(",", ".")
    if not value:
        return default
    try:
        number = float(value)
    except ValueError:
        number = float("nan")
    if not number >= minimum:  # também apanha NaN
        print(f"{UI_INDENT_1}Valor inválido (número >= {minimum:g} esperado).")
        pause()
        return None
    return number


def outliers():
    global table
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    header("VALORES ANÓMALOS (OUTLIERS)")
    print()
    col = ask_numeric_column("Coluna a analisar")
    if col is None:
        return
    print(f"{UI_INDENT_1}Método: [1] IQR (Q1 - k*IQR .. Q3 + k*IQR, k=1.5)  [2] Desvio-padrão (k=3)")
    method = {"1": "iqr", "2": "zscore"}.get(input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() or "1")
    if method is None:
        print(f"{UI_INDENT_1}Opção inválida.")
        pause()
        return
    default_factor = IQR_FACTOR if method == "iqr" else ZSCORE_LIMIT
    factor = ask_float(f"{UI_INDENT_1}Factor k (Enter = {default_factor:g}): ", default_factor, 0.1)
    if factor is None:
        return
    mask, low, high = detect_outliers(col, method, factor)
    print(f"\n{UI_INDENT_1}Intervalo normal: {format_number(float(low))} a {format_number(float(high))}")
    if not mask.any():
        print(f"{UI_INDENT_1}Nenhum valor anómalo em '{col}'.")
        pause()
        return
    print(f"{UI_INDENT_1}{int(mask.sum())} valor(es) anómalo(s):\n")
    draw_data_table(table[mask])
    print(f"\n{UI_INDENT_1}[1] Remover as linhas  [2] Deixar a célula em branco  [Enter] Só ver")
    action = input(f"{UI_INDENT_1}Opção: ").strip()
    if action == "1":
        table = table[~mask].reset_index(drop=True)
    elif action == "2":
        if col in meta["formulas"]:
            print(f"{UI_INDENT_1}'{col}' é uma coluna calculada; remova as linhas ou altere os dados de origem.")
            pause()
            return
        table.loc[mask, col] = pd.NA
    else:
        return
    apply_types()
    commit()
    print(f"\n{UI_INDENT_1}Feito.")
    pause()


def normalize_series(series, method):
    """Normaliza uma série numérica: "minmax" -> [0, 1]; "zscore" -> média 0 e desvio-padrão 1. Levanta ValueError se não variar."""
    values = pd.to_numeric(series, errors="coerce").astype("float64")
    if method == "minmax":
        span = values.max() - values.min()
        if pd.isna(span) or span == 0:
            raise ValueError("a coluna não varia (todos os valores são iguais ou está vazia)")
        return (values - values.min()) / span
    if method == "zscore":
        std = values.std()
        if pd.isna(std) or std == 0:
            raise ValueError("a coluna não varia (todos os valores são iguais ou está vazia)")
        return (values - values.mean()) / std
    raise ValueError(f"método desconhecido: {method}")


def _ask_new_column_name(default):
    """Pede o nome de uma coluna nova (Enter = `default`, que é tornado único). None se cancelar."""
    candidate, n = default, 2
    while resolve_column(candidate) is not None:
        candidate, n = f"{default}_{n}", n + 1
    name = input(f"{UI_INDENT_1}Nome da nova coluna (Enter = {candidate}): ").strip()
    if name.lower() == "sair":
        return None
    name = name or candidate
    if resolve_column(name) is not None:
        print(f"{UI_INDENT_1}Coluna '{name}' já existe.")
        pause()
        return None
    return name


def normalize_column():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    header("NORMALIZAR COLUNA")
    print()
    col = ask_numeric_column("Coluna a normalizar")
    if col is None:
        return
    print(f"{UI_INDENT_1}[1] Min-máx (0 a 1)  [2] Padronização z-score (média 0, desvio 1)")
    method = {"1": "minmax", "2": "zscore"}.get(input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() or "1")
    if method is None:
        print(f"{UI_INDENT_1}Opção inválida.")
        pause()
        return
    try:
        result = normalize_series(table[col], method)
    except ValueError as e:
        print(f"{UI_INDENT_1}{e}.")
        pause()
        return
    name = _ask_new_column_name(f"{col}_{'norm' if method == 'minmax' else 'z'}")
    if name is None:
        return
    table[name] = result
    col_types[name] = float
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' criada (a original não muda).")
    pause()


BIN_MAX = 50


def make_bins(series, mode, spec):
    """Divide uma coluna numérica em faixas. Devolve (rótulos por linha ou em branco fora das faixas, limites).

    Faixas (a, b]; a primeira inclui o mínimo. `mode`: "edges" (spec = lista de limites crescentes), "equal" (spec = nº de faixas
    de igual largura) ou "quantile" (spec = nº de faixas com igual nº de linhas). Levanta ValueError se os parâmetros forem inválidos.
    """
    values = pd.to_numeric(series, errors="coerce").astype("float64")
    if mode == "edges":
        try:
            edges = [float(e) for e in spec]
        except (TypeError, ValueError):
            raise ValueError("os limites têm de ser números (com ponto decimal)") from None
        if len(edges) < 2 or any(b <= a for a, b in zip(edges, edges[1:], strict=False)):
            raise ValueError("indique pelo menos 2 limites, por ordem crescente e sem repetições")
    elif mode in ("equal", "quantile"):
        if not isinstance(spec, int) or not 2 <= spec <= BIN_MAX:
            raise ValueError(f"o número de faixas tem de estar entre 2 e {BIN_MAX}")
        if values.notna().sum() < 2:
            raise ValueError("a coluna não tem valores suficientes")
        edges = list(np.linspace(values.min(), values.max(), spec + 1)) if mode == "equal" else list(np.unique(values.quantile(np.linspace(0, 1, spec + 1))))
        if len(edges) < 2 or edges[0] == edges[-1]:
            raise ValueError("a coluna não varia, não há faixas possíveis")
    else:
        raise ValueError(f"modo desconhecido: {mode}")
    labels = [f"{format_number(a)} a {format_number(b)}" for a, b in zip(edges, edges[1:], strict=False)]
    return pd.cut(values, bins=edges, labels=labels, include_lowest=True).astype(object), edges


def bin_column():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    header("FAIXAS (INTERVALOS)")
    print()
    col = ask_numeric_column("Coluna a dividir em faixas")
    if col is None:
        return
    print(f"{UI_INDENT_1}[1] Limites à escolha (ex.: 0,18,30,65)  [2] N faixas de igual largura  [3] N faixas com igual nº de linhas")
    option = input(f"{UI_INDENT_1}Opção: ").strip()
    try:
        if option == "1":
            raw = input(f"{UI_INDENT_1}Limites separados por vírgula: ").replace(";", ",")
            result, edges = make_bins(table[col], "edges", [x.strip().replace(" ", "") for x in raw.split(",") if x.strip()])
        elif option in ("2", "3"):
            count = ask_int_in_range(f"{UI_INDENT_1}Número de faixas (2-{BIN_MAX}): ", 2, BIN_MAX)
            if count is None:
                return
            result, edges = make_bins(table[col], "equal" if option == "2" else "quantile", count)
        else:
            return
    except ValueError as e:
        print(f"{UI_INDENT_1}{e}.")
        pause()
        return
    outside = int(table[col].notna().sum() - result.notna().sum())
    name = _ask_new_column_name(f"{col}_faixa")
    if name is None:
        return
    table[name] = result
    col_types[name] = str
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' criada com {len(edges) - 1} faixa(s) (a, b]; a primeira inclui o mínimo.")
    if outside:
        print(f"{UI_INDENT_1}{outside} valor(es) ficaram fora dos limites e estão em branco.")
    pause()


# Colunas calculadas (fórmulas)
FORMULA_HELP = (
    "Operadores: + - * / // % **   comparações: == != < <= > >=   lógica: and or not",
    "Colunas: escreva o nome (Salario) ou entre parênteses rectos se tiver espaços ([Preço unitário])",
    "Funções: abs round(x, casas) sqrt log exp floor ceil min(a,b) max(a,b) se(condição, então, senão)",
    "Datas:   year(d) month(d) day(d) days(d1, d2)   Exemplos: Preco * Qtd   se(Idade >= 18, 1, 0)   round(Salario / 12, 2)",
    "Células: B2 = coluna B (letras abaixo), linha 2 (o nº da coluna # do ecrã); $B$2 é igual. Intervalos: B2:B9",
    "Totais:  soma(B2:B9)  media(Salario)  contar(Nome)  min(B2:B9)  max(Salario)   Exemplo: Salario / soma(Salario) * 100",
)


def add_formula_column():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    header("COLUNA CALCULADA (FÓRMULA)")
    print()
    for line in FORMULA_HELP:
        print(f"{UI_INDENT_1}{line}")
    print(f"{UI_INDENT_1}Letras:   " + "  ".join(f"{column_letter(i)}={c}" for i, c in enumerate(table.columns)))
    if meta["formulas"]:
        print(f"\n{UI_INDENT_1}Fórmulas existentes:")
        for col, expression in meta["formulas"].items():
            print(f"{UI_INDENT_2}{col} = {expression}")
    print()
    name = input(f"{UI_INDENT_1}Nome da nova coluna ('sair' cancela): ").strip()
    if not name or name.lower() == "sair":
        return
    if resolve_column(name) is not None:
        print(f"{UI_INDENT_1}Coluna '{name}' já existe.")
        pause()
        return
    while True:
        expression = input(f"{UI_INDENT_1}{name} = ").strip()
        if not expression or expression.lower() == "sair":
            return
        try:
            expression = canonical_formula(expression, list(table.columns))
            series, col_type = evaluate_formula(expression, table, col_types)
            break
        except FormulaError as e:
            print(f"{UI_INDENT_1}Fórmula inválida: {e}")

    table[name] = series
    col_types[name] = col_type
    meta["formulas"][name] = expression
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' ({TYPE_NAME.get(col_type, 'texto')}) calculada. É recalculada sempre que a folha muda.")
    print(f"{UI_INDENT_1}Guardada como: {expression}")
    pause()


def freeze_formula():
    if not meta["formulas"]:
        print(f"{UI_INDENT_1}Não há colunas calculadas.")
        pause()
        return
    header("CONVERTER FÓRMULA EM VALORES")
    print()
    for col, expression in meta["formulas"].items():
        print(f"{UI_INDENT_2}{col} = {expression}")
    value = input(f"\n{UI_INDENT_1}Coluna ('sair' cancela): ").strip()
    if not value or value.lower() == "sair":
        return
    col = resolve_column(value, source=list(meta["formulas"]))
    if col is None:
        print(f"{UI_INDENT_1}Essa coluna não tem fórmula.")
        pause()
        return
    users = formula_users(col)
    if users:
        print(f"{UI_INDENT_1}'{col}' é usada pela fórmula de: {', '.join(users)}. Converta essas primeiro.")
        pause()
        return
    if not confirm(f"Converter '{col}' em valores fixos (deixa de ser recalculada)?"):
        return
    del meta["formulas"][col]
    commit()
    print(f"\n{UI_INDENT_1}'{col}' é agora uma coluna normal.")
    pause()


# Livro de folhas: escolher, adicionar, renomear, eliminar, duplicar, juntar, empilhar e tabela dinâmica
def ask_sheet(message, exclude=None):
    """Pede uma folha pelo nome ou número. Devolve o nome, ou None se cancelar/inválido."""
    names = [n for n in sheet_names() if n != exclude]
    print(f"{UI_INDENT_1}Folhas: " + ", ".join(f"[{sheet_names().index(n) + 1}] {n}" for n in names))
    value = input(f"{UI_INDENT_1}{message} (nome ou nº, 'sair' cancela): ").strip()
    if not value or value.lower() == "sair":
        return None
    if value.isdigit() and 1 <= int(value) <= len(sheet_names()) and sheet_names()[int(value) - 1] in names:
        return sheet_names()[int(value) - 1]
    chosen = resolve_column(value, source=names)
    if chosen is None:
        print(f"{UI_INDENT_1}Folha não existe.")
        pause()
    return chosen


def _ask_sheet_name(message):
    """Pede um nome de folha válido e livre. Devolve None se cancelar."""
    while True:
        name = input(f"{UI_INDENT_1}{message} ('sair' cancela): ").strip()
        if not name or name.lower() == "sair":
            return None
        error = sheet_name_error(name)
        if error is None:
            return name
        print(f"{UI_INDENT_1}Nome inválido: {error}.")


def switch_sheet_ui():
    name = ask_sheet("Mudar para")
    if name is not None and name != active_name:
        switch_sheet(name)
        commit()


def add_blank_sheet():
    name = _ask_sheet_name("Nome da nova folha")
    if name is None:
        return
    previous = active_name
    add_sheet(name)
    new_sheet()
    if table.columns.empty:  # o utilizador cancelou a definição das colunas: não deixa uma folha vazia para trás
        switch_sheet(previous)
        delete_sheet(name)
        save_state()


def rename_sheet_ui():
    old = ask_sheet("Folha a renomear")
    if old is None:
        return
    new = _ask_sheet_name(f"Novo nome para '{old}'")
    if new is None:
        return
    try:
        rename_sheet(old, new)
    except ValueError as e:
        print(f"{UI_INDENT_1}Nome inválido: {e}.")
        pause()
        return
    commit()
    print(f"\n{UI_INDENT_1}'{old}' renomeada para '{new}'.")
    pause()


def delete_sheet_ui():
    if len(sheet_names()) <= 1:
        print(f"{UI_INDENT_1}O livro tem de ter pelo menos uma folha.")
        pause()
        return
    name = ask_sheet("Folha a eliminar")
    if name is None or not confirm(f"Eliminar a folha '{name}' e os seus dados?"):
        return
    delete_sheet(name)
    commit()
    print(f"\n{UI_INDENT_1}Folha '{name}' eliminada (pode desfazer com [U]).")
    pause()


def duplicate_sheet_ui():
    source = ask_sheet("Folha a duplicar")
    if source is None:
        return
    new = _ask_sheet_name("Nome da cópia")
    if new is None:
        return
    duplicate_sheet(source, new)
    commit()
    print(f"\n{UI_INDENT_1}Cópia '{new}' criada e activa.")
    pause()


_MERGE_KINDS = {"1": ("inner", "só as linhas com chave nas duas folhas"), "2": ("left", "todas as da folha activa"), "3": ("right", "todas as da outra folha"), "4": ("outer", "todas")}
MAX_MERGE_ROWS = 1_000_000


def merge_tables(left_name, right_name, left_key, right_key, how):
    """Junta duas folhas por uma coluna chave. Devolve (tabela, tipos). Levanta ValueError se não for possível.

    As colunas da direita com o nome de uma da esquerda recebem o sufixo `_<folha da direita>`.
    """
    left, right = get_sheet(left_name), get_sheet(right_name)
    left_types, right_types = left["types"], right["types"]
    numeric = (int, float)
    kl, kr = left_types.get(left_key, str), right_types.get(right_key, str)
    if kl != kr and not (kl in numeric and kr in numeric):
        raise ValueError(f"as chaves têm tipos diferentes ({TYPE_NAME.get(kl)} e {TYPE_NAME.get(kr)}); converta uma delas primeiro")

    taken = {str(c).lower() for c in left["table"].columns}
    rename, new_right_names = {}, []
    for col in right["table"].columns:
        if col == right_key and left_key == right_key:
            new_right_names.append(col)
            continue
        new = col
        if col.lower() in taken:
            new = f"{col}_{right_name}"
            while new.lower() in taken:
                new += "_"
        rename[col] = new
        taken.add(new.lower())
        new_right_names.append(new)
    right_data = right["table"].rename(columns=rename)
    right_key_name = rename.get(right_key, right_key)

    left_counts = left["table"].groupby(left_key).size()
    right_counts = right_data.groupby(right_key_name).size()
    pairs = int((left_counts.reindex(right_counts.index, fill_value=0) * right_counts).sum()) + len(left["table"]) + len(right_data)
    if pairs > MAX_MERGE_ROWS:
        raise ValueError("a junção produziria demasiadas linhas (a chave tem muitos valores repetidos nas duas folhas)")

    merged = pd.merge(left["table"], right_data, how=how, left_on=left_key, right_on=right_key_name)
    types = dict(left_types)
    types.update({rename.get(c, c): right_types.get(c, str) for c in right["table"].columns})
    if right_key_name == left_key:
        types[left_key] = kl if kl == kr else float
    elif kl != kr:
        types[left_key] = types[right_key_name] = float
    merged = typed_frame(merged, types)
    return merged, {c: types.get(c, infer_type(merged[c])) for c in merged.columns}


def merge_sheets_ui():
    if len(sheet_names()) < 2:
        print(f"{UI_INDENT_1}São precisas pelo menos duas folhas (adicione ou importe outra).")
        pause()
        return
    header("JUNTAR TABELAS (JOIN)")
    print(f"{UI_INDENT_1}Folha da esquerda: '{active_name}'\n")
    right_name = ask_sheet("Folha da direita", exclude=active_name)
    if right_name is None:
        return
    print(f"{UI_INDENT_1}Colunas de '{active_name}': {', '.join(table.columns)}")
    left_key = resolve_column(input(f"{UI_INDENT_1}Coluna chave em '{active_name}': ").strip())
    right_columns = list(get_sheet(right_name)["table"].columns)
    print(f"{UI_INDENT_1}Colunas de '{right_name}': {', '.join(right_columns)}")
    right_key = resolve_column(input(f"{UI_INDENT_1}Coluna chave em '{right_name}': ").strip(), source=right_columns)
    if left_key is None or right_key is None:
        print(f"{UI_INDENT_1}Coluna não existe.")
        pause()
        return
    print(f"{UI_INDENT_1}Manter: " + "  ".join(f"[{k}] {desc}" for k, (_, desc) in _MERGE_KINDS.items()))
    how = _MERGE_KINDS.get(input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() or "1")
    if how is None:
        print(f"{UI_INDENT_1}Opção inválida.")
        pause()
        return
    try:
        merged, types = merge_tables(active_name, right_name, left_key, right_key, how[0])
    except ValueError as e:
        print(f"\n{UI_INDENT_1}{e}.")
        pause()
        return
    name = unique_sheet_name(f"{active_name}+{right_name}")
    add_sheet(name, merged, types)
    commit()
    print(f"\n{UI_INDENT_1}Criada a folha '{name}' com {len(merged)} linhas e {len(merged.columns)} colunas (as originais não mudam).")
    pause()


def concat_tables(first_name, second_name):
    """Empilha as linhas de duas folhas. As colunas são unidas pelo nome (sem distinguir maiúsculas); as que faltam ficam em branco.

    Se uma coluna tiver tipos diferentes: inteiro+decimal dá decimal, qualquer outro caso dá texto.
    Devolve (tabela, tipos, colunas convertidas para texto).
    """
    sources = [get_sheet(first_name), get_sheet(second_name)]
    names = []  # união pela ordem de aparecimento
    for source in sources:
        for col in source["table"].columns:
            if resolve_column(col, source=names) is None:
                names.append(col)
    final_types, to_text = {}, []
    for col in names:
        found = {sh["types"].get(c, str) for sh in sources for c in sh["table"].columns if c.lower() == col.lower()}
        if len(found) == 1:
            final_types[col] = found.pop()
        elif found == {int, float}:
            final_types[col] = float
        else:
            final_types[col] = str
            to_text.append(col)
    frames = []
    for source in sources:
        data = source["table"]
        part = pd.DataFrame(index=data.index)
        for col in names:
            actual = resolve_column(col, source=data.columns)
            if actual is None:
                part[col] = pd.Series(np.nan, index=data.index, dtype="object")
            elif final_types[col] is str and source["types"].get(actual, str) is not str:
                part[col] = convert_series(data[actual], source["types"].get(actual, str), str)[0]
            else:
                part[col] = data[actual]
        frames.append(part)
    result = typed_frame(pd.concat(frames, ignore_index=True), final_types) if names else pd.DataFrame()
    return result, final_types, to_text


def concat_sheets_ui():
    if len(sheet_names()) < 2:
        print(f"{UI_INDENT_1}São precisas pelo menos duas folhas (adicione ou importe outra).")
        pause()
        return
    header("EMPILHAR FOLHAS (CONCATENAR)")
    print(f"{UI_INDENT_1}As linhas da outra folha ficam por baixo das de '{active_name}'.\n")
    other = ask_sheet("Outra folha", exclude=active_name)
    if other is None:
        return
    result, types, to_text = concat_tables(active_name, other)
    if to_text:
        print(f"{UI_INDENT_1}Aviso: tipos diferentes em {', '.join(to_text)}; essas colunas passam a texto.")
        if not confirm("Continuar?"):
            return
    name = unique_sheet_name(f"{active_name}+{other}")
    add_sheet(name, result, types)
    commit()
    print(f"\n{UI_INDENT_1}Criada a folha '{name}' com {len(result)} linhas (as originais não mudam).")
    pause()


def menu_workbook():
    while True:
        header("FOLHAS (SEPARADORES)")
        for i, name in enumerate(sheet_names(), start=1):
            data = get_sheet(name)["table"]
            print(f"{UI_INDENT_2}{'*' if name == active_name else ' '} [{i}] {name}  ({len(data)} linhas × {len(data.columns)} colunas)")
        print()
        print(f"{UI_INDENT_1}[M] Mudar de folha   [A] Adicionar folha   [R] Renomear   [E] Eliminar   [D] Duplicar")
        print(f"{UI_INDENT_1}[J] Juntar tabelas (join)   [C] Empilhar folhas (concatenar)   [0] Voltar")
        print("-" * WIDTH)
        choice = ask_option({"M", "A", "R", "E", "D", "J", "C", "0"})
        if choice == "0":
            break
        {"M": switch_sheet_ui, "A": add_blank_sheet, "R": rename_sheet_ui, "E": delete_sheet_ui,
         "D": duplicate_sheet_ui, "J": merge_sheets_ui, "C": concat_sheets_ui}[choice]()


# Tabela dinâmica (pivot)
_PIVOT_FUNCTIONS = {"1": ("sum", "soma"), "2": ("mean", "média"), "3": ("count", "contagem"), "4": ("min", "mínimo"), "5": ("max", "máximo")}


def _unique_labels(labels):
    """Nomes de coluna como texto, sem repetições (acrescenta _2, _3... ao que se repetir)."""
    seen, result = set(), []
    for label in map(str, labels):
        candidate, n = label, 2
        while candidate.lower() in seen:
            candidate = f"{label}_{n}"
            n += 1
        seen.add(candidate.lower())
        result.append(candidate)
    return result


def pivot_data(index_col, columns_col, value_col, function, totals=False):
    """Tabela dinâmica da folha activa: linhas = `index_col`, colunas = valores de `columns_col` (opcional), células = `function` de `value_col`.

    Devolve (tabela, tipos). Levanta ValueError se as colunas forem repetidas ou `function` for desconhecida.
    """
    if function not in {f for f, _ in _PIVOT_FUNCTIONS.values()}:
        raise ValueError(f"função desconhecida: {function}")
    if len({c for c in (index_col, columns_col, value_col) if c is not None}) != len([c for c in (index_col, columns_col, value_col) if c is not None]):
        raise ValueError("as colunas das linhas, das colunas e dos valores têm de ser diferentes")
    result = pd.pivot_table(table, index=index_col, columns=columns_col, values=value_col, aggfunc=function, margins=totals, margins_name="Total", observed=True)
    if isinstance(result, pd.Series):
        result = result.to_frame(value_col)
    result = result.reset_index()
    result.columns = _unique_labels(result.columns)
    return result, {c: infer_type(result[c]) for c in result.columns}


def pivot_table_ui():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    header("TABELA DINÂMICA")
    print()
    index_col = ask_column("Coluna para as linhas")
    if index_col is None:
        return
    columns_col = None
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    value = input(f"{UI_INDENT_1}Coluna para as colunas (Enter = nenhuma): ").strip()
    if value.lower() == "sair":
        return
    if value:
        columns_col = resolve_column(value)
        if columns_col is None:
            print(f"{UI_INDENT_1}Coluna não existe.")
            pause()
            return
    value_col = ask_numeric_column("Coluna numérica dos valores")
    if value_col is None:
        return
    print(f"{UI_INDENT_1}Função: " + "  ".join(f"[{k}] {label.capitalize()}" for k, (_, label) in _PIVOT_FUNCTIONS.items()))
    function = _PIVOT_FUNCTIONS.get(input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() or "1")
    if function is None:
        print(f"{UI_INDENT_1}Opção inválida.")
        pause()
        return
    totals = confirm("Incluir totais?")
    try:
        result, types = pivot_data(index_col, columns_col, value_col, function[0], totals)
    except ValueError as e:
        print(f"{UI_INDENT_1}{e}.")
        pause()
        return
    print()
    draw_data_table(result.head(DISPLAY_ROW_LIMIT))
    if confirm("Guardar como nova folha?"):
        name = unique_sheet_name("Dinamica")
        add_sheet(name, result, types)
        commit()
        print(f"\n{UI_INDENT_1}Criada a folha '{name}'.")
    pause()


# Apresentação: formatos, colunas visíveis, páginas e linha de totais
def format_column():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    header("FORMATAÇÃO DE VALORES")
    print()
    col = ask_numeric_column("Coluna a formatar")
    if col is None:
        return
    current = meta["formats"].get(col)
    print(f"{UI_INDENT_1}Formato actual: {apply_number_format(1234.5, current) if current else '(nenhum)'}  (exemplo com 1234.5)")
    print(f"{UI_INDENT_1}[1] Definir formato  [2] Remover formato  [Enter] Cancelar")
    choice = input(f"{UI_INDENT_1}Opção: ").strip()
    if choice == "2":
        meta["formats"].pop(col, None)
        commit()
        print(f"\n{UI_INDENT_1}Formato de '{col}' removido.")
        pause()
        return
    if choice != "1":
        return

    default_decimals = 0 if type_of(col) is int else 2
    decimals = ask_int_in_range(f"{UI_INDENT_1}Casas decimais 0-{FORMAT_DECIMALS_MAX} (Enter = {default_decimals}): ", 0, FORMAT_DECIMALS_MAX, default_decimals)
    if decimals is None:
        return
    thousands = confirm("Separador de milhares?")
    print(f"{UI_INDENT_1}Estilo: [1] 1,234.50 (internacional)  [2] 1 234,50 (português)")
    locale = "pt" if input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() == "2" else "intl"
    prefix = input(f"{UI_INDENT_1}Prefixo (ex.: '€ ', Enter = nenhum; os espaços contam): ")
    suffix = input(f"{UI_INDENT_1}Sufixo (ex.: ' €', '%', Enter = nenhum; os espaços contam): ")
    meta["formats"][col] = {"decimals": decimals, "thousands": thousands, "locale": locale, "prefix": prefix, "suffix": suffix}
    commit()
    print(f"\n{UI_INDENT_1}Formato aplicado a '{col}': {apply_number_format(1234.5, meta['formats'][col])}")
    pause()


def choose_visible_columns():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    header("COLUNAS VISÍVEIS")
    print()
    print(f"{UI_INDENT_1}Colunas: {', '.join(table.columns)}")
    print(f"{UI_INDENT_1}Actualmente: {', '.join(visible_columns())}")
    value = input(f"{UI_INDENT_1}Colunas a mostrar, pela ordem desejada, separadas por vírgula (Enter = todas): ").strip()
    if value.lower() == "sair":
        return
    chosen = []
    for name in [v.strip() for v in value.split(",") if v.strip()]:
        col = resolve_column(name)
        if col is None or col in chosen:
            print(f"{UI_INDENT_1}Coluna inválida ou repetida: '{name}'.")
            pause()
            return
        chosen.append(col)
    meta["visible"] = chosen
    commit()
    print(f"\n{UI_INDENT_1}{'Todas as colunas visíveis.' if not chosen else 'A mostrar: ' + ', '.join(chosen)} (os dados não mudam)")
    pause()


def view_pages():
    if table.empty:
        print(f"{UI_INDENT_1}Sem dados.")
        pause()
        return
    view = table[visible_columns()]
    size = settings["page_size"]
    page = 0
    while True:
        start, end, pages = page_bounds(len(view), size, page)
        page = start // size
        header(f"PÁGINA {page + 1}/{pages}  (linhas {start}-{end - 1} de {len(view)})", show_table=False)
        print()
        draw_data_table(view.iloc[start:end], totals_row(list(view.columns)) if page == pages - 1 else None)
        print()
        answer = input(f"{UI_INDENT_1}[N] seguinte  [P] anterior  [nº] ir para a página  [sair]: ").strip().lower()
        if answer in ("sair", "q", "0"):
            return
        if answer in ("", "n"):
            page += 1
        elif answer == "p":
            page -= 1
        elif answer.isdigit():
            page = int(answer) - 1
        else:
            print(f"{UI_INDENT_1}Opção inválida.")
            pause()


def choose_totals():
    if table.columns.empty:
        print(f"{UI_INDENT_1}Crie uma folha primeiro.")
        pause()
        return
    header("LINHA DE TOTAIS")
    print()
    print(f"{UI_INDENT_1}Mostra uma linha sob a tabela (só apresentação; não altera os dados).")
    print(f"{UI_INDENT_1}Actual: {_TOTALS_LABELS.get(meta['totals'], 'desligada')}")
    options = {str(i): f for i, f in enumerate(_TOTALS_FUNCTIONS, start=1)}
    print(f"{UI_INDENT_1}[0] Desligar  " + "  ".join(f"[{k}] {_TOTALS_LABELS[f].capitalize()}" for k, f in options.items()))
    choice = input(f"{UI_INDENT_1}Opção: ").strip()
    if choice != "0" and choice not in options:
        return
    meta["totals"] = options.get(choice)
    commit()
    print(f"\n{UI_INDENT_1}Linha de totais: {_TOTALS_LABELS.get(meta['totals'], 'desligada')}.")
    pause()


# Definições da aplicação (guardadas no estado)
def set_history_limit(steps):
    """Define quantos passos de 'desfazer' se guardam (1 a HISTORY_LIMIT_MAX) e corta o histórico se já for maior."""
    if not isinstance(steps, int) or isinstance(steps, bool) or not 1 <= steps <= HISTORY_LIMIT_MAX:
        raise ValueError(f"o número de passos tem de estar entre 1 e {HISTORY_LIMIT_MAX}")
    settings["history_limit"] = steps
    del undo_stack[:max(0, len(undo_stack) - (steps + 1))]


def menu_settings():
    while True:
        header("DEFINIÇÕES")
        print(f"{UI_INDENT_1}[1] Passos de desfazer/refazer: {settings['history_limit']} (1-{HISTORY_LIMIT_MAX}; mais passos usam mais memória)")
        print(f"{UI_INDENT_1}[2] Linhas por página em 'Ver por páginas': {settings['page_size']} (5-200)")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)
        choice = ask_option({"1", "2", "0"})
        if choice == "0":
            break
        if choice == "1":
            steps = ask_int_in_range(f"{UI_INDENT_1}Novo valor (Enter = manter): ", 1, HISTORY_LIMIT_MAX)
            if steps is not None:
                set_history_limit(steps)
        else:
            size = ask_int_in_range(f"{UI_INDENT_1}Novo valor (Enter = manter): ", 5, 200)
            if size is not None:
                settings["page_size"] = size
        save_state()


# Menus secundários: folha, linhas, colunas, análise básica e avançada
def menu_sheet():
    while True:
        header("GESTÃO DE FOLHA")
        print(f"{UI_INDENT_1}[1] Nova folha (definir colunas)")
        print(f"{UI_INDENT_1}[2] Importar Excel")
        print(f"{UI_INDENT_1}[3] Exportar Excel")
        print(f"{UI_INDENT_1}[4] Exportar Word (tabela + estatísticas)")
        print(f"{UI_INDENT_1}[5] Exportar PDF  (tabela + estatísticas)")
        print(f"{UI_INDENT_1}[6] Folha dummy (dados de exemplo)")
        print(f"{UI_INDENT_1}[7] Limpar folha (reset)")
        print(f"{UI_INDENT_1}[8] Importar CSV")
        print(f"{UI_INDENT_1}[9] Exportar CSV")
        print(f"{UI_INDENT_1}[A] Importar JSON          [B] Colar dados (CSV / Excel)")
        print(f"{UI_INDENT_1}[C] Exportar JSON          [D] Exportar Markdown     [E] Exportar HTML")
        print(f"{UI_INDENT_1}[P] Guardar projecto       [O] Abrir projecto (todas as folhas)")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "P", "O", "0"})
        if choice == "0":
            break
        elif choice == "1":
            new_sheet()
        elif choice == "2":
            import_excel()
        elif choice == "3":
            export_excel()
        elif choice == "4":
            export_word()
        elif choice == "5":
            export_pdf()
        elif choice == "6":
            create_dummy_sheet()
        elif choice == "7":
            clear_sheet()
        elif choice == "8":
            import_csv()
        elif choice == "9":
            export_csv()
        else:
            {"A": import_json, "B": import_pasted, "C": export_json, "D": export_markdown, "E": export_html,
             "P": save_project, "O": open_project}[choice]()


def menu_rows():
    while True:
        header("LINHAS")
        print(f"{UI_INDENT_1}[1] Adicionar linha(s)")
        print(f"{UI_INDENT_1}[2] Editar célula")
        print(f"{UI_INDENT_1}[3] Remover linha")
        print(f"{UI_INDENT_1}[4] Procurar e substituir")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "0"})
        if choice == "0":
            break
        elif choice == "1":
            add_row()
        elif choice == "2":
            edit_cell()
        elif choice == "3":
            remove_row()
        elif choice == "4":
            find_replace()


def menu_columns():
    while True:
        header("COLUNAS")
        print(f"{UI_INDENT_1}[1] Adicionar coluna")
        print(f"{UI_INDENT_1}[2] Remover coluna")
        print(f"{UI_INDENT_1}[3] Renomear coluna")
        print(f"{UI_INDENT_1}[4] Formatação de valores (casas decimais, milhares, €)")
        print(f"{UI_INDENT_1}[5] Colunas visíveis (escolher e ordenar)")
        print(f"{UI_INDENT_1}[6] Coluna calculada (fórmula)")
        print(f"{UI_INDENT_1}[7] Converter fórmula em valores")
        print(f"{UI_INDENT_1}[8] Converter tipo da coluna")
        print(f"{UI_INDENT_1}[9] Regras de validação")
        print(f"{UI_INDENT_1}[A] Normalizar coluna (min-máx, z-score)")
        print(f"{UI_INDENT_1}[B] Faixas / intervalos (tipo 'faixas etárias')")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "0"})
        if choice == "0":
            break
        elif choice == "1":
            add_column()
        elif choice == "2":
            remove_column()
        elif choice == "3":
            rename_column()
        elif choice == "4":
            format_column()
        elif choice == "5":
            choose_visible_columns()
        elif choice == "6":
            add_formula_column()
        elif choice == "7":
            freeze_formula()
        elif choice == "8":
            convert_column_type()
        elif choice == "9":
            menu_rules()
        elif choice == "A":
            normalize_column()
        elif choice == "B":
            bin_column()


def menu_basic_analysis():
    while True:
        header("ANÁLISE BÁSICA")
        print(f"{UI_INDENT_1}[1] Ordenar (uma ou múltiplas colunas)")
        print(f"{UI_INDENT_1}[2] Filtrar / Pesquisar")
        print(f"{UI_INDENT_1}[3] Ver Início / Fim")
        print(f"{UI_INDENT_1}[4] Amostra aleatória")
        print(f"{UI_INDENT_1}[5] Ver por páginas")
        print(f"{UI_INDENT_1}[6] Linha de totais")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "0"})
        if choice == "0":
            break
        elif choice == "1":
            sort_rows()
        elif choice == "2":
            filter_rows()
        elif choice == "3":
            view_head_tail()
        elif choice == "4":
            random_sample()
        elif choice == "5":
            view_pages()
        elif choice == "6":
            choose_totals()


def menu_advanced():
    while True:
        header("ANÁLISE AVANÇADA")
        print(f"{UI_INDENT_1}[1] Info / Resumo do DataFrame")
        print(f"{UI_INDENT_1}[2] Agrupar (groupby + agregação)")
        print(f"{UI_INDENT_1}[3] Duplicados (detectar e remover)")
        print(f"{UI_INDENT_1}[4] Dados em falta (contar / preencher / remover)")
        print(f"{UI_INDENT_1}[5] Tabela dinâmica (pivot)")
        print(f"{UI_INDENT_1}[6] Valores anómalos (outliers)")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "0"})
        if choice == "0":
            break
        elif choice == "1":
            info_summary()
        elif choice == "2":
            group_by()
        elif choice == "3":
            duplicates()
        elif choice == "4":
            missing_data()
        elif choice == "5":
            pivot_table_ui()
        elif choice == "6":
            outliers()


# Ecrã "Sobre" com informação da aplicação
def about():
    header("SOBRE")
    print(f"{UI_INDENT_1}{APP_NAME}  —  Folha de Cálculo em Linha de Comando")
    print()
    print(f"{UI_INDENT_1}Versão   : {APP_VERSION}")
    print(f"{UI_INDENT_1}Autor    : {APP_AUTHOR}")
    print(f"{UI_INDENT_1}Licença  : MIT License")
    print()
    print(f"{UI_INDENT_1}Funcionalidades:")
    print(f"{UI_INDENT_2}- Várias folhas; juntar, empilhar e tabela dinâmica")
    print(f"{UI_INDENT_2}- Tipos: texto, inteiro, decimal, data; conversão e regras de validação")
    print(f"{UI_INDENT_2}- Colunas calculadas por fórmula, recalculadas automaticamente")
    print(f"{UI_INDENT_2}- Estatísticas, correlação, outliers e análise avançada")
    print(f"{UI_INDENT_2}- Gráficos no ecrã ou gravados em PNG")
    print(f"{UI_INDENT_2}- Importar: Excel, CSV, JSON, dados colados")
    print(f"{UI_INDENT_2}- Exportar: Excel, CSV, JSON, Markdown, HTML, Word, PDF")
    print(f"{UI_INDENT_2}- Projectos, Undo / Redo e gravação automática de estado")
    print(f"{UI_INDENT_2}- Linha de comandos: --importar / --exportar (ver --help)")
    print()
    print(f"{UI_INDENT_1}Requisitos: Python 3.12+ · pandas · numpy · matplotlib · openpyxl · python-docx · fpdf2")
    print("-" * WIDTH)
    pause()


# Ponto de entrada: menu principal e ciclo da aplicação
def configure_log():
    """Regista erros em .termifolha.log (rotativo). Chamado só em main(): importar o módulo não cria ficheiros."""
    from logging.handlers import RotatingFileHandler
    try:
        handler = RotatingFileHandler(LOG_FILE, maxBytes=256 * 1024, backupCount=1, encoding="utf-8")
    except OSError:
        return  # sem permissões de escrita: a aplicação funciona na mesma, sem registo
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(handler)
    log.setLevel(logging.INFO)


def run_interactive():
    load_state()

    if not undo_stack:
        undo_stack.append(snapshot())

    while True:
        header("MENU PRINCIPAL")
        print(f"{UI_INDENT_1}[1] Folha            (nova, importar, exportar, dummy, limpar)")
        print(f"{UI_INDENT_1}[2] Linhas           (adicionar, editar, remover)")
        print(f"{UI_INDENT_1}[3] Colunas          (adicionar, remover, renomear)")
        print(f"{UI_INDENT_1}[4] Análise Básica   (ordenar, filtrar, início/fim, amostra)")
        print(f"{UI_INDENT_1}[5] Estatísticas     (8 métodos pandas)")
        print(f"{UI_INDENT_1}[6] Análise Avançada (info, groupby, duplicados, missing)")
        print(f"{UI_INDENT_1}[7] Gráficos         (matplotlib + numpy)")
        print(f"{UI_INDENT_1}[8] Folhas           (separadores, juntar, empilhar)")
        print(f"{UI_INDENT_1}[9] Definições       (passos de desfazer, linhas por página)")
        print(f"{UI_INDENT_1}[S] Sobre")
        print(f"{UI_INDENT_1}[U] Desfazer  [{len(undo_stack) - 1} passo(s)]")
        print(f"{UI_INDENT_1}[R] Refazer   [{len(redo_stack)} passo(s)]")
        print(f"{UI_INDENT_1}[Q] Sair")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "S", "U", "R", "Q"})

        if choice == "1":
            menu_sheet()
        elif choice == "2":
            menu_rows()
        elif choice == "3":
            menu_columns()
        elif choice == "4":
            menu_basic_analysis()
        elif choice == "5":
            statistics_menu()
        elif choice == "6":
            menu_advanced()
        elif choice == "7":
            menu_charts()
        elif choice == "8":
            menu_workbook()
        elif choice == "9":
            menu_settings()
        elif choice == "S":
            about()
        elif choice == "U":
            undo_last()
        elif choice == "R":
            redo_last()
        elif choice == "Q":
            clear_screen()
            print("Obrigado por usar o TermiFolha. Até à próxima!")
            break


_EXPORTERS = {
    ".xlsx": lambda path, args: write_excel(path, args.todas_as_folhas),
    ".csv": lambda path, args: write_csv(path, "pt" if args.csv_pt else "intl"),
    ".json": lambda path, args: write_json(path),
    ".md": lambda path, args: write_markdown(path),
    ".html": lambda path, args: write_html(path),
    ".docx": lambda path, args: write_word(path),
    ".pdf": lambda path, args: write_pdf(path),
}


def build_parser():
    parser = argparse.ArgumentParser(
        prog="TermiFolha",
        description="TermiFolha — folha de cálculo em linha de comando. Sem argumentos abre o menu interactivo; "
                    "com --importar/--projecto/--exportar corre sem interacção (não lê nem grava o estado da sessão).",
        epilog="Exemplo: python TermiFolha.py --importar vendas.csv --exportar vendas.xlsx --exportar relatorio.html",
    )
    parser.add_argument("--importar", metavar="FICHEIRO", help="importa um .xlsx, .csv ou .json para a folha activa (o Excel com várias folhas importa todas)")
    parser.add_argument("--projecto", metavar="FICHEIRO", help=f"abre um ficheiro de projecto ({EXT_PROJECT}) com todas as folhas")
    parser.add_argument("--exportar", metavar="FICHEIRO", action="append", default=[],
                        help="exporta a folha activa; o formato vem da extensão (.xlsx .csv .json .md .html .docx .pdf); pode repetir-se")
    parser.add_argument("--csv-pt", action="store_true", help="CSV no formato do Excel português (';' e ',')")
    parser.add_argument("--todas-as-folhas", action="store_true", help="no .xlsx, grava todas as folhas (um separador cada)")
    parser.add_argument("--versao", action="version", version=f"{APP_NAME} {APP_VERSION}")
    return parser


def load_file_non_interactive(path):
    """Importa um ficheiro (por extensão) sem fazer perguntas. Levanta ValueError se o formato não for suportado."""
    extension = os.path.splitext(path)[1].lower()
    if extension == ".xlsx":
        frames = pd.read_excel(path, sheet_name=None)
        if len(frames) == 1:
            adopt_dataframe(next(iter(frames.values())))
        else:
            adopt_workbook(frames)
    elif extension == ".csv":
        adopt_dataframe(_read_csv(path), detect_dates=True)
    elif extension == ".json":
        adopt_dataframe(read_json_file(path), detect_dates=True)
    else:
        raise ValueError(f"formato não suportado para importar: '{extension or path}' (use .xlsx, .csv ou .json)")


def run_batch(args):
    """Modo não interactivo: abre/importa, exporta e termina. Devolve o código de saída (0 = tudo bem)."""
    global STATE_ENABLED
    STATE_ENABLED = False  # não mexe no estado da sessão interactiva
    unknown = [p for p in args.exportar if os.path.splitext(p)[1].lower() not in _EXPORTERS]
    if unknown:
        print(f"Erro: formato de exportação não suportado em {', '.join(unknown)} (use {' '.join(_EXPORTERS)}).", file=sys.stderr)
        return 2
    try:
        if args.projecto:
            sheet_list, active, _file_settings = _read_state_file(args.projecto)
            _install_workbook(sheet_list, active, dict(settings))
        if args.importar:
            load_file_non_interactive(args.importar)
        if args.exportar and table.columns.empty:
            print("Erro: não há dados para exportar (indique --importar ou --projecto).", file=sys.stderr)
            return 1
        for path in args.exportar:
            _EXPORTERS[os.path.splitext(path)[1].lower()](path, args)
            print(f"Exportado para '{path}'.")
    except FileNotFoundError as e:
        print(f"Erro: ficheiro '{e.filename}' não encontrado.", file=sys.stderr)
        return 1
    except Exception as e:
        log.exception("Falha no modo não interactivo")
        print(f"Erro: {e}", file=sys.stderr)
        return 1
    return 0


def main(argv=None):
    args = build_parser().parse_args(argv)
    configure_log()
    if args.importar or args.projecto or args.exportar:
        return run_batch(args)
    run_interactive()
    return 0


if __name__ == "__main__":
    sys.exit(main())
