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
# Versão: 1.1
#
# Descrição: TermiFolha é uma aplicação de folha de cálculo em linha de comando, permite criar, editar, filtrar e analisar dados de uma folha de cálculo (DataFrame) de forma simples e eficiente.
#
# Funcionalidades:
# - Criar nova folha de cálculo com colunas personalizadas
# - Adicionar, editar e remover linhas e colunas
# - Ordenar dados por múltiplas colunas
# - Filtrar dados com critérios avançados
# - Estatísticas descritivas e análise de dados
# - Visualizações simples (histogramas, gráficos de linha, heatmaps)
# - Suporte a tipos de dados: texto, inteiro, decimal e data
# - Importar e exportar Excel e CSV; exportar Word e PDF
# - Undo/Redo e guardar automático do estado (JSON)
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
# 20-05-2026

import copy
import csv
import json
import logging
import os
import pickle

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
APP_VERSION        = "1.2"
APP_AUTHOR         = "Gabriel Gomes Alves"
_PROGRAM_DIR   = os.path.dirname(os.path.abspath(__file__))
STATE_FILE   = os.path.join(_PROGRAM_DIR, ".termifolha_estado.json")
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
    """Metadados de uma folha: regras de validação, formatos de apresentação, colunas visíveis, fórmulas e linha de totais."""
    return {"rules": {}, "formats": {}, "visible": [], "formulas": {}, "totals": None}


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


# Desenho da tabela e cabeçalho no terminal
def column_widths(data):
    widths = {}
    for col in data.columns:
        w = len(str(col))
        if not data.empty:
            for idx in data.index:
                v = data.at[idx, col]
                disp_len = len(format_value(v))
                if disp_len > w:
                    w = disp_len
        widths[col] = min(w, COLUMN_WIDTH_LIMIT)
    return widths


def draw_data_table(data):
    if data.columns.empty:
        print(f"{UI_INDENT_1}(tabela vazia — crie uma folha primeiro)")
        return

    widths = column_widths(data)
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
        return

    for idx in data.index:
        line = UI_INDENT_1 + str(idx).rjust(index_width) + " |"
        for col in data.columns:
            val = data.at[idx, col]
            shown = truncate(format_value(val), widths[col])
            line += " " + shown.ljust(widths[col]) + " |"
        print(line)


def draw_table():
    draw_data_table(table)


def header(subtitle="MENU PRINCIPAL"):
    clear_screen()
    print("=" * WIDTH)
    print(APP_TITLE.center(WIDTH))
    print("=" * WIDTH)
    if table.columns.empty:
        state_text = "folha vazia"
    else:
        state_text = f"{len(table.columns)} coluna(s)  |  {len(table)} linha(s)"
    print(f"{UI_INDENT_1}Estado: {state_text}")
    print("-" * WIDTH)
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
        print(f"{UI_INDENT_1}Deve selecionar pelo menos {min_n} coluna(s).")
        pause()
        return None
    if max_n is not None and len(chosen) > max_n:
        print(f"{UI_INDENT_1}Deve selecionar no máximo {max_n} coluna(s).")
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
    global table, col_types, meta
    if os.path.exists(STATE_FILE):
        path, reader, migrated = STATE_FILE, _read_state_file, False
    elif os.path.exists(LEGACY_STATE_FILE):
        path, reader, migrated = LEGACY_STATE_FILE, _read_legacy_pickle, True
    else:
        return

    try:
        loaded = reader(path)
    except Exception as e:
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
    global sheets
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
    """Fecha uma alteração: grava o estado e regista o ponto de desfazer."""
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
    global table, col_types
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
        adopt_dataframe(pd.read_excel(name), ask_integers=True)
        print(f"\n{UI_INDENT_1}'{name}' importado: {len(table)} linhas, {len(table.columns)} colunas.")
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


def adopt_dataframe(df, detect_dates=False, ask_integers=False):
    """Torna `df` a folha activa (importação de Excel/CSV).

    O estado global só é alterado no fim, por isso uma falha a meio não deixa
    a folha actual meio substituída.
    """
    global table, col_types, meta
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
    new_col_types = {col: infer_type(df[col]) for col in df.columns}
    table, col_types, meta = df, new_col_types, new_meta()
    apply_types()
    commit()


def _as_dates(series):
    """Converte uma coluna de texto em datas se TODOS os valores forem DD/MM/AAAA; senão None."""
    valid_values = series.dropna()
    if valid_values.empty:
        return None
    try:
        pd.to_datetime(valid_values, format="%d/%m/%Y")
    except (ValueError, TypeError):
        return None
    return pd.to_datetime(series, format="%d/%m/%Y", errors="coerce")


def _csv_separator(name, file_encoding):
    """Detecta o separador entre os mais comuns; ',' se não for possível (ex.: uma só coluna)."""
    with open(name, encoding=file_encoding, newline="") as f:
        random_sample = f.read(8192)
    try:
        return csv.Sniffer().sniff(random_sample, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


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
        for col in raw.columns:
            raw[col] = _numeric_or_text(raw[col])
        return raw
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
    sep, decimal = (";", ",") if input(f"{UI_INDENT_1}Opção (Enter = 1): ").strip() == "2" else (",", ".")

    try:
        # utf-8-sig: o Excel reconhece os acentos; datas em DD/MM/AAAA, como na aplicação.
        table.to_csv(name, index=False, sep=sep, decimal=decimal, encoding="utf-8-sig", date_format="%d/%m/%Y")
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_csv")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


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

        doc.save(name)
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

        pdf.output(name)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_pdf")
        print(f"\n{UI_INDENT_1}Erro: {e}")
    pause()


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
        table.to_excel(name, index=False)
        print(f"\n{UI_INDENT_1}Exportado para '{name}'.")
    except Exception as e:
        log.exception("Falha em export_excel")
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
            col_type = type_of(col)
            type_label = TYPE_NAME.get(col_type, "texto")
            while True:
                val = input(f"{UI_INDENT_1}{col} ({type_label}): ").strip()
                if val.lower() == "sair":
                    cancelled = True
                    break
                try:
                    new_row[col] = convert_value(val, col_type)
                    break
                except Exception:
                    print(f"{UI_INDENT_1}Valor inválido para {type_label}.")
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
    global table
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

    col_type = type_of(col)
    current = table.at[idx, col]
    print(f"{UI_INDENT_1}Valor atual: {format_value(current)}")

    while True:
        loaded = input(f"{UI_INDENT_1}Novo valor: ").strip()
        if loaded.lower() == "sair":
            print(f"\n{UI_INDENT_1}Operação cancelada.")
            pause()
            return
        try:
            table.at[idx, col] = convert_value(loaded, col_type)
            break
        except Exception:
            print(f"{UI_INDENT_1}Valor inválido.")

    commit()
    print(f"\n{UI_INDENT_1}Célula atualizada.")
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
    global table, col_types
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
            except Exception:
                print(f"{UI_INDENT_1}Valor inválido para {type_label}.")

    table[name] = values
    apply_types()
    commit()
    print(f"\n{UI_INDENT_1}Coluna '{name}' adicionada.")
    pause()


def remove_column():
    global table, col_types
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
    global table, col_types
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
    except Exception:
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
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1","2","3","4","5","6","7","8","0"})

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
            except Exception as e:
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
    plt.show()


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
    plt.show()


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
    plt.show()


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
    plt.show()


def plot_correlation_heatmap():
    cols = ask_numeric_columns(min_n=2, message="Colunas para correlação")
    if cols is None:
        return
    data_rows = table[cols].dropna()
    if len(data_rows) < 2:
        print(f"{UI_INDENT_1}Necessário pelo menos 2 linhas com dados válidos nas colunas selecionadas.")
        pause()
        return
    matrix = data_rows.corr().to_numpy()
    if np.isnan(matrix).all():
        print(f"{UI_INDENT_1}Correlação indefinida para as colunas selecionadas.")
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
    plt.show()


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
    plt.show()


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
    plt.show()


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
    plt.show()


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
    plt.show()


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
    plt.show()


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
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "0"})
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
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"})
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


def menu_rows():
    while True:
        header("LINHAS")
        print(f"{UI_INDENT_1}[1] Adicionar linha(s)")
        print(f"{UI_INDENT_1}[2] Editar célula")
        print(f"{UI_INDENT_1}[3] Remover linha")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "0"})
        if choice == "0":
            break
        elif choice == "1":
            add_row()
        elif choice == "2":
            edit_cell()
        elif choice == "3":
            remove_row()


def menu_columns():
    while True:
        header("COLUNAS")
        print(f"{UI_INDENT_1}[1] Adicionar coluna")
        print(f"{UI_INDENT_1}[2] Remover coluna")
        print(f"{UI_INDENT_1}[3] Renomear coluna")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "0"})
        if choice == "0":
            break
        elif choice == "1":
            add_column()
        elif choice == "2":
            remove_column()
        elif choice == "3":
            rename_column()


def menu_basic_analysis():
    while True:
        header("ANÁLISE BÁSICA")
        print(f"{UI_INDENT_1}[1] Ordenar (uma ou múltiplas colunas)")
        print(f"{UI_INDENT_1}[2] Filtrar / Pesquisar")
        print(f"{UI_INDENT_1}[3] Ver Início / Fim")
        print(f"{UI_INDENT_1}[4] Amostra aleatória")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "0"})
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


def menu_advanced():
    while True:
        header("ANÁLISE AVANÇADA")
        print(f"{UI_INDENT_1}[1] Info / Resumo do DataFrame")
        print(f"{UI_INDENT_1}[2] Agrupar (groupby + agregação)")
        print(f"{UI_INDENT_1}[3] Duplicados (detectar e remover)")
        print(f"{UI_INDENT_1}[4] Dados em falta (contar / preencher / remover)")
        print(f"{UI_INDENT_1}[0] Voltar")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "0"})
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
    print(f"{UI_INDENT_2}- Criar, editar, filtrar e analisar dados")
    print(f"{UI_INDENT_2}- Tipos: texto, inteiro, decimal, data")
    print(f"{UI_INDENT_2}- Estatísticas e análise avançada (groupby, duplicados)")
    print(f"{UI_INDENT_2}- Gráficos: barras, linhas, circular, dispersão, área,")
    print(f"{UI_INDENT_2}  histograma, boxplot, heatmap, tendência polinomial")
    print(f"{UI_INDENT_2}- Importar / Exportar Excel e CSV; exportar Word e PDF")
    print(f"{UI_INDENT_2}- Undo / Redo e gravação automática de estado")
    print()
    print(f"{UI_INDENT_1}Requisitos: Python 3.x · pandas · numpy · matplotlib · python-docx · fpdf2")
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


def main():
    global table, col_types

    configure_log()
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
        print(f"{UI_INDENT_1}[S] Sobre")
        print(f"{UI_INDENT_1}[U] Desfazer  [{len(undo_stack) - 1} passo(s)]")
        print(f"{UI_INDENT_1}[R] Refazer   [{len(redo_stack)} passo(s)]")
        print(f"{UI_INDENT_1}[Q] Sair")
        print("-" * WIDTH)

        choice = ask_option({"1", "2", "3", "4", "5", "6", "7", "S", "U", "R", "Q"})

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


if __name__ == "__main__":
    main()
