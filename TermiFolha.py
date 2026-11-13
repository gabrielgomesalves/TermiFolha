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
# Versão: 1.0
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
# - Undo/Redo e guardar automático do estado
#
# Requisitos: Python 3.x, pickle, numpy, pandas, matplotlib, python-docx, fpdf2
#
# Uso: Execute o script e siga as instruções no terminal para interagir com a folha de cálculo.
#
# Nota: Este código é um exemplo de implementação e pode ser expandido com mais funcionalidades ou melhorias na interface.
#
# Copyright (c) 2026 Gabriel Gomes Alves. Todos os direitos reservados.
#
# Licença: MIT License (https://opensource.org/licenses/MIT)
#
# (ASCII art gerado por https://patorjk.com/software/taag/#p=display&f=DOS+Rebel&t=TermiFolha&x=none&v=4&h=4&w=80&we=false)
#
# 20-05-2026

import csv
import json
import logging
import os
import pickle

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from fpdf import FPDF

# Constantes da aplicação (nome, versão, ficheiros, dimensões UI)
NOME_APP          = "TermiFolha"
TITULO_APP        = "TermiFolha  —  Folha de Cálculo CLI"
VERSAO_APP        = "1.0"
AUTOR_APP         = "Gabriel Gomes Alves"
_PASTA_PROGRAMA   = os.path.dirname(os.path.abspath(__file__))
FICHEIRO_ESTADO   = os.path.join(_PASTA_PROGRAMA, ".termifolha_estado.json")
FICHEIRO_ESTADO_LEGADO = os.path.join(_PASTA_PROGRAMA, ".termifolha_estado.pkl")  # formato antigo, só lido para migrar
FICHEIRO_LOG      = os.path.join(_PASTA_PROGRAMA, ".termifolha.log")
EXT_EXCEL         = ".xlsx"
EXT_CSV           = ".csv"
EXT_WORD          = ".docx"
EXT_PDF           = ".pdf"
LARGURA           = 72
COL_LIMITE        = 15
TIPO_DATA         = "date"  # sentinel para colunas data (pandas sem dtype Python nativo) - https://pandas.pydata.org/docs/user_guide/missing_data.html

TIPOS             = {
    "1": (str,       "Texto"),
    "2": (int,       "Inteiro"),
    "3": (float,     "Decimal"),
    "4": (TIPO_DATA, "Data (DD/MM/AAAA)"),
}
NOME_TIPO         = {tipo: nome.lower() for tipo, nome in TIPOS.values()}
HIST_LIMITE       = 4
UI_INDENTACAO_1   = "  "
UI_INDENTACAO_2   = "    "

log = logging.getLogger("termifolha")  # sem handlers por omissão; main() configura o ficheiro de registo
log.addHandler(logging.NullHandler())

# Estado global: tabela de dados, tipos por coluna e histórico undo/redo
tabela = pd.DataFrame()
tipos = {}

undo = []
redo = []


# Utilitários de terminal e formatação de valores
def limpar():
    os.system("cls" if os.name == "nt" else "clear")


def pausar():
    input(f"\n{UI_INDENTACAO_1}Prima [ENTER] para continuar...")


def truncar(valor, largura):
    s = str(valor)
    return s if len(s) <= largura else s[:largura - 1] + "~"


def formatar_valor(val):
    if pd.isna(val):
        return "--"
    if hasattr(val, "strftime"):
        return val.strftime("%d/%m/%Y")
    return str(val)


def formatar_numero(val):
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
def larguras(data):
    larg = {}
    for col in data.columns:
        w = len(str(col))
        if not data.empty:
            for idx in data.index:
                v = data.at[idx, col]
                disp_len = len(formatar_valor(v))
                if disp_len > w:
                    w = disp_len
        larg[col] = min(w, COL_LIMITE)
    return larg


def desenhar_tabela_dados(data):
    if data.columns.empty:
        print(f"{UI_INDENTACAO_1}(tabela vazia — crie uma folha primeiro)")
        return

    larg = larguras(data)
    idx_w = max(3, len(str(len(data))))

    cab = UI_INDENTACAO_1 + "#".rjust(idx_w) + " |"
    sep = UI_INDENTACAO_1 + "-" * idx_w + "-+"
    for col in data.columns:
        w = larg[col]
        cab += " " + truncar(col, w).ljust(w) + " |"
        sep += "-" * w + "--+"

    print(cab)
    print(sep)

    if data.empty:
        print(UI_INDENTACAO_1 + " " * idx_w + f"{UI_INDENTACAO_1}(sem linhas)")
        return

    for idx in data.index:
        linha = UI_INDENTACAO_1 + str(idx).rjust(idx_w) + " |"
        for col in data.columns:
            val = data.at[idx, col]
            mostrado = truncar(formatar_valor(val), larg[col])
            linha += " " + mostrado.ljust(larg[col]) + " |"
        print(linha)


def desenhar_tabela():
    desenhar_tabela_dados(tabela)


def cabecalho(subtitulo="MENU PRINCIPAL"):
    limpar()
    print("=" * LARGURA)
    print(TITULO_APP.center(LARGURA))
    print("=" * LARGURA)
    if tabela.columns.empty:
        estado = "folha vazia"
    else:
        estado = f"{len(tabela.columns)} coluna(s)  |  {len(tabela)} linha(s)"
    print(f"{UI_INDENTACAO_1}Estado: {estado}")
    print("-" * LARGURA)
    print()
    desenhar_tabela()
    print()
    print("=" * LARGURA)
    print(subtitulo.center(LARGURA))
    print("-" * LARGURA)


def imprimir_tabela_pandas(obj, prefixo=UI_INDENTACAO_2):
    for linha in obj.to_string().split("\n"):
        print(f"{prefixo}{linha}")


# Entrada do utilizador: opções, tipos, conversão e prompts
def pedir_opcao(validas):
    while True:
        escolha = input(f"{UI_INDENTACAO_1}Opção: ").strip().upper()
        if escolha in validas:
            return escolha
        print(f"{UI_INDENTACAO_1}Opção inválida.")


def pedir_tipo():
    opcoes = "  ".join(f"[{k}] {nome}" for k, (_, nome) in TIPOS.items())
    while True:
        print(f"{UI_INDENTACAO_1}Tipo: {opcoes}")
        t = input(f"{UI_INDENTACAO_1}Tipo: ").strip()
        if t in TIPOS:
            return TIPOS[t][0]
        print(f"{UI_INDENTACAO_1}Tipo inválido.")


def converter(valor, tipo):
    if tipo == TIPO_DATA:
        return pd.to_datetime(valor.strip(), format="%d/%m/%Y")
    return tipo(valor.strip())


def resolver_coluna(nome_input, fonte=None):
    fonte = tabela.columns if fonte is None else fonte
    nome_lower = nome_input.lower()
    for col in fonte:
        if str(col).lower() == nome_lower:
            return col
    return None


def confirmar(mensagem):
    resp = input(f"{UI_INDENTACAO_1}{mensagem} [S/N]: ").strip().lower()
    return resp in ("s", "sim", "y", "yes")


def pedir_coluna(mensagem="Coluna"):
    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}")
    valor = input(f"{UI_INDENTACAO_1}{mensagem}: ").strip()
    if valor.lower() == "sair":
        return None
    col = resolver_coluna(valor)
    if col is None:
        print(f"{UI_INDENTACAO_1}Coluna não existe.")
        pausar()
        return None
    return col


_TODAS = object()  # sentinela: "todas as colunas"


def pedir_coluna_ou_todas(mensagem="Coluna"):
    """Como pedir_coluna, mas Enter significa todas as colunas (devolve _TODAS). None se cancelar."""
    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}")
    valor = input(f"{UI_INDENTACAO_1}{mensagem} (Enter = todas): ").strip()
    if valor.lower() == "sair":
        return None
    if not valor:
        return _TODAS
    col = resolver_coluna(valor)
    if col is None:
        print(f"{UI_INDENTACAO_1}Coluna não existe.")
        pausar()
        return None
    return col


def pedir_coluna_numerica(mensagem="Coluna numérica"):
    num_cols = tabela.select_dtypes(include="number").columns.tolist()
    if not num_cols:
        print(f"{UI_INDENTACAO_1}Sem colunas numéricas.")
        pausar()
        return None
    print(f"{UI_INDENTACAO_1}Colunas numéricas: {', '.join(num_cols)}")
    valor = input(f"{UI_INDENTACAO_1}{mensagem}: ").strip()
    if valor.lower() == "sair":
        return None
    col = resolver_coluna(valor, fonte=num_cols)
    if col is None:
        print(f"{UI_INDENTACAO_1}Coluna inválida ou não numérica.")
        pausar()
        return None
    return col


def pedir_colunas_numericas(min_n=1, max_n=None, mensagem="Colunas numéricas (separadas por vírgula)"):
    num_cols = tabela.select_dtypes(include="number").columns.tolist()
    if len(num_cols) < min_n:
        print(f"{UI_INDENTACAO_1}Necessário pelo menos {min_n} coluna(s) numérica(s).")
        pausar()
        return None

    validas = [c for c in num_cols if tabela[c].notna().any()]
    if len(validas) < min_n:
        print(f"{UI_INDENTACAO_1}Necessário pelo menos {min_n} coluna(s) numérica(s) com dados.")
        pausar()
        return None

    print(f"{UI_INDENTACAO_1}Colunas numéricas disponíveis: {', '.join(validas)}")
    if max_n != 1:
        print(f"{UI_INDENTACAO_1}(pode indicar várias colunas separadas por vírgula)")
    valor = input(f"{UI_INDENTACAO_1}{mensagem}: ").strip()
    if valor.lower() == "sair":
        return None

    pedidos = [v.strip() for v in valor.split(",") if v.strip()]
    if not pedidos:
        print(f"{UI_INDENTACAO_1}Deve indicar pelo menos uma coluna.")
        pausar()
        return None

    escolhidas = []
    for nome in pedidos:
        col = resolver_coluna(nome, fonte=validas)
        if col is None:
            print(f"{UI_INDENTACAO_1}Coluna inválida ou sem dados numéricos: '{nome}'.")
            pausar()
            return None
        if col in escolhidas:
            print(f"{UI_INDENTACAO_1}Coluna repetida: '{col}'.")
            pausar()
            return None
        escolhidas.append(col)

    if len(escolhidas) < min_n:
        print(f"{UI_INDENTACAO_1}Deve selecionar pelo menos {min_n} coluna(s).")
        pausar()
        return None
    if max_n is not None and len(escolhidas) > max_n:
        print(f"{UI_INDENTACAO_1}Deve selecionar no máximo {max_n} coluna(s).")
        pausar()
        return None
    return escolhidas


def pedir_inteiro_intervalo(mensagem, minimo, maximo, default=None):
    valor = input(mensagem).strip()
    if not valor:
        return default
    try:
        numero = int(valor)
    except ValueError:
        print(f"{UI_INDENTACAO_1}Valor inválido (inteiro esperado).")
        pausar()
        return None
    if numero < minimo or numero > maximo:
        print(f"{UI_INDENTACAO_1}Valor fora do intervalo [{minimo}, {maximo}].")
        pausar()
        return None
    return numero


def tipo_de(col):
    return tipos.get(col, str)


_DTYPE_INFER = [
    (pd.api.types.is_integer_dtype,        int),
    (pd.api.types.is_float_dtype,          float),
    (pd.api.types.is_datetime64_any_dtype, TIPO_DATA),
]


def inferir_tipo(serie):
    for check, tipo in _DTYPE_INFER:
        if check(serie):
            return tipo
    return str


# Aplicação de tipos, persistência em ficheiro e histórico undo/redo
def aplicar_tipos():
    global tabela
    for col in tabela.columns:
        tipo = tipo_de(col)
        try:
            if tipo is int:
                tabela[col] = pd.to_numeric(tabela[col], errors="coerce").astype("Int64")
            elif tipo is float:
                tabela[col] = pd.to_numeric(tabela[col], errors="coerce")
            elif tipo == TIPO_DATA:
                tabela[col] = pd.to_datetime(tabela[col], errors="coerce")
        except Exception:
            pass


# Persistência: o estado é gravado em JSON (nunca em pickle, que executa código ao carregar).
_NOME_TIPO_JSON = {str: "str", int: "int", float: "float", TIPO_DATA: "date"}
_TIPO_DE_NOME_JSON = {nome: tipo for tipo, nome in _NOME_TIPO_JSON.items()}
_VERSAO_ESTADO = 1


def _valor_json(v):
    if pd.isna(v):
        return None
    if hasattr(v, "isoformat"):
        return v.isoformat()
    if hasattr(v, "item"):
        return v.item()
    return v


def _estado_para_dict():
    colunas = [str(c) for c in tabela.columns]
    linhas = [[_valor_json(v) for v in registo] for registo in tabela.itertuples(index=False, name=None)]
    return {
        "versao": _VERSAO_ESTADO,
        "colunas": colunas,
        "tipos": {str(c): _NOME_TIPO_JSON.get(t, "str") for c, t in tipos.items()},
        "dados": linhas,
    }


def _estado_de_dict(d):
    """Valida e reconstrói (tabela, tipos). Levanta ValueError se o conteúdo não for um estado válido."""
    if not isinstance(d, dict) or d.get("versao") != _VERSAO_ESTADO:
        raise ValueError("versão de estado desconhecida")
    colunas, nomes_tipos, dados = d.get("colunas"), d.get("tipos"), d.get("dados")
    if not (isinstance(colunas, list) and all(isinstance(c, str) for c in colunas)):
        raise ValueError("lista de colunas inválida")
    if not (isinstance(nomes_tipos, dict) and isinstance(dados, list)):
        raise ValueError("estrutura inválida")
    if any(not isinstance(r, list) or len(r) != len(colunas) for r in dados):
        raise ValueError("linha com número de valores diferente do número de colunas")
    novos_tipos = {c: _TIPO_DE_NOME_JSON.get(nomes_tipos.get(c), str) for c in colunas}
    df = pd.DataFrame(dados, columns=colunas) if dados else pd.DataFrame(columns=colunas)
    return df, novos_tipos


def guardar_estado():
    """Grava o estado de forma atómica. Devolve True se gravou, False se falhou (e avisa o utilizador)."""
    temporario = FICHEIRO_ESTADO + ".tmp"
    try:
        with open(temporario, "w", encoding="utf-8") as f:
            json.dump(_estado_para_dict(), f, ensure_ascii=False)
        os.replace(temporario, FICHEIRO_ESTADO)
        return True
    except Exception as e:
        log.exception("Falha ao guardar o estado")
        print(f"{UI_INDENTACAO_1}Aviso: não foi possível guardar o estado ({e}).")
        print(f"{UI_INDENTACAO_1}As alterações mantêm-se em memória; exporte para Excel para não as perder.")
        try:
            os.remove(temporario)
        except OSError:
            pass
        return False


# Tipos que um estado antigo (pickle) pode legitimamente referenciar. Tudo o resto é recusado.
_PICKLE_BUILTINS = {"str", "int", "float", "bool", "dict", "list", "tuple", "set", "frozenset", "slice", "object", "complex", "bytes", "range"}
_PICKLE_NUMPY = {"_reconstruct", "ndarray", "dtype", "_frombuffer", "scalar"}
_PICKLE_PANDAS_TOPO = {  # nunca "eval", "read_pickle", ...
    "DataFrame", "Series", "Index", "RangeIndex", "MultiIndex", "DatetimeIndex",
    "StringDtype", "Int64Dtype", "Float64Dtype", "BooleanDtype", "DatetimeTZDtype", "CategoricalDtype",
}


class _UnpicklerSeguro(pickle.Unpickler):
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


def _preservar_ficheiro_invalido(caminho, erro):
    copia = caminho + ".corrompido"
    log.error("Estado inválido em %s: %s", caminho, erro)
    print(f"{UI_INDENTACAO_1}Aviso: o estado guardado está corrompido ({erro}).")
    try:
        os.replace(caminho, copia)
        print(f"{UI_INDENTACAO_1}O ficheiro foi preservado como '{copia}'. A começar com uma folha vazia.")
    except OSError as erro_copia:
        print(f"{UI_INDENTACAO_1}Não foi possível preservar o ficheiro ({erro_copia}).")


def carregar_estado():
    global tabela, tipos
    if os.path.exists(FICHEIRO_ESTADO):
        try:
            with open(FICHEIRO_ESTADO, encoding="utf-8") as f:
                novo = _estado_de_dict(json.load(f))
        except Exception as e:
            _preservar_ficheiro_invalido(FICHEIRO_ESTADO, e)
            tabela, tipos = pd.DataFrame(), {}
            pausar()
            return
        migrado = False
    elif os.path.exists(FICHEIRO_ESTADO_LEGADO):
        try:
            with open(FICHEIRO_ESTADO_LEGADO, "rb") as f:
                antigo_tabela, antigo_tipos = _UnpicklerSeguro(f).load()
            if not isinstance(antigo_tabela, pd.DataFrame) or not isinstance(antigo_tipos, dict):
                raise ValueError("conteúdo inesperado")
            antigo_tabela = antigo_tabela.copy()
            antigo_tabela.columns = [str(c) for c in antigo_tabela.columns]
            novo = (antigo_tabela, {str(c): t for c, t in antigo_tipos.items()})
        except Exception as e:
            _preservar_ficheiro_invalido(FICHEIRO_ESTADO_LEGADO, e)
            tabela, tipos = pd.DataFrame(), {}
            pausar()
            return
        migrado = True
    else:
        return

    tabela, tipos = novo
    aplicar_tipos()
    if migrado and guardar_estado():
        try:
            os.replace(FICHEIRO_ESTADO_LEGADO, FICHEIRO_ESTADO_LEGADO + ".migrado")
            print(f"{UI_INDENTACAO_1}Estado antigo migrado para o novo formato (cópia em '{FICHEIRO_ESTADO_LEGADO}.migrado').")
            pausar()
        except OSError as e:
            log.warning("Não foi possível renomear o estado antigo: %s", e)
    undo.clear()
    redo.clear()
    undo.append((tabela.copy(), tipos.copy()))


def registar_ponto():
    redo.clear()
    undo.append((tabela.copy(), tipos.copy()))
    if len(undo) > HIST_LIMITE:
        undo.pop(0)


def desfazer():
    global tabela, tipos
    if len(undo) <= 1:
        print(f"{UI_INDENTACAO_1}Nada para desfazer.")
        pausar()
        return
    redo.append(undo.pop())
    tabela, tipos = undo[-1][0].copy(), undo[-1][1].copy()
    guardar_estado()
    print(f"{UI_INDENTACAO_1}Desfeito. ({len(undo) - 1} passo(s) atrás)")
    pausar()


def refazer():
    global tabela, tipos
    if not redo:
        print(f"{UI_INDENTACAO_1}Nada para refazer.")
        pausar()
        return
    state = redo.pop()
    undo.append(state)
    tabela, tipos = state[0].copy(), state[1].copy()
    guardar_estado()
    print(f"{UI_INDENTACAO_1}Refeito. ({len(redo)} passo(s) para refazer)")
    pausar()


# Gestão da folha: criar, importar, exportar, dummy, limpar
def nova_folha():
    global tabela, tipos
    cabecalho("NOVA FOLHA")
    print(f"{UI_INDENTACAO_1}Defina as colunas (nome vazio ou 'sair' para terminar | 'sair' sem colunas para cancelar)\n")

    colunas = []
    novos_tipos = {}

    while True:
        nome = input(f"{UI_INDENTACAO_1}Nome da coluna {len(colunas) + 1}: ").strip()
        if nome.lower() == "sair":
            if not colunas:
                print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
                pausar()
                return
            break
        if not nome:
            if not colunas:
                print(f"{UI_INDENTACAO_1}Precisa de pelo menos uma coluna.")
                continue
            break
        if any(c.lower() == nome.lower() for c in colunas):
            print(f"{UI_INDENTACAO_1}Coluna '{nome}' já existe. Por favor escolha um nome diferente.")
            continue
        novos_tipos[nome] = pedir_tipo()
        colunas.append(nome)
        print(f"{UI_INDENTACAO_1}+ coluna '{nome}' do tipo '{NOME_TIPO.get(novos_tipos[nome], 'texto')}' adicionada.\n")

    tabela = pd.DataFrame(columns=colunas)
    tipos = novos_tipos
    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Folha criada com {len(colunas)} coluna(s).")
    pausar()


def confirmar_substituir_folha():
    """Se já existir folha, pede confirmação antes de a substituir. Devolve False se o utilizador recusar."""
    if tabela.columns.empty or confirmar("Folha actual será substituída. Continuar?"):
        return True
    print(f"{UI_INDENTACAO_1}Operação cancelada.")
    pausar()
    return False


def importar():
    global tabela, tipos
    cabecalho("IMPORTAR EXCEL")
    print()
    if not confirmar_substituir_folha():
        return
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return

    if not nome.endswith(EXT_EXCEL):
        nome += EXT_EXCEL

    try:
        adoptar_dataframe(pd.read_excel(nome), perguntar_inteiros=True)
        print(f"\n{UI_INDENTACAO_1}'{nome}' importado: {len(tabela)} linhas, {len(tabela.columns)} colunas.")
    except FileNotFoundError:
        print(f"\n{UI_INDENTACAO_1}Ficheiro '{nome}' não encontrado.")
    except Exception as e:
        log.exception("Falha em importar")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def _inteiros_com_falhas(serie):
    """True se for decimal com células vazias e todos os valores presentes forem inteiros (ex.: 25.0, 40.0, NaN)."""
    validos = serie.dropna()
    return pd.api.types.is_float_dtype(serie) and 0 < len(validos) < len(serie) and bool((validos % 1 == 0).all())


def adoptar_dataframe(df, detectar_datas=False, perguntar_inteiros=False):
    """Torna `df` a folha activa (importação de Excel/CSV).

    O estado global só é alterado no fim, por isso uma falha a meio não deixa
    a folha actual meio substituída.
    """
    global tabela, tipos
    df = df.copy()
    # Cabeçalhos numéricos (ex.: anos) são tratados como texto no resto da aplicação.
    df.columns = [str(c) for c in df.columns]
    if perguntar_inteiros:
        # O pandas converte inteiros com células vazias em decimais; só o utilizador sabe se era intencional.
        for col in df.columns:
            if _inteiros_com_falhas(df[col]) and confirmar(f"A coluna '{col}' só tem valores inteiros (e células vazias). Tratar como inteiro?"):
                df[col] = df[col].astype("Int64")
    if detectar_datas:
        for col in df.columns:
            if inferir_tipo(df[col]) is str:
                datas = _como_datas(df[col])
                if datas is not None:
                    df[col] = datas
    novos_tipos = {col: inferir_tipo(df[col]) for col in df.columns}
    tabela, tipos = df, novos_tipos
    aplicar_tipos()
    guardar_estado()
    registar_ponto()


def _como_datas(serie):
    """Converte uma coluna de texto em datas se TODOS os valores forem DD/MM/AAAA; senão None."""
    validos = serie.dropna()
    if validos.empty:
        return None
    try:
        pd.to_datetime(validos, format="%d/%m/%Y")
    except (ValueError, TypeError):
        return None
    return pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")


def _separador_csv(nome, codificacao):
    """Detecta o separador entre os mais comuns; ',' se não for possível (ex.: uma só coluna)."""
    with open(nome, encoding=codificacao, newline="") as f:
        amostra = f.read(8192)
    try:
        return csv.Sniffer().sniff(amostra, delimiters=",;\t|").delimiter
    except csv.Error:
        return ","


def _numerica_ou_texto(serie):
    """Inferência a partir do texto cru: inteiros (com falhas) -> Int64, decimais -> float64, resto -> texto.

    Códigos com zeros à esquerda (ex.: 007) ficam como texto para não perder os zeros.
    """
    serie = serie.str.strip()
    validos = serie.dropna()
    if validos.empty or validos.str.contains(r"^[+-]?0\d").any():
        return serie
    if validos.str.fullmatch(r"[+-]?\d+(,\d+)?").all() and validos.str.contains(",").any():  # vírgula decimal (Excel PT)
        serie = serie.str.replace(",", ".", regex=False)
        validos = serie.dropna()
    if validos.str.fullmatch(r"[+-]?\d+").all():
        return pd.to_numeric(serie).astype("Int64")
    convertida = pd.to_numeric(serie, errors="coerce")
    if convertida[serie.notna()].notna().all():
        return convertida
    return serie


def _ler_csv(nome):
    """Lê um CSV (UTF-8 ou Latin-1, separador detectado) e infere os tipos das colunas."""
    for codificacao in ("utf-8-sig", "latin-1"):
        try:
            sep = _separador_csv(nome, codificacao)
            cru = pd.read_csv(nome, sep=sep, encoding=codificacao, dtype=str)
        except UnicodeDecodeError:
            continue
        except pd.errors.EmptyDataError:
            raise ValueError("o ficheiro está vazio") from None
        for col in cru.columns:
            cru[col] = _numerica_ou_texto(cru[col])
        return cru
    raise ValueError("não foi possível descodificar o ficheiro (esperado UTF-8 ou Latin-1)")


def importar_csv():
    cabecalho("IMPORTAR CSV")
    print()
    if not confirmar_substituir_folha():
        return
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return

    if not nome.endswith(EXT_CSV):
        nome += EXT_CSV

    try:
        adoptar_dataframe(_ler_csv(nome), detectar_datas=True)
        print(f"\n{UI_INDENTACAO_1}'{nome}' importado: {len(tabela)} linhas, {len(tabela.columns)} colunas.")
    except FileNotFoundError:
        print(f"\n{UI_INDENTACAO_1}Ficheiro '{nome}' não encontrado.")
    except Exception as e:
        log.exception("Falha em importar_csv")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def exportar_csv():
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para exportar.")
        pausar()
        return

    cabecalho("EXPORTAR CSV")
    print()
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return
    if not nome.endswith(EXT_CSV):
        nome += EXT_CSV

    print(f"{UI_INDENTACAO_1}Formato: [1] Internacional (',' e '.')  [2] Excel português (';' e ',')")
    sep, decimal = (";", ",") if input(f"{UI_INDENTACAO_1}Opção (Enter = 1): ").strip() == "2" else (",", ".")

    try:
        # utf-8-sig: o Excel reconhece os acentos; datas em DD/MM/AAAA, como na aplicação.
        tabela.to_csv(nome, index=False, sep=sep, decimal=decimal, encoding="utf-8-sig", date_format="%d/%m/%Y")
        print(f"\n{UI_INDENTACAO_1}Exportado para '{nome}'.")
    except Exception as e:
        log.exception("Falha em exportar_csv")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def exportar_word():
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para exportar.")
        pausar()
        return

    cabecalho("EXPORTAR WORD")
    print()
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return
    if not nome.endswith(EXT_WORD):
        nome += EXT_WORD

    try:
        doc = Document()
        doc.add_heading(f"{NOME_APP} — Relatório", 0)
        doc.add_paragraph(f"Dimensões: {len(tabela)} linhas × {len(tabela.columns)} colunas")

        doc.add_heading("Colunas e tipos", level=1)
        for col in tabela.columns:
            doc.add_paragraph(f"• {col}: {NOME_TIPO.get(tipo_de(col), 'texto')}")

        if not tabela.empty:
            doc.add_heading("Dados", level=1)
            cols = list(tabela.columns)
            tbl = doc.add_table(rows=1 + len(tabela), cols=len(cols))
            tbl.style = "Table Grid"
            for j, col in enumerate(cols):
                tbl.rows[0].cells[j].text = str(col)
            for i, (_, row) in enumerate(tabela.iterrows()):
                for j, col in enumerate(cols):
                    tbl.rows[i + 1].cells[j].text = formatar_valor(row[col])

        num_cols = tabela.select_dtypes(include="number").columns.tolist()
        if num_cols:
            doc.add_heading("Estatísticas básicas (colunas numéricas)", level=1)
            for col in num_cols:
                serie = tabela[col].dropna()
                doc.add_paragraph(
                    f"{col}:  média={formatar_numero(serie.mean())}  "
                    f"mín={formatar_numero(serie.min())}  "
                    f"máx={formatar_numero(serie.max())}  "
                    f"desvio={formatar_numero(serie.std())}  "
                    f"mediana={formatar_numero(serie.median())}"
                )

        doc.save(nome)
        print(f"\n{UI_INDENTACAO_1}Exportado para '{nome}'.")
    except Exception as e:
        log.exception("Falha em exportar_word")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def _latin1(texto):
    """Fontes core do PDF só suportam Latin-1: o resto é substituído por '?'."""
    return str(texto).encode("latin-1", "replace").decode("latin-1")


def _pdf_linha(pdf, altura, texto, align=""):
    """Escreve uma linha e avança. Funciona em fpdf e fpdf2 (evita o `ln=True` deprecado)."""
    pdf.cell(0, altura, _latin1(texto), align=align)
    pdf.ln(altura)


def exportar_pdf():
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para exportar.")
        pausar()
        return

    cabecalho("EXPORTAR PDF")
    print()
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return
    if not nome.endswith(EXT_PDF):
        nome += EXT_PDF

    try:
        pdf = FPDF()
        pdf.add_page()

        # Título
        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(16, 41, 230)
        _pdf_linha(pdf, 10, f"{NOME_APP} - Relatorio", align="C")

        # Dimensões
        pdf.set_font("Helvetica", size=11)
        pdf.set_text_color(0, 0, 0)
        _pdf_linha(pdf, 8, f"Dimensoes: {len(tabela)} linhas x {len(tabela.columns)} colunas")

        # Colunas e tipos
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(14, 122, 230)
        _pdf_linha(pdf, 8, "Colunas e tipos")
        pdf.set_font("Helvetica", size=10)
        pdf.set_text_color(0, 0, 0)
        for col in tabela.columns:
            nome_tipo = NOME_TIPO.get(tipo_de(col), "texto")
            _pdf_linha(pdf, 7, f"  {col}: {nome_tipo}")

        # Dados
        if not tabela.empty:
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_text_color(14, 122, 230)
            _pdf_linha(pdf, 8, "Dados")
            pdf.set_text_color(0, 0, 0)

            cols = list(tabela.columns)
            col_w = min(35, (pdf.w - 20) / len(cols))
            pdf.set_font("Helvetica", "B", 8)
            for col in cols:
                pdf.cell(col_w, 7, _latin1(str(col)[:14]), border=1)
            pdf.ln()
            pdf.set_font("Helvetica", size=8)
            for _, row in tabela.iterrows():
                for col in cols:
                    pdf.cell(col_w, 6, _latin1(str(formatar_valor(row[col]))[:14]), border=1)
                pdf.ln()

        # Estatísticas
        num_cols = tabela.select_dtypes(include="number").columns.tolist()
        if num_cols:
            pdf.ln(3)
            pdf.set_font("Helvetica", "B", 12)
            pdf.set_text_color(14, 122, 230)
            _pdf_linha(pdf, 8, "Estatisticas basicas (colunas numericas)")
            pdf.set_font("Helvetica", size=10)
            pdf.set_text_color(0, 0, 0)
            for col in num_cols:
                serie = tabela[col].dropna()
                linha = (
                    f"{col}: media={formatar_numero(serie.mean())}  "
                    f"min={formatar_numero(serie.min())}  "
                    f"max={formatar_numero(serie.max())}  "
                    f"desvio={formatar_numero(serie.std())}  "
                    f"mediana={formatar_numero(serie.median())}"
                )
                _pdf_linha(pdf, 7, linha)

        pdf.output(nome)
        print(f"\n{UI_INDENTACAO_1}Exportado para '{nome}'.")
    except Exception as e:
        log.exception("Falha em exportar_pdf")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def exportar():
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para exportar.")
        pausar()
        return

    cabecalho("EXPORTAR EXCEL")
    print()
    nome = input(f"{UI_INDENTACAO_1}Nome do ficheiro (sem extensão): ").strip()
    if not nome:
        return

    if not nome.endswith(EXT_EXCEL):
        nome += EXT_EXCEL

    try:
        tabela.to_excel(nome, index=False)
        print(f"\n{UI_INDENTACAO_1}Exportado para '{nome}'.")
    except Exception as e:
        log.exception("Falha em exportar")
        print(f"\n{UI_INDENTACAO_1}Erro: {e}")
    pausar()


def criar_folha_dummy():
    global tabela, tipos

    cabecalho("CRIAR FOLHA DUMMY")
    print()
    if not tabela.columns.empty:
        if not confirmar("Folha actual será substituída. Continuar?"):
            print(f"\n{UI_INDENTACAO_1}  Operação cancelada.")
            pausar()
            return

    tabela = pd.DataFrame({
        "Nome": ["João", "Maria", "Pedro", "Ana", "Carlos", "Sofia", "João", "Rita"],
        "Idade": [25, 30, 35, None, 45, 28, 25, 33],
        "Salario": [1500.50, 2000.00, 1800.75, 2200.00, 3500.00, 1700.25, 1500.50, 1950.00],
        "Cidade": ["Lisboa", "Porto", "Faro", "Coimbra", "Lisboa", "Porto", "Lisboa", "Braga"],
        "Admissao": pd.to_datetime([
            "15/01/2020", "01/06/2019", "10/03/2021", "20/11/2018",
            "30/05/2017", "14/02/2022", "15/01/2020", "25/09/2019"
        ], format="%d/%m/%Y"),
    })
    tipos = {"Nome": str, "Idade": int, "Salario": float, "Cidade": str, "Admissao": TIPO_DATA}
    aplicar_tipos()
    guardar_estado()
    registar_ponto()

    print(f"\n{UI_INDENTACAO_1}Folha dummy criada: 8 linhas | 5 colunas")
    print(f"{UI_INDENTACAO_1}Inclui:")
    print(f"{UI_INDENTACAO_2}• Tipos variados: texto, inteiro, decimal, data")
    print(f"{UI_INDENTACAO_2}• 1 valor em falta na coluna 'Idade'")
    print(f"{UI_INDENTACAO_2}• 1 linha duplicada (João, 25, Lisboa)")
    print(f"{UI_INDENTACAO_2}• Cidades repetidas (ideal para 'groupby')")
    pausar()


def limpar_folha():
    global tabela, tipos
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Folha já está vazia.")
        pausar()
        return

    cabecalho("LIMPAR FOLHA")
    print()
    if not confirmar("Confirma apagar toda a folha actual (perdem-se dados)?"):
        print(f"\n{UI_INDENTACAO_1}  Operação cancelada.")
        pausar()
        return

    tabela = pd.DataFrame()
    tipos = {}
    try:
        if os.path.exists(FICHEIRO_ESTADO):
            os.remove(FICHEIRO_ESTADO)
    except Exception:
        pass
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Folha limpa.")
    pausar()


# Operações sobre linhas: adicionar, editar célula, remover
def adicionar_linha():
    global tabela
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Crie uma folha primeiro.")
        pausar()
        return

    cabecalho("ADICIONAR LINHA")
    print(f"{UI_INDENTACAO_1}('sair' em qualquer campo para terminar a introdução de linhas)\n")

    linhas_adicionadas = 0

    while True:
        nova_linha = {}
        cancelar = False

        for col in tabela.columns:
            tipo = tipo_de(col)
            nome_tipo = NOME_TIPO.get(tipo, "texto")
            while True:
                val = input(f"{UI_INDENTACAO_1}{col} ({nome_tipo}): ").strip()
                if val.lower() == "sair":
                    cancelar = True
                    break
                try:
                    nova_linha[col] = converter(val, tipo)
                    break
                except Exception:
                    print(f"{UI_INDENTACAO_1}Valor inválido para {nome_tipo}.")
            if cancelar:
                break

        if cancelar:
            break

        nova = pd.DataFrame([nova_linha])
        tabela = pd.concat([tabela, nova], ignore_index=True)
        aplicar_tipos()
        guardar_estado()
        registar_ponto()
        linhas_adicionadas += 1
        print(f"\n{UI_INDENTACAO_1}Linha adicionada. (total introduzido: {linhas_adicionadas})\n")

    if linhas_adicionadas > 0:
        print(f"\n{UI_INDENTACAO_1}{linhas_adicionadas} linha(s) adicionada(s) com sucesso.")
    else:
        print(f"\n{UI_INDENTACAO_1}Nenhuma linha adicionada.")
    pausar()


def editar_celula():
    global tabela
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para editar.")
        pausar()
        return

    cabecalho("EDITAR CÉLULA")
    print(f"{UI_INDENTACAO_1}('sair' em qualquer campo para cancelar)\n")

    idx_input = input(f"{UI_INDENTACAO_1}Índice da linha: ").strip()
    if idx_input.lower() == "sair":
        print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
        pausar()
        return
    try:
        idx = int(idx_input)
    except ValueError:
        print(f"{UI_INDENTACAO_1}Índice deve ser inteiro.")
        pausar()
        return

    if idx not in tabela.index:
        print(f"{UI_INDENTACAO_1}Índice inválido.")
        pausar()
        return

    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}")
    col_input = input(f"{UI_INDENTACAO_1}Coluna: ").strip()
    if col_input.lower() == "sair":
        print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
        pausar()
        return
    col = resolver_coluna(col_input)
    if col is None:
        print(f"{UI_INDENTACAO_1}Coluna não existe.")
        pausar()
        return

    tipo = tipo_de(col)
    atual = tabela.at[idx, col]
    print(f"{UI_INDENTACAO_1}Valor atual: {formatar_valor(atual)}")

    while True:
        novo = input(f"{UI_INDENTACAO_1}Novo valor: ").strip()
        if novo.lower() == "sair":
            print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
            pausar()
            return
        try:
            tabela.at[idx, col] = converter(novo, tipo)
            break
        except Exception:
            print(f"{UI_INDENTACAO_1}Valor inválido.")

    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Célula atualizada.")
    pausar()


def remover_linha():
    global tabela
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem linhas para remover.")
        pausar()
        return

    cabecalho("REMOVER LINHA")
    print()

    try:
        idx = int(input(f"{UI_INDENTACAO_1}Índice da linha a remover: ").strip())
    except ValueError:
        print(f"{UI_INDENTACAO_1}Índice deve ser inteiro.")
        pausar()
        return

    if idx not in tabela.index:
        print(f"{UI_INDENTACAO_1}Índice inválido.")
        pausar()
        return

    print()
    desenhar_tabela_dados(tabela.loc[[idx]])
    print()
    if not confirmar(f"Confirma remover linha {idx}?"):
        print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
        pausar()
        return

    tabela = tabela.drop(index=idx).reset_index(drop=True)
    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Linha {idx} removida.")
    pausar()


# Operações sobre colunas: adicionar, remover, renomear
def adicionar_coluna():
    global tabela, tipos
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Crie uma folha primeiro.")
        pausar()
        return

    cabecalho("ADICIONAR COLUNA")
    print(f"{UI_INDENTACAO_1}('sair' para cancelar)\n")

    nome = input(f"{UI_INDENTACAO_1}Nome da nova coluna: ").strip()
    if not nome or nome.lower() == "sair":
        return
    if resolver_coluna(nome) is not None:
        print(f"{UI_INDENTACAO_1}Coluna '{nome}' já existe. Por favor escolha um nome diferente.")
        pausar()
        return

    tipo = pedir_tipo()
    tipos[nome] = tipo
    nome_tipo = NOME_TIPO.get(tipo, "texto")

    if tabela.empty:
        tabela[nome] = pd.Series(dtype=object)
        aplicar_tipos()
        guardar_estado()
        registar_ponto()
        print(f"\n{UI_INDENTACAO_1}Coluna '{nome}' adicionada.")
        pausar()
        return

    valores = []
    for idx in tabela.index:
        while True:
            val = input(f"{UI_INDENTACAO_1}Valor para linha {idx} ({nome_tipo}) ('sair' cancela, Enter = em branco): ").strip()
            if val.lower() == "sair":
                tipos.pop(nome, None)
                print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
                pausar()
                return
            if not val:
                valores.append(None)  # em branco = valor em falta (apresentado como "--")
                break
            try:
                valores.append(converter(val, tipo))
                break
            except Exception:
                print(f"{UI_INDENTACAO_1}Valor inválido para {nome_tipo}.")

    tabela[nome] = valores
    aplicar_tipos()
    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Coluna '{nome}' adicionada.")
    pausar()


def remover_coluna():
    global tabela, tipos
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem colunas para remover.")
        pausar()
        return

    cabecalho("REMOVER COLUNA")
    print()
    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}")
    nome = resolver_coluna(input(f"{UI_INDENTACAO_1}Coluna a remover: ").strip())

    if nome is None:
        print(f"{UI_INDENTACAO_1}Coluna não existe.")
        pausar()
        return

    if not confirmar(f"Confirma remover coluna '{nome}'?"):
        print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
        pausar()
        return

    tabela = tabela.drop(columns=[nome])
    tipos.pop(nome, None)
    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}Coluna '{nome}' removida.")
    pausar()


def renomear_coluna():
    global tabela, tipos
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem colunas para renomear.")
        pausar()
        return

    cabecalho("RENOMEAR COLUNA")
    print()
    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}")
    nome = resolver_coluna(input(f"{UI_INDENTACAO_1}Coluna a renomear: ").strip())

    if nome is None:
        print(f"{UI_INDENTACAO_1}Coluna não existe.")
        pausar()
        return

    while True:
        novo_nome = input(f"{UI_INDENTACAO_1}Novo nome ('sair' para cancelar): ").strip()
        if novo_nome.lower() == "sair":
            print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
            pausar()
            return
        if not novo_nome:
            print(f"{UI_INDENTACAO_1}Nome não pode ser vazio.")
        elif resolver_coluna(novo_nome) is not None:
            print(f"{UI_INDENTACAO_1}Já existe uma coluna com esse nome.")
        else:
            break

    tabela = tabela.rename(columns={nome: novo_nome})
    if nome in tipos:
        tipos[novo_nome] = tipos.pop(nome)
    guardar_estado()
    registar_ponto()
    print(f"\n{UI_INDENTACAO_1}'{nome}' renomeada para '{novo_nome}'.")
    pausar()


# Análise básica: ordenar, filtrar (numérico/data/texto), ver início/fim, amostra
def ordenar():
    global tabela
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para ordenar.")
        pausar()
        return

    cabecalho("ORDENAR")
    print(f"{UI_INDENTACAO_1}Adicione colunas em sequência ('sair' para terminar)\n")
    print(f"{UI_INDENTACAO_1}Colunas: {', '.join(tabela.columns)}\n")

    cols_ord = []
    asc_ord = []

    while True:
        valor = input(f"{UI_INDENTACAO_1}Coluna {len(cols_ord) + 1}: ").strip()
        if valor.lower() == "sair":
            if not cols_ord:
                print(f"\n{UI_INDENTACAO_1}Operação cancelada.")
                pausar()
                return
            break
        if not valor:
            if not cols_ord:
                continue
            break
        col = resolver_coluna(valor)
        if col is None:
            print(f"{UI_INDENTACAO_1}Coluna não existe.")
            continue
        if col in cols_ord:
            print(f"{UI_INDENTACAO_1}Coluna já adicionada.")
            continue
        direccao = input(f"{UI_INDENTACAO_2}Direcção [A]scendente / [D]escendente: ").strip().upper()
        cols_ord.append(col)
        asc_ord.append(direccao != "D")

    tabela = tabela.sort_values(by=cols_ord, ascending=asc_ord, ignore_index=True)
    guardar_estado()
    registar_ponto()

    descr = ", ".join(f"{c} ({'asc' if a else 'desc'})" for c, a in zip(cols_ord, asc_ord, strict=True))
    print(f"\n{UI_INDENTACAO_1}Ordenado por: {descr}")
    pausar()


def filtro_numerico(col):
    print()
    print(f"{UI_INDENTACAO_1}Operador: [1] ==  [2] >  [3] <  [4] >=  [5] <=  [6] intervalo [min,max]")
    op = input(f"{UI_INDENTACAO_1}Opção: ").strip()
    try:
        if op == "6":
            mn = float(input(f"{UI_INDENTACAO_1}Mínimo: ").strip())
            mx = float(input(f"{UI_INDENTACAO_1}Máximo: ").strip())
            return tabela[tabela[col].between(mn, mx)]
        val = float(input(f"{UI_INDENTACAO_1}Valor: ").strip())
        if op == "1":
            return tabela[tabela[col] == val]
        elif op == "2":
            return tabela[tabela[col] > val]
        elif op == "3":
            return tabela[tabela[col] < val]
        elif op == "4":
            return tabela[tabela[col] >= val]
        elif op == "5":
            return tabela[tabela[col] <= val]
        print(f"{UI_INDENTACAO_1}Operador inválido.")
        pausar()
        return None
    except ValueError:
        print(f"{UI_INDENTACAO_1}Valor inválido.")
        pausar()
        return None


def filtro_data(col):
    print()
    print(f"{UI_INDENTACAO_1}Operador: [1] Antes de  [2] Depois de  [3] Entre duas datas")
    op = input(f"{UI_INDENTACAO_1}Opção: ").strip()
    try:
        if op == "3":
            d1 = pd.to_datetime(input(f"{UI_INDENTACAO_1}De (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
            d2 = pd.to_datetime(input(f"{UI_INDENTACAO_1}Até (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
            return tabela[(tabela[col] >= d1) & (tabela[col] <= d2)]
        d = pd.to_datetime(input(f"{UI_INDENTACAO_1}Data (DD/MM/AAAA): ").strip(), format="%d/%m/%Y")
        if op == "1":
            return tabela[tabela[col] < d]
        elif op == "2":
            return tabela[tabela[col] > d]
        print(f"{UI_INDENTACAO_1}Operador inválido.")
        pausar()
        return None
    except Exception:
        print(f"{UI_INDENTACAO_1}Data inválida.")
        pausar()
        return None


def filtro_texto(col):
    print()
    print(f"{UI_INDENTACAO_1}Operador: [1] Contém  [2] Igual a  [3] Começa com  [4] Termina com")
    op = input(f"{UI_INDENTACAO_1}Opção: ").strip()
    if op not in ("1", "2", "3", "4"):
        print(f"{UI_INDENTACAO_1}Operador inválido.")
        pausar()
        return None
    termo = input(f"{UI_INDENTACAO_1}Termo: ").strip()
    serie = tabela[col].astype(str)
    termo_lower = termo.lower()
    if op == "2":
        return tabela[serie.str.lower() == termo_lower]
    elif op == "3":
        return tabela[serie.str.lower().str.startswith(termo_lower, na=False)]
    elif op == "4":
        return tabela[serie.str.lower().str.endswith(termo_lower, na=False)]
    return tabela[serie.str.contains(termo, case=False, na=False, regex=False)]


def filtrar():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados para filtrar.")
        pausar()
        return

    cabecalho("FILTRAR / PESQUISAR")
    print()
    col = pedir_coluna("Coluna a filtrar")
    if col is None:
        return

    serie = tabela[col]
    if pd.api.types.is_numeric_dtype(serie):
        resultado = filtro_numerico(col)
    elif pd.api.types.is_datetime64_any_dtype(serie):
        resultado = filtro_data(col)
    else:
        resultado = filtro_texto(col)

    if resultado is None:
        print(f"{UI_INDENTACAO_1}Filtro inválido ou cancelado.")
        pausar()
        return

    print()
    if resultado.empty:
        print(f"{UI_INDENTACAO_1}Nenhum resultado encontrado.")
    else:
        print(f"{UI_INDENTACAO_1}{len(resultado)} resultado(s) em '{col}':\n")
        desenhar_tabela_dados(resultado)
    pausar()


def ver_inicio_fim():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    cabecalho("INÍCIO / FIM")
    print()
    print(f"{UI_INDENTACAO_1}Mostra as primeiras ou últimas N linhas da tabela.")
    print()
    n_input = input(f"{UI_INDENTACAO_1}N — nº de linhas a mostrar (Enter = 5): ").strip()
    try:
        n = int(n_input) if n_input else 5
        if n <= 0:
            raise ValueError
    except ValueError:
        n = 5

    modo = input(f"{UI_INDENTACAO_1}[I] Início  [F] Fim: ").strip().upper()
    if modo == "F":
        print(f"\n{UI_INDENTACAO_1}Últimas {n} linhas:\n")
        desenhar_tabela_dados(tabela.tail(n))
    else:
        print(f"\n{UI_INDENTACAO_1}Primeiras {n} linhas:\n")
        desenhar_tabela_dados(tabela.head(n))
    pausar()


def amostra():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    cabecalho("AMOSTRA ALEATÓRIA")
    print()
    n_input = input(f"{UI_INDENTACAO_1}Tamanho: nº de linhas (1-{len(tabela)}) ou percentagem, ex. 20% (Enter = 5): ").strip()
    try:
        if n_input.endswith("%"):
            pct = float(n_input[:-1])
            if pct <= 0 or pct > 100:
                raise ValueError
            n = max(1, round(len(tabela) * pct / 100))
        else:
            n = int(n_input) if n_input else 5
            if n <= 0 or n > len(tabela):
                raise ValueError
    except ValueError:
        n = min(5, len(tabela))

    resultado = tabela.sample(n=n)
    print(f"\n{UI_INDENTACAO_1} Amostra aleatória de {n} linha(s):\n")
    desenhar_tabela_dados(resultado)
    pausar()


# Estatísticas: etiquetas, métodos individuais e menu
ETIQUETAS_ESTAT = {
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


def _resumo_todas_numericas():
    numericas = tabela.select_dtypes(include="number")
    if numericas.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem colunas numéricas.")
        pausar()
        return
    desc = numericas.describe()
    larg = {c: max(len(str(c)), 10) for c in desc.columns}
    print()
    print(f"{UI_INDENTACAO_1}Resumo estatístico de todas as colunas numéricas:\n")
    print(UI_INDENTACAO_2 + " " * 22 + " " + " ".join(str(c).rjust(larg[c]) for c in desc.columns))
    for stat in desc.index:
        celulas = " ".join(formatar_numero(float(desc.at[stat, c])).rjust(larg[c]) for c in desc.columns)
        print(UI_INDENTACAO_2 + ETIQUETAS_ESTAT.get(stat, stat).ljust(22) + " " + celulas)
    pausar()


def estat_resumo():
    col = pedir_coluna_ou_todas()
    if col is None:
        return
    if col is _TODAS:
        _resumo_todas_numericas()
        return

    desc = tabela[col].describe()
    tipo_col = "numérica" if pd.api.types.is_numeric_dtype(tabela[col]) else (
        "data" if pd.api.types.is_datetime64_any_dtype(tabela[col]) else "texto"
    )
    print()
    print(f"{UI_INDENTACAO_1}Resumo estatístico de '{col}' (coluna {tipo_col}):\n")

    for stat, val in desc.items():
        label = ETIQUETAS_ESTAT.get(stat, stat)
        if pd.isna(val):
            valor_str = "--"
        elif hasattr(val, "strftime"):
            valor_str = val.strftime("%d/%m/%Y")
        elif isinstance(val, float):
            valor_str = formatar_numero(val)
        else:
            valor_str = str(val)
        print(UI_INDENTACAO_2 + label.ljust(22) + " " + valor_str)
    pausar()


def estat_base():
    col = pedir_coluna_numerica()
    if col is None:
        return
    print()
    print(f"{UI_INDENTACAO_1}Estatísticas base de '{col}':")
    print(f"{UI_INDENTACAO_2}Média:   {formatar_numero(tabela[col].mean())}")
    print(f"{UI_INDENTACAO_2}Soma:    {formatar_numero(tabela[col].sum())}")
    print(f"{UI_INDENTACAO_2}Mínimo:  {formatar_numero(tabela[col].min())}")
    print(f"{UI_INDENTACAO_2}Máximo:  {formatar_numero(tabela[col].max())}")
    print(f"{UI_INDENTACAO_2}Mediana: {formatar_numero(tabela[col].median())}")
    pausar()


def estat_desv_var():
    col = pedir_coluna_numerica()
    if col is None:
        return
    print()
    print(f"{UI_INDENTACAO_1}Desvio padrão de '{col}': {formatar_numero(tabela[col].std())}")
    print(f"{UI_INDENTACAO_1}Variância de '{col}':    {formatar_numero(tabela[col].var())}")
    pausar()


def estat_moda():
    col = pedir_coluna()
    if col is None:
        return
    modas = tabela[col].mode()
    print()
    if modas.empty:
        print(f"{UI_INDENTACAO_1}Sem moda calculável.")
    else:
        valores = ", ".join(str(m) for m in modas.tolist())
        print(f"{UI_INDENTACAO_1}Moda de '{col}': {valores}")
    pausar()


def estat_quantis():
    col = pedir_coluna_numerica()
    if col is None:
        return
    pers_input = input(f"{UI_INDENTACAO_1}Percentis separados por vírgula (ex: 0.25,0.5,0.75) [Enter = quartis]: ").strip()
    if not pers_input:
        percentis = [0.25, 0.5, 0.75]
    else:
        try:
            percentis = [float(p.strip()) for p in pers_input.split(",")]
            if any(p < 0 or p > 1 for p in percentis):
                raise ValueError
        except ValueError:
            print(f"{UI_INDENTACAO_1}Percentis inválidos (devem ser entre 0 e 1). A usar quartis.")
            percentis = [0.25, 0.5, 0.75]

    quants = tabela[col].quantile(percentis)
    print()
    print(f"{UI_INDENTACAO_1}Percentis de '{col}':")
    for p, v in quants.items():
        print(UI_INDENTACAO_2 + f"{p * 100:.1f}%".rjust(6) + ": " + formatar_numero(v))
    pausar()


def estat_contagem():
    col = pedir_coluna()
    if col is None:
        return
    n_input = input(f"{UI_INDENTACAO_1}Top N (Enter = 10): ").strip()
    try:
        n = int(n_input) if n_input else 10
        if n <= 0:
            raise ValueError
    except ValueError:
        n = 10

    contagens = tabela[col].value_counts().head(n)
    print()
    print(f"{UI_INDENTACAO_1}Top {n} valores em '{col}':\n")
    if contagens.empty:
        print(f"{UI_INDENTACAO_2}(vazio)")
    else:
        imprimir_tabela_pandas(contagens)
    pausar()


def estat_unicos():
    col = pedir_coluna_ou_todas()
    if col is None:
        return
    if col is _TODAS:
        print()
        print(f"{UI_INDENTACAO_1}Valores únicos por coluna:\n")
        for c in tabela.columns:
            print(UI_INDENTACAO_2 + str(c).ljust(22) + " " + f"{tabela[c].nunique()} (de {int(tabela[c].notna().sum())} não-nulos)")
        pausar()
        return
    n_unicos = tabela[col].nunique()
    n_total = int(tabela[col].notna().sum())
    print()
    print(f"{UI_INDENTACAO_1}Valores únicos em '{col}': {n_unicos} (de {n_total} não-nulos)")
    pausar()


def estat_asim_kurt():
    col = pedir_coluna_numerica()
    if col is None:
        return
    print()
    print(f"{UI_INDENTACAO_1}Assimetria (skew) de '{col}': {formatar_numero(tabela[col].skew())}")
    print(f"{UI_INDENTACAO_1}Curtose (kurt) de '{col}':    {formatar_numero(tabela[col].kurt())}")
    pausar()



def estatisticas():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    while True:
        cabecalho("ESTATÍSTICAS")
        print(f"{UI_INDENTACAO_1}[1] Describe (resumo completo)")
        print(f"{UI_INDENTACAO_1}[2] Estatísticas base (média/soma/mín/máx/mediana)")
        print(f"{UI_INDENTACAO_1}[3] Desvio padrão + variância")
        print(f"{UI_INDENTACAO_1}[4] Moda")
        print(f"{UI_INDENTACAO_1}[5] Quartis / Percentis")
        print(f"{UI_INDENTACAO_1}[6] Top N valores mais frequentes")
        print(f"{UI_INDENTACAO_1}[7] Nº de valores únicos")
        print(f"{UI_INDENTACAO_1}[8] Assimetria + Curtose")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1","2","3","4","5","6","7","8","0"})

        if escolha == "0":
            break
        elif escolha == "1":
            estat_resumo()
        elif escolha == "2":
            estat_base()
        elif escolha == "3":
            estat_desv_var()
        elif escolha == "4":
            estat_moda()
        elif escolha == "5":
            estat_quantis()
        elif escolha == "6":
            estat_contagem()
        elif escolha == "7":
            estat_unicos()
        elif escolha == "8":
            estat_asim_kurt()


# Análise avançada: info do DataFrame, groupby, duplicados, missing data
def info_resumo():
    if tabela.columns.empty:
        print(f"{UI_INDENTACAO_1}Sem folha.")
        pausar()
        return

    cabecalho("INFO / RESUMO")
    print()
    print(f"{UI_INDENTACAO_1}Dimensões: {len(tabela)} linhas × {len(tabela.columns)} colunas")
    print(f"{UI_INDENTACAO_1}Memória:   {tabela.memory_usage(deep=True).sum() / 1024:.2f} KB")
    print()
    print(UI_INDENTACAO_1 + "Coluna".ljust(22) + " " + "Tipo".ljust(14) + " " + "Não-nulos".rjust(12) + " " + "Missing %".rjust(11))
    print(UI_INDENTACAO_1 + "-" * 62)
    n_total = len(tabela)
    for col in tabela.columns:
        n_nao_null = int(tabela[col].notna().sum())
        dtype_str = str(tabela[col].dtype)
        miss_pct = (1 - n_nao_null / n_total) * 100 if n_total > 0 else 0.0
        print(UI_INDENTACAO_1 + str(col).ljust(22) + " " + dtype_str.ljust(14) + " " + str(n_nao_null).rjust(12) + " " + f"{miss_pct:.1f}%".rjust(11))
    pausar()


def agrupar():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    cabecalho("AGRUPAR (groupby)")
    print()

    col_grupo = pedir_coluna("Coluna para agrupar")
    if col_grupo is None:
        return

    col_num = pedir_coluna_numerica("Coluna numérica a agregar")
    if col_num is None:
        return

    print()
    print(f"{UI_INDENTACAO_1}Agregação: [1] Soma  [2] Média  [3] Contagem  [4] Mínimo  [5] Máximo")
    op = input(f"{UI_INDENTACAO_1}Opção: ").strip()
    agg_map = {"1": "sum", "2": "mean", "3": "count", "4": "min", "5": "max"}
    agg = agg_map.get(op, "mean")

    resultado = tabela.groupby(col_grupo)[col_num].agg(agg).reset_index()
    resultado.columns = [col_grupo, f"{col_num} ({agg})"]
    print()
    print(f"{UI_INDENTACAO_1}groupby('{col_grupo}')['{col_num}'].{agg}():\n")
    desenhar_tabela_dados(resultado)
    pausar()


def duplicados():
    global tabela
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    cabecalho("DUPLICADOS")
    print()

    dup_mask = tabela.duplicated(keep=False)
    n_dup = int(dup_mask.sum())

    if n_dup == 0:
        print(f"{UI_INDENTACAO_1}Nenhum duplicado encontrado.")
        pausar()
        return

    print(f"{UI_INDENTACAO_1}{n_dup} linha(s) duplicada(s):\n")
    desenhar_tabela_dados(tabela[dup_mask])
    print()

    if confirmar("Remover duplicados (mantém primeira ocorrência)?"):
        n_antes = len(tabela)
        tabela = tabela.drop_duplicates(ignore_index=True)
        aplicar_tipos()
        guardar_estado()
        registar_ponto()
        print(f"\n{UI_INDENTACAO_1}{n_antes - len(tabela)} duplicado(s) removido(s).")
    else:
        print(f"\n{UI_INDENTACAO_1}Nenhuma alteração.")
    pausar()


def valor_para_coluna(col, valor):
    """Ajusta um valor calculado (média, mediana) ao tipo da coluna.

    Colunas inteiras não aceitam fracções, por isso o valor é arredondado.
    """
    if pd.isna(valor):
        return valor
    if pd.api.types.is_integer_dtype(tabela[col]):
        return round(valor)
    return valor


def dados_em_falta():
    global tabela
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    cabecalho("DADOS EM FALTA")
    print()

    contagens = tabela.isna().sum()
    n_total = len(tabela)

    print(f"{UI_INDENTACAO_1}Missing por coluna (total {n_total} linhas):\n")
    for col in tabela.columns:
        n_miss = int(contagens[col])
        pct = (n_miss / n_total * 100) if n_total else 0
        print(UI_INDENTACAO_2 + str(col).ljust(22) + " " + str(n_miss).rjust(5) + f"{UI_INDENTACAO_1} (" + f"{pct:.1f}%".rjust(6) + ")")
    print()

    if int(contagens.sum()) == 0:
        print(f"{UI_INDENTACAO_1}Sem missing data. Nada a tratar.")
        pausar()
        return

    print(f"{UI_INDENTACAO_1}Acções:")
    print(f"{UI_INDENTACAO_2}[1] Preencher missing (fillna)")
    print(f"{UI_INDENTACAO_2}[2] Remover linhas com missing (dropna)")
    print(f"{UI_INDENTACAO_2}[0] Voltar")
    op = input(f"{UI_INDENTACAO_1}Opção: ").strip()

    if op == "1":
        col = pedir_coluna("Coluna a preencher")
        if col is None:
            return
        print()
        print(f"{UI_INDENTACAO_2}[1] Valor específico")
        print(f"{UI_INDENTACAO_2}[2] Média (se numérica)")
        print(f"{UI_INDENTACAO_2}[3] Mediana (se numérica)")
        print(f"{UI_INDENTACAO_2}[4] Forward-fill (ffill)")
        sub_op = input(f"{UI_INDENTACAO_1}Opção: ").strip()

        if sub_op == "1":
            valor = input(f"{UI_INDENTACAO_1}Valor a usar: ").strip()
            tipo = tipo_de(col)
            try:
                tabela[col] = tabela[col].fillna(converter(valor, tipo))
            except Exception as e:
                print(f"{UI_INDENTACAO_1}Erro: {e}")
                pausar()
                return
        elif sub_op == "2":
            if not pd.api.types.is_numeric_dtype(tabela[col]):
                print(f"{UI_INDENTACAO_1}Coluna não numérica.")
                pausar()
                return
            tabela[col] = tabela[col].fillna(valor_para_coluna(col, tabela[col].mean()))
        elif sub_op == "3":
            if not pd.api.types.is_numeric_dtype(tabela[col]):
                print(f"{UI_INDENTACAO_1}Coluna não numérica.")
                pausar()
                return
            tabela[col] = tabela[col].fillna(valor_para_coluna(col, tabela[col].median()))
        elif sub_op == "4":
            tabela[col] = tabela[col].ffill()
        else:
            return

        aplicar_tipos()
        guardar_estado()
        registar_ponto()
        print(f"\n{UI_INDENTACAO_1}Coluna '{col}' preenchida.")
    elif op == "2":
        if not confirmar(f"Remover linhas com qualquer missing ({int(contagens.sum())} células)?"):
            return
        n_antes = len(tabela)
        tabela = tabela.dropna().reset_index(drop=True)
        guardar_estado()
        registar_ponto()
        print(f"\n{UI_INDENTACAO_1}{n_antes - len(tabela)} linha(s) removida(s).")
    pausar()


# Gráficos (matplotlib + numpy): dispersão, área, histograma, boxplot, heatmap, tendência
def gerar_grafico_dispersao():
    cols = pedir_colunas_numericas(
        min_n=2,
        max_n=2,
        mensagem="Colunas X,Y para dispersão (ex.: idade,salario)"
    )
    if cols is None:
        return
    col_x, col_y = cols
    pares = tabela[[col_x, col_y]].dropna()
    if pares.empty:
        print(f"{UI_INDENTACAO_1}Sem pares de valores válidos para dispersão.")
        pausar()
        return

    x = pares[col_x].to_numpy(dtype=float)
    y = pares[col_y].to_numpy(dtype=float)
    plt.figure()
    plt.scatter(x, y, color="purple", alpha=0.7, edgecolor="black")
    plt.xlabel(col_x)
    plt.ylabel(col_y)
    plt.title(f"Dispersão: {col_x} vs {col_y}")
    plt.grid(linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()


def gerar_grafico_area():
    cols = pedir_colunas_numericas(min_n=1, mensagem="Coluna(s) para área")
    if cols is None:
        return
    x = np.arange(len(tabela))
    plt.figure()
    for c in cols:
        vals = tabela[c].fillna(0).to_numpy(dtype=float)
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
def gerar_histograma():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return
    resultado = pedir_colunas_numericas(min_n=1, max_n=1, mensagem="Coluna para histograma")
    if resultado is None:
        return
    col = resultado[0]
    vals = tabela[col].dropna().to_numpy(dtype=float)
    if len(vals) == 0:
        print(f"{UI_INDENTACAO_1}Sem valores numéricos.")
        pausar()
        return
    q1, mediana, q3 = np.percentile(vals, [25, 50, 75])
    bins_auto = max(5, int(np.sqrt(len(vals))))
    bins_max = max(10, min(100, len(vals)))
    bins = pedir_inteiro_intervalo(f"{UI_INDENTACAO_1}Nº de bins [{2}-{bins_max}] (Enter = {bins_auto}): ", 2, bins_max, default=bins_auto)
    if bins is None:
        return
    plt.figure()
    plt.hist(vals, bins=bins, color="teal", edgecolor="black", alpha=0.75)
    plt.axvline(q1, color="orange", linestyle="--", label=f"Q1 ({formatar_numero(q1)})")
    plt.axvline(mediana, color="red", linestyle="--", label=f"Mediana ({formatar_numero(mediana)})")
    plt.axvline(q3, color="purple", linestyle="--", label=f"Q3 ({formatar_numero(q3)})")
    plt.xlabel(col)
    plt.ylabel("Frequência")
    plt.title(f"Distribuição de '{col}'")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()


def gerar_diagrama_caixas():
    cols = pedir_colunas_numericas(min_n=1, mensagem="Coluna(s) para boxplot")
    if cols is None:
        return
    dados = [tabela[c].dropna().to_numpy(dtype=float) for c in cols]
    plt.figure()
    plt.boxplot(dados, tick_labels=cols)
    for i, arr in enumerate(dados, start=1):
        if len(arr) == 0:
            continue
        q1, _, q3 = np.percentile(arr, [25, 50, 75])
        plt.text(i, q3, f"Q3={formatar_numero(q3)}", ha="center", va="bottom", fontsize=8)
        plt.text(i, q1, f"Q1={formatar_numero(q1)}", ha="center", va="top", fontsize=8)
    plt.xlabel("Coluna")
    plt.ylabel("Valor")
    plt.title("Boxplot — quartis e mediana por coluna")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()


def gerar_mapa_calor_correlacao():
    cols = pedir_colunas_numericas(min_n=2, mensagem="Colunas para correlação")
    if cols is None:
        return
    dados = tabela[cols].dropna()
    if len(dados) < 2:
        print(f"{UI_INDENTACAO_1}Necessário pelo menos 2 linhas com dados válidos nas colunas selecionadas.")
        pausar()
        return
    matriz = dados.corr().to_numpy()
    if np.isnan(matriz).all():
        print(f"{UI_INDENTACAO_1}Correlação indefinida para as colunas selecionadas.")
        pausar()
        return
    n = len(cols)
    plt.figure()
    plt.imshow(matriz, cmap="coolwarm", vmin=-1, vmax=1)
    plt.colorbar(label="Correlação")
    posicoes = np.arange(n)
    plt.xticks(posicoes, cols, rotation=45, ha="right")
    plt.yticks(posicoes, cols)
    for i in range(n):
        for j in range(n):
            plt.text(j, i, f"{matriz[i, j]:.2f}", ha="center", va="center", color="black")
    plt.title("Heatmap de correlação")
    plt.tight_layout()
    plt.show()


def gerar_linhas_media_movel():
    cols = pedir_colunas_numericas(min_n=1, mensagem="Coluna(s) para média móvel")
    if cols is None:
        return
    n_linhas = len(tabela)
    if n_linhas < 2:
        print(f"{UI_INDENTACAO_1}Necessário pelo menos 2 linhas.")
        pausar()
        return
    janela_max = min(50, n_linhas)
    janela_default = min(5, janela_max)
    janela = pedir_inteiro_intervalo(f"{UI_INDENTACAO_1}Janela da média móvel [2-{janela_max}] (Enter = {janela_default}): ", 2, janela_max, default=janela_default)
    if janela is None:
        return
    kernel = np.ones(janela) / janela
    x = np.arange(n_linhas)
    plt.figure()
    for c in cols:
        vals = tabela[c].to_numpy(dtype=float)
        if np.isnan(vals).all():
            print(f"{UI_INDENTACAO_1}Coluna '{c}' sem valores válidos — ignorada.")
            continue
        if np.isnan(vals).any():
            vals = np.where(np.isnan(vals), np.nanmean(vals), vals)
        suavizado = np.convolve(vals, kernel, mode="same")
        plt.plot(x, suavizado, marker="o", label=c)
    plt.xlabel("Índice da linha")
    plt.ylabel("Valor suavizado")
    plt.title(f"Média móvel (np.convolve, janela={janela})")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()


def gerar_tendencia_polinomial():
    cols = pedir_colunas_numericas(min_n=1, mensagem="Coluna(s) para tendência")
    if cols is None:
        return
    n = len(tabela)
    if n < 2:
        print(f"{UI_INDENTACAO_1}Necessário pelo menos 2 linhas.")
        pausar()
        return
    grau_max = min(5, n - 1)
    grau_default = min(2, grau_max)
    grau = pedir_inteiro_intervalo(f"{UI_INDENTACAO_1}Grau do polinómio [1-{grau_max}] (Enter = {grau_default}): ", 1, grau_max, default=grau_default)
    if grau is None:
        return
    x = np.arange(n)
    plt.figure()
    for c in cols:
        y = tabela[c].to_numpy(dtype=float)
        if np.isnan(y).all():
            print(f"{UI_INDENTACAO_1}Coluna '{c}' sem valores válidos — ignorada.")
            continue
        if np.isnan(y).any():
            y = np.where(np.isnan(y), np.nanmean(y), y)
        coefs = np.polyfit(x, y, grau)
        poly = np.poly1d(coefs)
        tendencia = poly(x)
        plt.scatter(x, y, alpha=0.3, s=20)
        plt.plot(x, tendencia, linewidth=2, label=c)
    plt.xlabel("Índice da linha")
    plt.ylabel("Valor")
    plt.title("Tendência polinomial (np.polyfit)")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()


# Gráficos com eixos personalizados (escolha de eixo X e agregações)
def gerar_barras_eixos():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    print()
    print(f"{UI_INDENTACAO_1}Eixo X: [1] Índice da linha  [2] Coluna da tabela")
    op_x = input(f"{UI_INDENTACAO_1}Opção: ").strip()

    if op_x == "2":
        col_x = pedir_coluna("Coluna para eixo X")
        if col_x is None:
            return
        col_y = pedir_coluna_numerica("Coluna numérica para eixo Y")
        if col_y is None:
            return
        print()
        print(f"{UI_INDENTACAO_1}Agregação: [1] Soma  [2] Média  [3] Contagem  [4] Mínimo  [5] Máximo")
        agg_map = {"1": "sum", "2": "mean", "3": "count", "4": "min", "5": "max"}
        agg = agg_map.get(input(f"{UI_INDENTACAO_1}Opção: ").strip(), "mean")
        resultado = tabela.groupby(col_x)[col_y].agg(agg)
        x_vals = [str(v) for v in resultado.index]
        y_vals = resultado.values
        xlabel = str(col_x)
        ylabel = f"{col_y} ({agg})"
        titulo = f"Barras: {col_y} por {col_x} ({agg})"
    else:
        col_y = pedir_coluna_numerica("Coluna numérica para eixo Y")
        if col_y is None:
            return
        x_vals = list(tabela.index.astype(str))
        y_vals = tabela[col_y].to_numpy(dtype=float)
        xlabel = "Índice"
        ylabel = str(col_y)
        titulo = f"Barras: {col_y} por índice"

    plt.figure()
    plt.bar(x_vals, y_vals, color="steelblue", edgecolor="black")
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.title(titulo)
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.show()


def gerar_linhas_eixos():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    print()
    print(f"{UI_INDENTACAO_1}Eixo X: [1] Índice da linha  [2] Coluna da tabela")
    op_x = input(f"{UI_INDENTACAO_1}Opção: ").strip()

    if op_x == "2":
        col_x = pedir_coluna("Coluna para eixo X")
        if col_x is None:
            return
        x_vals = tabela[col_x].tolist()
        xlabel = str(col_x)
    else:
        x_vals = list(tabela.index)
        xlabel = "Índice"

    cols_y = pedir_colunas_numericas(min_n=1, mensagem="Coluna(s) para eixo Y")
    if cols_y is None:
        return

    plt.figure()
    for c in cols_y:
        y_vals = tabela[c].to_numpy(dtype=float)
        plt.plot(x_vals, y_vals, marker="o", label=c)
    plt.xlabel(xlabel)
    plt.ylabel("Valor")
    plt.title("Linhas — eixo X personalizado")
    plt.grid(linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.show()


def gerar_circular_eixos():
    if tabela.empty:
        print(f"{UI_INDENTACAO_1}Sem dados.")
        pausar()
        return

    col_labels = pedir_coluna("Coluna de categorias (labels)")
    if col_labels is None:
        return
    col_val = pedir_coluna_numerica("Coluna numérica (valores)")
    if col_val is None:
        return

    print()
    print(f"{UI_INDENTACAO_1}Agregação: [1] Soma  [2] Média  [3] Contagem")
    agg_map = {"1": "sum", "2": "mean", "3": "count"}
    agg = agg_map.get(input(f"{UI_INDENTACAO_1}Opção: ").strip(), "sum")

    resultado = tabela.groupby(col_labels)[col_val].agg(agg)
    resultado = resultado[resultado > 0]
    if resultado.empty:
        print(f"{UI_INDENTACAO_1}Sem valores positivos para desenhar.")
        pausar()
        return

    plt.figure()
    plt.pie(resultado.values, labels=[str(rotulo) for rotulo in resultado.index], autopct="%1.1f%%", startangle=90)
    plt.title(f"Circular: {col_val} por {col_labels} ({agg})")
    plt.tight_layout()
    plt.show()


# Menu de gráficos: orquestra todas as visualizações disponíveis
def menu_graficos():
    while True:
        cabecalho("GRÁFICOS (matplotlib + numpy)")
        print(f"{UI_INDENTACAO_1}[1] Barras (eixos à escolha)")
        print(f"{UI_INDENTACAO_1}[2] Linhas (eixo X à escolha)")
        print(f"{UI_INDENTACAO_1}[3] Circular (categorias reais)")
        print(f"{UI_INDENTACAO_1}[4] Dispersão (2 colunas)")
        print(f"{UI_INDENTACAO_1}[5] Área (soma acumulada)")
        print(f"{UI_INDENTACAO_1}[6] Histograma (1 coluna + quartis)")
        print(f"{UI_INDENTACAO_1}[7] Boxplot (quartis e mediana)")
        print(f"{UI_INDENTACAO_1}[8] Heatmap de correlação")
        print(f"{UI_INDENTACAO_1}[9] Linhas com média móvel")
        print(f"{UI_INDENTACAO_1}[A] Tendência polinomial")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            gerar_barras_eixos()
        elif escolha == "2":
            gerar_linhas_eixos()
        elif escolha == "3":
            gerar_circular_eixos()
        elif escolha == "4":
            gerar_grafico_dispersao()
        elif escolha == "5":
            gerar_grafico_area()
        elif escolha == "6":
            gerar_histograma()
        elif escolha == "7":
            gerar_diagrama_caixas()
        elif escolha == "8":
            gerar_mapa_calor_correlacao()
        elif escolha == "9":
            gerar_linhas_media_movel()
        elif escolha == "A":
            gerar_tendencia_polinomial()


# Menus secundários: folha, linhas, colunas, análise básica e avançada
def menu_folha():
    while True:
        cabecalho("GESTÃO DE FOLHA")
        print(f"{UI_INDENTACAO_1}[1] Nova folha (definir colunas)")
        print(f"{UI_INDENTACAO_1}[2] Importar Excel")
        print(f"{UI_INDENTACAO_1}[3] Exportar Excel")
        print(f"{UI_INDENTACAO_1}[4] Exportar Word (tabela + estatísticas)")
        print(f"{UI_INDENTACAO_1}[5] Exportar PDF  (tabela + estatísticas)")
        print(f"{UI_INDENTACAO_1}[6] Folha dummy (dados de exemplo)")
        print(f"{UI_INDENTACAO_1}[7] Limpar folha (reset)")
        print(f"{UI_INDENTACAO_1}[8] Importar CSV")
        print(f"{UI_INDENTACAO_1}[9] Exportar CSV")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "4", "5", "6", "7", "8", "9", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            nova_folha()
        elif escolha == "2":
            importar()
        elif escolha == "3":
            exportar()
        elif escolha == "4":
            exportar_word()
        elif escolha == "5":
            exportar_pdf()
        elif escolha == "6":
            criar_folha_dummy()
        elif escolha == "7":
            limpar_folha()
        elif escolha == "8":
            importar_csv()
        elif escolha == "9":
            exportar_csv()


def menu_linhas():
    while True:
        cabecalho("LINHAS")
        print(f"{UI_INDENTACAO_1}[1] Adicionar linha(s)")
        print(f"{UI_INDENTACAO_1}[2] Editar célula")
        print(f"{UI_INDENTACAO_1}[3] Remover linha")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            adicionar_linha()
        elif escolha == "2":
            editar_celula()
        elif escolha == "3":
            remover_linha()


def menu_colunas():
    while True:
        cabecalho("COLUNAS")
        print(f"{UI_INDENTACAO_1}[1] Adicionar coluna")
        print(f"{UI_INDENTACAO_1}[2] Remover coluna")
        print(f"{UI_INDENTACAO_1}[3] Renomear coluna")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            adicionar_coluna()
        elif escolha == "2":
            remover_coluna()
        elif escolha == "3":
            renomear_coluna()


def menu_analise():
    while True:
        cabecalho("ANÁLISE BÁSICA")
        print(f"{UI_INDENTACAO_1}[1] Ordenar (uma ou múltiplas colunas)")
        print(f"{UI_INDENTACAO_1}[2] Filtrar / Pesquisar")
        print(f"{UI_INDENTACAO_1}[3] Ver Início / Fim")
        print(f"{UI_INDENTACAO_1}[4] Amostra aleatória")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "4", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            ordenar()
        elif escolha == "2":
            filtrar()
        elif escolha == "3":
            ver_inicio_fim()
        elif escolha == "4":
            amostra()


def menu_avancado():
    while True:
        cabecalho("ANÁLISE AVANÇADA")
        print(f"{UI_INDENTACAO_1}[1] Info / Resumo do DataFrame")
        print(f"{UI_INDENTACAO_1}[2] Agrupar (groupby + agregação)")
        print(f"{UI_INDENTACAO_1}[3] Duplicados (detectar e remover)")
        print(f"{UI_INDENTACAO_1}[4] Dados em falta (contar / preencher / remover)")
        print(f"{UI_INDENTACAO_1}[0] Voltar")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "4", "0"})
        if escolha == "0":
            break
        elif escolha == "1":
            info_resumo()
        elif escolha == "2":
            agrupar()
        elif escolha == "3":
            duplicados()
        elif escolha == "4":
            dados_em_falta()


# Ecrã "Sobre" com informação da aplicação
def sobre():
    cabecalho("SOBRE")
    print(f"{UI_INDENTACAO_1}{NOME_APP}  —  Folha de Cálculo em Linha de Comando")
    print()
    print(f"{UI_INDENTACAO_1}Versão   : {VERSAO_APP}")
    print(f"{UI_INDENTACAO_1}Autor    : {AUTOR_APP}")
    print(f"{UI_INDENTACAO_1}Licença  : MIT License")
    print()
    print(f"{UI_INDENTACAO_1}Funcionalidades:")
    print(f"{UI_INDENTACAO_2}- Criar, editar, filtrar e analisar dados")
    print(f"{UI_INDENTACAO_2}- Tipos: texto, inteiro, decimal, data")
    print(f"{UI_INDENTACAO_2}- Estatísticas e análise avançada (groupby, duplicados)")
    print(f"{UI_INDENTACAO_2}- Gráficos: barras, linhas, circular, dispersão, área,")
    print(f"{UI_INDENTACAO_2}  histograma, boxplot, heatmap, tendência polinomial")
    print(f"{UI_INDENTACAO_2}- Importar / Exportar Excel e CSV; exportar Word e PDF")
    print(f"{UI_INDENTACAO_2}- Undo / Redo e gravação automática de estado")
    print()
    print(f"{UI_INDENTACAO_1}Requisitos: Python 3.x · pandas · numpy · matplotlib · python-docx · fpdf2")
    print("-" * LARGURA)
    pausar()


# Ponto de entrada: menu principal e ciclo da aplicação
def configurar_log():
    """Regista erros em .termifolha.log (rotativo). Chamado só em main(): importar o módulo não cria ficheiros."""
    from logging.handlers import RotatingFileHandler
    try:
        handler = RotatingFileHandler(FICHEIRO_LOG, maxBytes=256 * 1024, backupCount=1, encoding="utf-8")
    except OSError:
        return  # sem permissões de escrita: a aplicação funciona na mesma, sem registo
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(handler)
    log.setLevel(logging.INFO)


def main():
    global tabela, tipos

    configurar_log()
    carregar_estado()

    if not undo:
        undo.append((tabela.copy(), tipos.copy()))

    while True:
        cabecalho("MENU PRINCIPAL")
        print(f"{UI_INDENTACAO_1}[1] Folha            (nova, importar, exportar, dummy, limpar)")
        print(f"{UI_INDENTACAO_1}[2] Linhas           (adicionar, editar, remover)")
        print(f"{UI_INDENTACAO_1}[3] Colunas          (adicionar, remover, renomear)")
        print(f"{UI_INDENTACAO_1}[4] Análise Básica   (ordenar, filtrar, início/fim, amostra)")
        print(f"{UI_INDENTACAO_1}[5] Estatísticas     (8 métodos pandas)")
        print(f"{UI_INDENTACAO_1}[6] Análise Avançada (info, groupby, duplicados, missing)")
        print(f"{UI_INDENTACAO_1}[7] Gráficos         (matplotlib + numpy)")
        print(f"{UI_INDENTACAO_1}[S] Sobre")
        print(f"{UI_INDENTACAO_1}[U] Desfazer  [{len(undo) - 1} passo(s)]")
        print(f"{UI_INDENTACAO_1}[R] Refazer   [{len(redo)} passo(s)]")
        print(f"{UI_INDENTACAO_1}[Q] Sair")
        print("-" * LARGURA)

        escolha = pedir_opcao({"1", "2", "3", "4", "5", "6", "7", "S", "U", "R", "Q"})

        if escolha == "1":
            menu_folha()
        elif escolha == "2":
            menu_linhas()
        elif escolha == "3":
            menu_colunas()
        elif escolha == "4":
            menu_analise()
        elif escolha == "5":
            estatisticas()
        elif escolha == "6":
            menu_avancado()
        elif escolha == "7":
            menu_graficos()
        elif escolha == "S":
            sobre()
        elif escolha == "U":
            desfazer()
        elif escolha == "R":
            refazer()
        elif escolha == "Q":
            limpar()
            print("Obrigado por usar o TermiFolha. Até à próxima!")
            break


if __name__ == "__main__":
    main()
