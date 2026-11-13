"""Livro de folhas (separadores), formato de estado v2 e histórico configurável."""
import json
from pathlib import Path

import pandas as pd
import pytest


def _duas_folhas(app):
    app.dummy()
    app.m.add_sheet("Segunda", pd.DataFrame({"X": [1, 2]}), {"X": int})
    app.m.switch_sheet("Folha1")


def test_nome_de_folha_invalido_e_repetido(app):
    erro = app.m.sheet_name_error
    assert erro("") and erro("  ") and erro(" a") and erro("a" * 32) and erro("a/b") and erro("a:b")
    assert erro("Folha1")  # já existe
    assert erro("folha1")  # sem distinguir maiúsculas
    assert erro("Vendas") is None


def test_nome_unico_acrescenta_numero(app):
    assert app.m.unique_sheet_name("Folha1") == "Folha1 (2)"
    assert app.m.unique_sheet_name("a/b") == "a_b"
    assert len(app.m.unique_sheet_name("x" * 40)) <= 31


def test_adicionar_e_trocar_de_folha_preserva_os_dados(app):
    _duas_folhas(app)

    assert app.m.sheet_names() == ["Folha1", "Segunda"]
    assert len(app.m.table) == 8

    app.m.switch_sheet("Segunda")
    assert app.m.active_name == "Segunda" and app.m.table["X"].tolist() == [1, 2]

    app.m.switch_sheet("Folha1")
    assert len(app.m.table) == 8 and list(app.m.col_types) == ["Nome", "Idade", "Salario", "Cidade", "Admissao"]


def test_adicionar_folha_com_nome_invalido_levanta(app):
    with pytest.raises(ValueError, match="já existe"):
        app.m.add_sheet("Folha1")


def test_renomear_folha_mantem_posicao_e_activa(app):
    _duas_folhas(app)

    app.m.rename_sheet("Folha1", "Pessoas")

    assert app.m.sheet_names() == ["Pessoas", "Segunda"] and app.m.active_name == "Pessoas"
    with pytest.raises(ValueError):
        app.m.rename_sheet("Pessoas", "segunda")


def test_eliminar_folha_activa_passa_para_a_vizinha(app):
    _duas_folhas(app)

    app.m.delete_sheet("Folha1")

    assert app.m.sheet_names() == ["Segunda"] and app.m.active_name == "Segunda"
    with pytest.raises(ValueError, match="pelo menos uma"):
        app.m.delete_sheet("Segunda")


def test_duplicar_folha_copia_dados_e_metadados_sem_partilhar_estado(app):
    app.dummy()
    app.m.meta["formats"]["Salario"] = {"decimals": 2}

    app.m.duplicate_sheet("Folha1", "Copia")
    app.m.table.loc[0, "Nome"] = "ALTERADO"
    app.m.meta["formats"]["Salario"]["decimals"] = 0
    app.m.switch_sheet("Folha1")

    assert app.m.table.loc[0, "Nome"] == "João"
    assert app.m.meta["formats"]["Salario"] == {"decimals": 2}


def test_desfazer_repoe_a_folha_activa_e_as_outras(app):
    app.dummy()
    app.m.add_sheet("Segunda", pd.DataFrame({"X": [1]}), {"X": int})
    app.m.commit()

    app.m.switch_sheet("Folha1")
    app.run(app.m.remove_row, "0", "s")
    assert len(app.m.table) == 7

    app.m.undo_last()

    assert app.m.active_name == "Segunda" and app.m.sheet_names() == ["Folha1", "Segunda"]
    app.m.switch_sheet("Folha1")
    assert len(app.m.table) == 8


def test_estado_v2_guarda_todas_as_folhas_e_metadados(app):
    _duas_folhas(app)
    app.m.meta["formats"]["Salario"] = {"decimals": 1}
    app.m.meta["rules"]["Idade"] = {"min": 0}
    app.m.meta["visible"] = ["Nome", "Idade"]
    app.m.meta["totals"] = "sum"
    app.m.commit()

    app.m.reset_workbook()
    app.m.load_state()

    assert app.m.sheet_names() == ["Folha1", "Segunda"] and app.m.active_name == "Folha1"
    assert app.m.meta["formats"] == {"Salario": {"decimals": 1}}
    assert app.m.meta["rules"] == {"Idade": {"min": 0}}
    assert app.m.meta["visible"] == ["Nome", "Idade"] and app.m.meta["totals"] == "sum"
    app.m.switch_sheet("Segunda")
    assert app.m.table["X"].tolist() == [1, 2] and str(app.m.table["X"].dtype) == "Int64"


def test_le_estado_v1_de_uma_so_folha(app):
    v1 = {"versao": 1, "colunas": ["A"], "tipos": {"A": "int"}, "dados": [[1], [None]]}
    Path(app.m.STATE_FILE).write_text(json.dumps(v1), encoding="utf-8")

    app.m.load_state()

    assert app.m.sheet_names() == ["Folha1"] and app.m.table["A"].isna().sum() == 1
    assert app.m.settings["history_limit"] == app.m.HISTORY_LIMIT


@pytest.mark.parametrize("alteracao", [
    lambda d: d.update(folhas=[]),
    lambda d: d["folhas"].append(dict(d["folhas"][0])),  # nome repetido
    lambda d: d["folhas"][0].update(regras=[1]),
    lambda d: d["folhas"][0].update(formulas={"A": 5}),
    lambda d: d["folhas"][0].update(visiveis="A"),
])
def test_estado_v2_invalido_e_preservado_como_corrompido(app, alteracao):
    estado = {"versao": 2, "activa": "Folha1", "folhas": [{"nome": "Folha1", "colunas": ["A"], "tipos": {"A": "int"}, "dados": [[1]]}]}
    alteracao(estado)
    caminho = Path(app.m.STATE_FILE)
    caminho.write_text(json.dumps(estado), encoding="utf-8")

    app.m.load_state()

    assert not caminho.exists() and Path(str(caminho) + ".corrompido").exists()
    assert app.m.table.columns.empty


def test_definicoes_invalidas_voltam_ao_valor_por_omissao(app):
    limpas = app.m._sanitize_settings({"history_limit": 0, "page_size": "x"})
    assert limpas == {"history_limit": app.m.HISTORY_LIMIT, "page_size": app.m.PAGE_SIZE}
    assert app.m._sanitize_settings({"history_limit": 7, "page_size": 50}) == {"history_limit": 7, "page_size": 50}
    assert app.m._sanitize_settings({"history_limit": True})["history_limit"] == app.m.HISTORY_LIMIT


def test_limite_do_historico_e_configuravel(app):
    app.m.settings["history_limit"] = 3
    app.dummy()
    for _ in range(6):
        app.run(app.m.remove_row, "0", "s")

    assert len(app.m.undo_stack) == 4  # estado actual + 3 passos para trás


def test_limpar_folha_com_varias_folhas_preserva_as_outras(app):
    _duas_folhas(app)

    app.run(app.m.clear_sheet, "s")

    assert app.m.table.columns.empty and Path(app.m.STATE_FILE).exists()
    app.m.switch_sheet("Segunda")
    assert len(app.m.table) == 2
