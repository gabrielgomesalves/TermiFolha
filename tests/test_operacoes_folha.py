"""Operações sobre linhas/colunas e análise que antes não tinham testes
(e que usavam `inplace=True`, deprecado no pandas 3.1)."""
import pandas as pd


def test_remover_linha_reindexa(app):
    app.dummy()
    n = len(app.m.table)

    app.run(app.m.remove_row, "0", "s")

    assert len(app.m.table) == n - 1
    assert list(app.m.table.index) == list(range(n - 1))
    assert app.m.table.loc[0, "Nome"] == "Maria"


def test_remover_linha_cancelada_nao_altera(app):
    app.dummy()
    n = len(app.m.table)

    app.run(app.m.remove_row, "0", "n")

    assert len(app.m.table) == n


def test_remover_coluna_actualiza_tipos(app):
    app.dummy()

    app.run(app.m.remove_column, "cidade", "s")

    assert "Cidade" not in app.m.table.columns
    assert "Cidade" not in app.m.col_types


def test_renomear_coluna_move_o_tipo(app):
    app.dummy()

    app.run(app.m.rename_column, "idade", "Anos")

    assert "Anos" in app.m.table.columns
    assert app.m.col_types["Anos"] is int
    assert "Idade" not in app.m.col_types


def test_ordenar_por_duas_colunas(app):
    app.dummy()

    # Cidade ascendente, depois Salario descendente
    app.run(app.m.sort_rows, "cidade", "a", "salario", "d", "")

    cidades = app.m.table["Cidade"].tolist()
    assert cidades == sorted(cidades)
    lisboa = app.m.table[app.m.table["Cidade"] == "Lisboa"]["Salario"].tolist()
    assert lisboa == sorted(lisboa, reverse=True)
    assert list(app.m.table.index) == list(range(len(app.m.table)))


def test_dropna_remove_linhas_com_valores_em_falta(app):
    app.dummy()
    n = len(app.m.table)

    app.run(app.m.missing_data, "2", "s")

    assert len(app.m.table) == n - 1
    assert not app.m.table.isna().any().any()


def test_duplicados_remove_mantendo_primeira_ocorrencia(app):
    app.dummy()
    n = len(app.m.table)

    app.run(app.m.duplicates, "s")

    assert len(app.m.table) == n - 1


def test_desfazer_e_refazer_repoem_estados(app):
    app.dummy()
    app.run(app.m.remove_row, "0", "s")
    depois = len(app.m.table)

    app.m.undo_last()
    assert len(app.m.table) == depois + 1

    app.m.redo_last()
    assert len(app.m.table) == depois


def test_historico_limitado_a_quatro_passos(app):
    app.dummy()
    for _ in range(6):
        app.run(app.m.remove_row, "0", "s")

    assert len(app.m.undo_stack) <= app.m.HISTORY_LIMIT
    assert isinstance(app.m.table, pd.DataFrame)
