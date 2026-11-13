"""Operações sobre linhas/colunas e análise que antes não tinham testes
(e que usavam `inplace=True`, deprecado no pandas 3.1)."""
import pandas as pd


def test_remover_linha_reindexa(app):
    app.dummy()
    n = len(app.m.tabela)

    app.run(app.m.remover_linha, "0", "s")

    assert len(app.m.tabela) == n - 1
    assert list(app.m.tabela.index) == list(range(n - 1))
    assert app.m.tabela.loc[0, "Nome"] == "Maria"


def test_remover_linha_cancelada_nao_altera(app):
    app.dummy()
    n = len(app.m.tabela)

    app.run(app.m.remover_linha, "0", "n")

    assert len(app.m.tabela) == n


def test_remover_coluna_actualiza_tipos(app):
    app.dummy()

    app.run(app.m.remover_coluna, "cidade", "s")

    assert "Cidade" not in app.m.tabela.columns
    assert "Cidade" not in app.m.tipos


def test_renomear_coluna_move_o_tipo(app):
    app.dummy()

    app.run(app.m.renomear_coluna, "idade", "Anos")

    assert "Anos" in app.m.tabela.columns
    assert app.m.tipos["Anos"] is int
    assert "Idade" not in app.m.tipos


def test_ordenar_por_duas_colunas(app):
    app.dummy()

    # Cidade ascendente, depois Salario descendente
    app.run(app.m.ordenar, "cidade", "a", "salario", "d", "")

    cidades = app.m.tabela["Cidade"].tolist()
    assert cidades == sorted(cidades)
    lisboa = app.m.tabela[app.m.tabela["Cidade"] == "Lisboa"]["Salario"].tolist()
    assert lisboa == sorted(lisboa, reverse=True)
    assert list(app.m.tabela.index) == list(range(len(app.m.tabela)))


def test_dropna_remove_linhas_com_valores_em_falta(app):
    app.dummy()
    n = len(app.m.tabela)

    app.run(app.m.dados_em_falta, "2", "s")

    assert len(app.m.tabela) == n - 1
    assert not app.m.tabela.isna().any().any()


def test_duplicados_remove_mantendo_primeira_ocorrencia(app):
    app.dummy()
    n = len(app.m.tabela)

    app.run(app.m.duplicados, "s")

    assert len(app.m.tabela) == n - 1


def test_desfazer_e_refazer_repoem_estados(app):
    app.dummy()
    app.run(app.m.remover_linha, "0", "s")
    depois = len(app.m.tabela)

    app.m.desfazer()
    assert len(app.m.tabela) == depois + 1

    app.m.refazer()
    assert len(app.m.tabela) == depois


def test_historico_limitado_a_quatro_passos(app):
    app.dummy()
    for _ in range(6):
        app.run(app.m.remover_linha, "0", "s")

    assert len(app.m.undo) <= app.m.HIST_LIMITE
    assert isinstance(app.m.tabela, pd.DataFrame)
