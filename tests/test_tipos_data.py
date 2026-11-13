import pandas as pd


def _nova_folha_com_data(app):
    # Colunas: Dia (data), Valor (decimal)
    app.run(app.m.nova_folha, "Dia", "4", "Valor", "3", "")
    app.run(app.m.adicionar_linha, "01/02/2024", "1.5", "05/02/2024", "2.5", "sair")


def test_coluna_data_criada_em_nova_folha_tem_dtype_datetime(app):
    """Regressão A2: ficava como `object`, e o filtro de data nunca era usado."""
    _nova_folha_com_data(app)

    assert pd.api.types.is_datetime64_any_dtype(app.m.tabela["Dia"])


def test_filtro_de_data_aplica_se_a_coluna_criada_em_nova_folha(app, capsys):
    _nova_folha_com_data(app)

    # [1] Antes de 03/02/2024 -> só a linha de 01/02/2024
    app.run(app.m.filtrar, "dia", "1", "03/02/2024")

    saida = capsys.readouterr().out
    assert "Antes de" in saida  # menu do filtro de data
    assert "1 resultado(s)" in saida


def test_edicao_de_celula_de_data_mantem_dtype_datetime(app):
    _nova_folha_com_data(app)

    app.run(app.m.editar_celula, "0", "dia", "10/03/2024")

    assert pd.api.types.is_datetime64_any_dtype(app.m.tabela["Dia"])
    assert app.m.tabela.loc[0, "Dia"] == pd.Timestamp("2024-03-10")
