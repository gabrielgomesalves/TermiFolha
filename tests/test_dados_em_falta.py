import pytest


def test_preencher_com_media_em_coluna_inteira(app):
    """Regressão A1: a média (31.57) não cabe em Int64 e fazia rebentar o programa."""
    app.dummy()
    assert int(app.m.tabela["Idade"].isna().sum()) == 1

    app.run(app.m.dados_em_falta, "1", "idade", "2")

    assert int(app.m.tabela["Idade"].isna().sum()) == 0
    assert str(app.m.tabela["Idade"].dtype) == "Int64"
    assert app.m.tabela.loc[3, "Idade"] == 32  # média 31.57 arredondada


def test_preencher_com_mediana_fraccionaria_em_coluna_inteira(app):
    """Regressão A1 (mediana): mediana de [1, 2, em falta] é 1.5."""
    app.m.tabela = app.m.pd.DataFrame({"N": [1, 2, None]})
    app.m.tipos = {"N": int}
    app.m.aplicar_tipos()

    app.run(app.m.dados_em_falta, "1", "n", "3")

    assert int(app.m.tabela["N"].isna().sum()) == 0
    assert str(app.m.tabela["N"].dtype) == "Int64"
    assert app.m.tabela.loc[2, "N"] == 2  # 1.5 arredondado


def test_preencher_com_media_em_coluna_decimal_mantem_valor_exacto(app):
    app.dummy()
    app.m.tabela.loc[0, "Salario"] = None
    esperado = app.m.tabela["Salario"].mean()

    app.run(app.m.dados_em_falta, "1", "salario", "2")

    assert app.m.tabela.loc[0, "Salario"] == pytest.approx(esperado)
