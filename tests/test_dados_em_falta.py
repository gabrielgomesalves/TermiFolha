import pytest


def test_preencher_com_media_em_coluna_inteira(app):
    """Regressão A1: a média (31.57) não cabe em Int64 e fazia rebentar o programa."""
    app.dummy()
    assert int(app.m.table["Idade"].isna().sum()) == 1

    app.run(app.m.missing_data, "1", "idade", "2")

    assert int(app.m.table["Idade"].isna().sum()) == 0
    assert str(app.m.table["Idade"].dtype) == "Int64"
    assert app.m.table.loc[3, "Idade"] == 32  # média 31.57 arredondada


def test_preencher_com_mediana_fraccionaria_em_coluna_inteira(app):
    """Regressão A1 (mediana): mediana de [1, 2, em falta] é 1.5."""
    app.m.table = app.m.pd.DataFrame({"N": [1, 2, None]})
    app.m.col_types = {"N": int}
    app.m.apply_types()

    app.run(app.m.missing_data, "1", "n", "3")

    assert int(app.m.table["N"].isna().sum()) == 0
    assert str(app.m.table["N"].dtype) == "Int64"
    assert app.m.table.loc[2, "N"] == 2  # 1.5 arredondado


def test_preencher_com_media_em_coluna_decimal_mantem_valor_exacto(app):
    app.dummy()
    app.m.table.loc[0, "Salario"] = None
    esperado = app.m.table["Salario"].mean()

    app.run(app.m.missing_data, "1", "salario", "2")

    assert app.m.table.loc[0, "Salario"] == pytest.approx(esperado)
