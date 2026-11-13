import pandas as pd
import pytest


@pytest.mark.parametrize("tipo_menu", ["1", "2", "3", "4"])  # texto, inteiro, decimal, data
def test_adicionar_coluna_aceita_valores_em_branco(app, tipo_menu):
    """Regressão M4: o manual diz que o valor pode ficar em branco (`--`)."""
    app.dummy()
    n = len(app.m.table)

    app.run(app.m.add_column, "Extra", tipo_menu, *([""] * n))

    assert "Extra" in app.m.table.columns
    assert app.m.table["Extra"].isna().all()
    assert app.m.format_value(app.m.table.loc[0, "Extra"]) == "--"


def test_adicionar_coluna_mistura_valores_e_brancos(app):
    app.dummy()
    n = len(app.m.table)
    respostas = ["7"] + [""] * (n - 1)

    app.run(app.m.add_column, "Nivel", "2", *respostas)

    assert app.m.table.loc[0, "Nivel"] == 7
    assert int(app.m.table["Nivel"].isna().sum()) == n - 1
    assert str(app.m.table["Nivel"].dtype) == "Int64"
    assert pd.isna(app.m.table.loc[1, "Nivel"])
