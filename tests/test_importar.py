import pandas as pd
import pytest


@pytest.fixture
def xlsx_cabecalhos_numericos(tmp_path):
    path = tmp_path / "anos.xlsx"
    pd.DataFrame([[1, 2], [3, 4]], columns=[2020, 2021]).to_excel(path, index=False)
    return str(path)


def test_importar_converte_cabecalhos_numericos_para_texto(app, xlsx_cabecalhos_numericos):
    """Regressão A3: cabeçalhos inteiros faziam `', '.join(colunas)` rebentar."""
    app.run(app.m.import_excel, xlsx_cabecalhos_numericos)

    assert list(app.m.table.columns) == ["2020", "2021"]
    assert set(app.m.col_types) == {"2020", "2021"}


def test_operacoes_por_coluna_funcionam_apos_importar_cabecalhos_numericos(
    app, xlsx_cabecalhos_numericos, capsys
):
    app.run(app.m.import_excel, xlsx_cabecalhos_numericos)

    app.run(app.m.stat_mode, "2020")
    assert "Moda de '2020'" in capsys.readouterr().out

    app.run(app.m.rename_column, "2021", "ano_b")
    assert "ano_b" in app.m.table.columns
